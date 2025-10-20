import os
import pandas as pd
import logging
import traceback
from datetime import datetime
from django.shortcuts import render, redirect
from django.http import FileResponse, HttpResponse
from django.conf import settings
from django.contrib import messages
from .models import UploadedFile, GeneratedFile
from .utils.excel_processor import ExcelProcessor
from .utils.pdf_generator import PDFGenerator
import shutil
import zipfile

logger = logging.getLogger(__name__)

def home(request):
    # Gestion du téléchargement direct
    file_type = request.GET.get('download')
    file_path = request.GET.get('file_path')
    if file_type and file_path and os.path.exists(file_path):
        return FileResponse(open(file_path, 'rb'), as_attachment=True)
    
    if request.method == 'POST':
        prep_file = request.FILES.get('preparation_pl')
        zzz_file = request.FILES.get('zzz_file')

        if not prep_file:
            messages.error(request, "Le fichier Preparation PL est obligatoire.")
            return render(request, 'upload.html')

        # Options de génération
        generer_excel = request.POST.get('generer_excel') == 'on'
        generer_pdf = request.POST.get('generer_pdf') == 'on'
        if not generer_excel and not generer_pdf:
            messages.error(request, "Veuillez sélectionner au moins un format (Excel ou PDF).")
            return render(request, 'upload.html')

        # Champs de formulaire
        cariste = request.POST.get('cariste', 'FIRN')
        fournisseur = request.POST.get('fournisseur', 'FIRN')
        numero_dossier = request.POST.get('numero_dossier', 'DSF23044')
        type_certification = request.POST.get('type_certification', 'FSC RECYCLED 100%')
        numero_certificat = request.POST.get('numero_certificat', 'CU-COC-903458')

        try:
            # Sauvegarde fichiers uploadés
            prep_obj = UploadedFile.objects.create(
                file=prep_file,
                file_type='Préparation_PL',
                original_name=prep_file.name
            )
            zzz_obj = None
            zzz_template_path = None
            if zzz_file:
                zzz_obj = UploadedFile.objects.create(
                    file=zzz_file,
                    file_type='zzzz',
                    original_name=zzz_file.name
                )
                zzz_template_path = zzz_obj.file.path

            processor = ExcelProcessor()
            pdf_generator = PDFGenerator()

            # Lecture du fichier principal
            prep_data, _ = processor.read_excel_file(prep_obj.file.path)
            containers = processor.extract_containers(prep_data)
            if not containers:
                messages.error(request, "Aucun conteneur trouvé dans le fichier.")
                return render(request, 'upload.html')

            # 📁 Dossier commun de sortie
            base_output_dir = getattr(settings, 'CUSTOM_DOWNLOAD_DIR',
                                      os.path.join(settings.MEDIA_ROOT, 'generated'))
            os.makedirs(base_output_dir, exist_ok=True)

            # Créer un dossier session horodaté
            session_timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            session_dir = os.path.join(base_output_dir, f"session_{session_timestamp}")
            os.makedirs(session_dir, exist_ok=True)

            results = []
            successful_containers = 0
            failed_containers = []

            # 🔄 Traitement de chaque conteneur
            for container in containers:
                container_data = processor.filter_by_container(prep_data, container)

                excel_path = None
                pdf_path = None
                container_success = True

                # ✅ Génération Excel
                if generer_excel:
                    try:
                        excel_path = processor.create_excel(
                            data=container_data,
                            container=container,
                            output_dir=session_dir,
                            cariste=cariste,
                            fournisseur=fournisseur,
                            numero_dossier=numero_dossier,
                            type_certification=type_certification,
                            numero_certificat=numero_certificat,
                            template_path=zzz_template_path
                        )
                        logger.info(f"✓ Excel généré avec succès pour {container}")
                    except Exception as excel_error:
                        logger.error(f"Erreur Excel pour {container}: {excel_error}")
                        messages.warning(request, f"Erreur génération Excel pour {container}: {str(excel_error)}")
                        container_success = False
                        failed_containers.append(container)
                        continue  # Passer au conteneur suivant

                # ✅ Génération PDF (fidèle à Excel)
                if generer_pdf and excel_path:
                    try:
                        pdf_path = pdf_generator.convert_excel_to_pdf(
                            excel_path=excel_path,
                            output_dir=session_dir,
                            container_name=container
                        )
                        if pdf_path:
                            logger.info(f"✓ PDF généré avec succès pour {container}")
                        else:
                            logger.warning(f"PDF non généré pour {container}")
                            messages.warning(request, f"PDF non généré pour {container}")
                    except Exception as pdf_error:
                        logger.error(f"Erreur PDF pour {container}: {pdf_error}")
                        pdf_path = None
                        messages.warning(request, f"Erreur génération PDF pour {container}: {str(pdf_error)}")

                # Sauvegarde dans la base
                if excel_path and os.path.exists(excel_path):
                    try:
                        GeneratedFile.objects.create(
                            file=os.path.relpath(excel_path, settings.MEDIA_ROOT),
                            file_type='excel',
                            container_name=container
                        )
                    except Exception as db_error:
                        logger.warning(f"Erreur sauvegarde BD Excel {container}: {db_error}")

                if pdf_path and os.path.exists(pdf_path):
                    try:
                        GeneratedFile.objects.create(
                            file=os.path.relpath(pdf_path, settings.MEDIA_ROOT),
                            file_type='pdf',
                            container_name=container
                        )
                    except Exception as db_error:
                        logger.warning(f"Erreur sauvegarde BD PDF {container}: {db_error}")

                if container_success:
                    successful_containers += 1
                    
                results.append({
                    'container': container,
                    'excel_path': excel_path,
                    'pdf_path': pdf_path,
                    'excel_filename': os.path.basename(excel_path) if excel_path and os.path.exists(excel_path) else 'Non généré',
                    'pdf_filename': os.path.basename(pdf_path) if pdf_path and os.path.exists(pdf_path) else 'Non généré',
                    'success': container_success
                })

            # ✅ Créer un ZIP global contenant tous les fichiers (pas de sous-dossiers)
            zip_path = None
            if successful_containers > 0:
                try:
                    zip_path = create_session_zip(session_dir, session_timestamp)
                    logger.info(f"✓ ZIP créé avec succès : {zip_path}")
                except Exception as zip_error:
                    logger.error(f"Erreur création ZIP: {zip_error}")
                    messages.warning(request, f"Erreur lors de la création du fichier ZIP: {str(zip_error)}")

            # Messages de résumé
            if successful_containers > 0:
                messages.success(request, f"Traitement terminé ! {successful_containers} conteneur(s) généré(s) avec succès.")
            if failed_containers:
                messages.error(request, f"{len(failed_containers)} conteneur(s) en échec: {', '.join(failed_containers)}")

            return render(request, 'upload.html', {
                'results': results,
                'show_results': True,
                'session_dir': session_dir,
                'zip_path': zip_path,
                'total_containers': len(containers),
                'successful_containers': successful_containers,
                'failed_containers': failed_containers,
                'cariste_utilise': cariste,
                'fournisseur_utilise': fournisseur,
                'numero_dossier_utilise': numero_dossier,
                'type_certification_utilise': type_certification,
                'numero_certificat_utilise': numero_certificat,
            })

        except Exception as e:
            logger.error(f"Erreur lors du traitement: {traceback.format_exc()}")
            messages.error(request, f"Erreur lors du traitement: {str(e)}")
            return render(request, 'upload.html')

    return render(request, 'upload.html')

def download_file(request):
    """Vue pour télécharger des fichiers individuels"""
    file_path = request.GET.get('file_path')
    
    if not file_path or not os.path.exists(file_path):
        messages.error(request, "Fichier non trouvé.")
        return redirect('home')
    
    try:
        response = FileResponse(
            open(file_path, 'rb'),
            as_attachment=True
        )
        response['Content-Disposition'] = f'attachment; filename="{os.path.basename(file_path)}"'
        return response
        
    except Exception as e:
        logger.error(f"Erreur lors du téléchargement: {traceback.format_exc()}")
        messages.error(request, f"Erreur lors du téléchargement: {str(e)}")
        return redirect('home')

def download_container_folder(request):
    """Vue pour télécharger le dossier complet d'un conteneur"""
    container_name = request.GET.get('container')
    session_dir = request.GET.get('session_dir')
    
    if not container_name or not session_dir:
        messages.error(request, "Paramètres manquants.")
        return redirect('home')
    
    try:
        # Créer un ZIP pour le conteneur spécifique
        container_zip_path = os.path.join(os.path.dirname(session_dir), f"{container_name}.zip")
        
        with zipfile.ZipFile(container_zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            # Ajouter tous les fichiers Excel et PDF de ce conteneur
            for filename in os.listdir(session_dir):
                if filename.startswith(container_name):
                    file_path = os.path.join(session_dir, filename)
                    zipf.write(file_path, filename)
        
        # Retourner le ZIP en téléchargement
        response = FileResponse(
            open(container_zip_path, 'rb'),
            as_attachment=True
        )
        response['Content-Disposition'] = f'attachment; filename="{container_name}.zip"'
        
        # Nettoyer le fichier ZIP temporaire après envoi
        try:
            os.remove(container_zip_path)
        except:
            pass
            
        return response
        
    except Exception as e:
        logger.error(f"Erreur lors du téléchargement du conteneur: {traceback.format_exc()}")
        messages.error(request, f"Erreur lors du téléchargement: {str(e)}")
        return redirect('home')

def download_session_zip(request):
    """Vue pour télécharger le ZIP complet de la session"""
    zip_path = request.GET.get('zip_path')
    
    if not zip_path or not os.path.exists(zip_path):
        messages.error(request, "Fichier ZIP non trouvé.")
        return redirect('home')
    
    try:
        response = FileResponse(
            open(zip_path, 'rb'),
            as_attachment=True
        )
        response['Content-Disposition'] = f'attachment; filename="{os.path.basename(zip_path)}"'
        return response
        
    except Exception as e:
        logger.error(f"Erreur lors du téléchargement du ZIP: {traceback.format_exc()}")
        messages.error(request, f"Erreur lors du téléchargement: {str(e)}")
        return redirect('home')

def create_session_zip(session_dir, session_timestamp):
    """Crée un ZIP contenant tous les fichiers Excel/PDF de la session."""
    zip_filename = f"fichiers_conteneurs_{session_timestamp}.zip"
    zip_path = os.path.join(os.path.dirname(session_dir), zip_filename)

    try:
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for root, dirs, files in os.walk(session_dir):
                for file in files:
                    file_path = os.path.join(root, file)
                    # Utiliser seulement le nom du fichier (pas le chemin complet)
                    arcname = os.path.basename(file_path)
                    zipf.write(file_path, arcname)
        
        logger.info(f"✓ ZIP créé avec succès: {zip_path}")
        return zip_path
        
    except Exception as e:
        logger.error(f"Erreur lors de la création du ZIP: {e}")
        raise

def clear_old_sessions(request):
    """Vue pour nettoyer les anciennes sessions (admin)"""
    try:
        base_output_dir = getattr(settings, 'CUSTOM_DOWNLOAD_DIR',
                                os.path.join(settings.MEDIA_ROOT, 'generated'))
        
        if not os.path.exists(base_output_dir):
            messages.info(request, "Aucun dossier de session à nettoyer.")
            return redirect('home')
        
        deleted_count = 0
        deleted_size = 0
        
        # Supprimer les sessions de plus de 7 jours
        for item in os.listdir(base_output_dir):
            item_path = os.path.join(base_output_dir, item)
            if os.path.isdir(item_path) and item.startswith('session_'):
                try:
                    # Calculer l'âge du dossier
                    stat = os.stat(item_path)
                    creation_time = stat.st_ctime
                    current_time = datetime.now().timestamp()
                    age_days = (current_time - creation_time) / (24 * 3600)
                    
                    if age_days > 7:  # Plus de 7 jours
                        # Calculer la taille avant suppression
                        size = get_folder_size(item_path)
                        shutil.rmtree(item_path)
                        deleted_count += 1
                        deleted_size += size
                        logger.info(f"Session supprimée: {item} ({size/1024/1024:.2f} MB)")
                        
                except Exception as e:
                    logger.warning(f"Impossible de supprimer {item}: {e}")
        
        # Supprimer les anciens fichiers ZIP
        for item in os.listdir(base_output_dir):
            item_path = os.path.join(base_output_dir, item)
            if os.path.isfile(item_path) and item.endswith('.zip'):
                try:
                    stat = os.stat(item_path)
                    creation_time = stat.st_ctime
                    current_time = datetime.now().timestamp()
                    age_days = (current_time - creation_time) / (24 * 3600)
                    
                    if age_days > 7:
                        size = os.path.getsize(item_path)
                        os.remove(item_path)
                        deleted_count += 1
                        deleted_size += size
                        logger.info(f"ZIP supprimé: {item} ({size/1024/1024:.2f} MB)")
                        
                except Exception as e:
                    logger.warning(f"Impossible de supprimer {item}: {e}")
        
        if deleted_count > 0:
            messages.success(request, f"{deleted_count} anciens fichiers/dossiers supprimés ({deleted_size/1024/1024:.2f} MB libérés).")
        else:
            messages.info(request, "Aucun fichier ancien à supprimer.")
            
    except Exception as e:
        logger.error(f"Erreur lors du nettoyage: {e}")
        messages.error(request, f"Erreur lors du nettoyage: {str(e)}")
    
    return redirect('home')

def get_folder_size(folder_path):
    """Calcule la taille totale d'un dossier"""
    total_size = 0
    for dirpath, dirnames, filenames in os.walk(folder_path):
        for filename in filenames:
            filepath = os.path.join(dirpath, filename)
            try:
                total_size += os.path.getsize(filepath)
            except OSError:
                pass
    return total_size

def list_sessions(request):
    """Vue pour lister toutes les sessions disponibles (admin)"""
    try:
        base_output_dir = getattr(settings, 'CUSTOM_DOWNLOAD_DIR',
                                os.path.join(settings.MEDIA_ROOT, 'generated'))
        
        if not os.path.exists(base_output_dir):
            messages.info(request, "Aucune session disponible.")
            return render(request, 'upload.html')
        
        sessions = []
        for item in os.listdir(base_output_dir):
            item_path = os.path.join(base_output_dir, item)
            if os.path.isdir(item_path) and item.startswith('session_'):
                try:
                    stat = os.stat(item_path)
                    creation_time = datetime.fromtimestamp(stat.st_ctime)
                    file_count = len([f for f in os.listdir(item_path) if os.path.isfile(os.path.join(item_path, f))])
                    size = get_folder_size(item_path)
                    
                    sessions.append({
                        'name': item,
                        'path': item_path,
                        'creation_time': creation_time,
                        'file_count': file_count,
                        'size_mb': size / 1024 / 1024
                    })
                except Exception as e:
                    logger.warning(f"Impossible d'analyser {item}: {e}")
        
        # Trier par date de création (plus récent en premier)
        sessions.sort(key=lambda x: x['creation_time'], reverse=True)
        
        return render(request, 'sessions.html', {
            'sessions': sessions,
            'total_sessions': len(sessions)
        })
        
    except Exception as e:
        logger.error(f"Erreur lors du listing des sessions: {e}")
        messages.error(request, f"Erreur lors du listing des sessions: {str(e)}")
        return redirect('home')