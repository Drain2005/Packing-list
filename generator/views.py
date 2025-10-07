import os
import pandas as pd
import logging
import traceback
from django.shortcuts import render, redirect
from django.http import FileResponse, HttpResponse
from django.conf import settings
from django.contrib import messages
from .models import UploadedFile, GeneratedFile
from .utils.excel_processor import ExcelProcessor
from .utils.pdf_generator import PDFGenerator
from datetime import datetime
import shutil

logger = logging.getLogger(__name__)

def home(request):
    # Gestion du téléchargement de fichiers
    file_type = request.GET.get('download')
    file_path = request.GET.get('file_path')
    
    if file_type and file_path:
        if os.path.exists(file_path):
            response = FileResponse(open(file_path, 'rb'), as_attachment=True)
            return response
        else:
            messages.error(request, "Fichier non trouvé")
    
    # Traitement du formulaire POST
    if request.method == 'POST':
        # Récupération des fichiers
        prep_file = request.FILES.get('preparation_pl')
        zzz_file = request.FILES.get('zzz_file')
        download_type = request.POST.get('download_type')
        file_to_download = request.POST.get('file_to_download')
        
        # Gestion du téléchargement depuis les résultats
        if download_type and file_to_download:
            if os.path.exists(file_to_download):
                response = FileResponse(open(file_to_download, 'rb'), as_attachment=True)
                return response
            else:
                messages.error(request, "Fichier non trouvé")
                return render(request, 'upload.html')
        
        # Validation du fichier obligatoire
        if not prep_file:
            messages.error(request, "Le fichier Preparation PL est obligatoire")
            return render(request, 'upload.html')
        
        # Récupération des paramètres
        cariste = request.POST.get('cariste', 'FIRN')
        fournisseur = request.POST.get('fournisseur', 'FIRN')
        numero_dossier = request.POST.get('numero_dossier', 'DSF23044')
        type_certification = request.POST.get('type_certification', 'FSC RECYCLED 100%')
        numero_certificat = request.POST.get('numero_certificat', 'CU-COC-903458')
        generer_excel = request.POST.get('generer_excel') == 'on'
        generer_pdf = request.POST.get('generer_pdf') == 'on'
        file_prefix = request.POST.get('file_prefix', 'SEMBA_RECEPTION')
        
        # Validation qu'au moins un format est sélectionné
        if not generer_excel and not generer_pdf:
            messages.error(request, "Veuillez sélectionner au moins un format de sortie (Excel ou PDF)")
            return render(request, 'upload.html')
        
        try:
            # Sauvegarde des fichiers uploadés
            prep_obj = UploadedFile.objects.create(
                file=prep_file,
                file_type='Préparation_PL',
                original_name=prep_file.name
            )
            
            zzz_obj = None
            if zzz_file:
                zzz_obj = UploadedFile.objects.create(
                    file=zzz_file,
                    file_type='zzzz',
                    original_name=zzz_file.name
                )
            
            # Traitement des fichiers Excel
            processor = ExcelProcessor()
            pdf_generator = PDFGenerator()
            
            # Lecture des fichiers
            prep_data, prep_headers = processor.read_excel_file(prep_obj.file.path)
            zzz_data, zzz_headers = None, None
            if zzz_obj:
                zzz_data, zzz_headers = processor.read_excel_file(zzz_obj.file.path)
            
            # Extraction des conteneurs
            containers = processor.extract_containers(prep_data)
            
            if not containers:
                messages.error(request, "Aucun conteneur trouvé dans le fichier")
                return render(request, 'upload.html')
            
            results = []
            
            # Dossier de sortie principal
            base_output_dir = getattr(settings, 'CUSTOM_DOWNLOAD_DIR', 
                                    os.path.join(settings.MEDIA_ROOT, 'generated'))
            
            # Créer un sous-dossier avec timestamp pour cette session
            session_timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            session_dir = os.path.join(base_output_dir, f"session_{session_timestamp}")
            os.makedirs(session_dir, exist_ok=True)
            
            # Traitement par conteneur
            for container in containers:
                # Créer un dossier spécifique pour ce conteneur
                container_dir = os.path.join(session_dir, f"conteneur_{container}")
                os.makedirs(container_dir, exist_ok=True)
                
                container_data = processor.filter_by_container(prep_data, container)
                
                excel_path = None
                pdf_path = None
                
                # Création du fichier Excel si demandé
                if generer_excel:
                    try:
                        # Utiliser le template ZZZZ si disponible
                        if zzz_obj and zzz_obj.file.path:
                            excel_path = processor.create_excel_from_zzz_template(
                                data=container_data,
                                container=container,
                                output_dir=container_dir,  # Dossier du conteneur
                                zzz_file_path=zzz_obj.file.path,
                                cariste=cariste,
                                fournisseur=fournisseur,
                                numero_dossier=numero_dossier,
                                type_certification=type_certification,
                                numero_certificat=numero_certificat
                            )
                        else:
                            # Fallback : création manuelle
                            excel_path = processor.create_excel(
                                data=container_data,
                                container=container,
                                output_dir=container_dir,  # Dossier du conteneur
                                cariste=cariste,
                                fournisseur=fournisseur,
                                numero_dossier=numero_dossier,
                                type_certification=type_certification,
                                numero_certificat=numero_certificat
                            )
                    except Exception as e:
                        logger.error(f"Erreur création Excel {container}: {traceback.format_exc()}")
                        messages.error(request, f"Erreur création Excel pour {container}: {str(e)}")
                        continue
                
                # Création du fichier PDF si demandé
                if generer_pdf and excel_path:
                    try:
                        pdf_path = pdf_generator.create_pdf_from_excel(
                            excel_path=excel_path,
                            container=container,
                            output_dir=container_dir,  # Dossier du conteneur
                            cariste=cariste,
                            fournisseur=fournisseur,
                            numero_dossier=numero_dossier,
                            type_certification=type_certification,
                            numero_certificat=numero_certificat
                        )
                    except Exception as e:
                        logger.error(f"Erreur création PDF {container}: {e}")
                        messages.warning(request, f"Erreur lors de la génération du PDF: {e}")
                        pdf_path = None
                
                # Sauvegarde en base de données
                if excel_path:
                    excel_filename = os.path.basename(excel_path)
                    GeneratedFile.objects.create(
                        file=f'telechargements/{excel_filename}',
                        file_type='excel',
                        container_name=container
                    )
                
                if pdf_path:
                    pdf_filename = os.path.basename(pdf_path)
                    GeneratedFile.objects.create(
                        file=f'telechargements/{pdf_filename}',
                        file_type='pdf',
                        container_name=container
                    )
                
                results.append({
                    'container': container,
                    'container_dir': container_dir,  # Ajout du chemin du dossier
                    'excel_path': excel_path,
                    'pdf_path': pdf_path,
                    'excel_filename': os.path.basename(excel_path) if excel_path else 'Non généré',
                    'pdf_filename': os.path.basename(pdf_path) if pdf_path else 'Non généré',
                })
            
            # Créer un fichier ZIP de tous les dossiers de conteneurs
            zip_path = None
            if results:
                try:
                    zip_path = create_containers_zip(session_dir, session_timestamp)
                except Exception as e:
                    logger.error(f"Erreur création ZIP: {e}")
            
            context = {
                'results': results,
                'show_results': True,
                'cariste_utilise': cariste,
                'fournisseur_utilise': fournisseur,
                'numero_dossier_utilise': numero_dossier,
                'type_certification_utilise': type_certification,
                'numero_certificat_utilise': numero_certificat,
                'session_dir': session_dir,
                'zip_path': zip_path,
                'total_containers': len(containers),
            }
            return render(request, 'upload.html', context)
            
        except Exception as e:
            logger.error(f"Erreur lors du traitement: {traceback.format_exc()}")
            messages.error(request, f"Erreur lors du traitement: {str(e)}")
            return render(request, 'upload.html')
    
    return render(request, 'upload.html')

def create_containers_zip(session_dir, session_timestamp):
    """Crée un ZIP de tous les dossiers de conteneurs"""
    import zipfile
    
    zip_filename = f"conteneurs_complets_{session_timestamp}.zip"
    zip_path = os.path.join(os.path.dirname(session_dir), zip_filename)
    
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root, dirs, files in os.walk(session_dir):
            for file in files:
                file_path = os.path.join(root, file)
                # Créer le chemin relatif dans le ZIP
                arcname = os.path.relpath(file_path, os.path.dirname(session_dir))
                zipf.write(file_path, arcname)
    
    return zip_path

def download_file(request):
    """Télécharge un fichier individuel"""
    file_path = request.GET.get('file_path')
    
    if not file_path or not os.path.exists(file_path):
        messages.error(request, "Fichier non trouvé")
        return redirect('home')
    
    try:
        response = FileResponse(open(file_path, 'rb'), as_attachment=True)
        filename = os.path.basename(file_path)
        response['Content-Disposition'] = f'attachment; filename="{filename}"'
        return response
    except Exception as e:
        logger.error(f"Erreur téléchargement fichier {file_path}: {e}")
        messages.error(request, f"Erreur lors du téléchargement: {str(e)}")
        return redirect('home')

def download_container_folder(request):
    """Télécharge un dossier conteneur complet"""
    container_name = request.GET.get('container_name')  # Récupérer depuis GET
    session_dir = request.GET.get('session_dir')
    
    if not container_name:
        return HttpResponse("Nom du conteneur manquant", status=400)
    
    if not session_dir or not os.path.exists(session_dir):
        return HttpResponse("Dossier session non trouvé", status=404)
    
    container_dir = os.path.join(session_dir, f"conteneur_{container_name}")
    if not os.path.exists(container_dir):
        return HttpResponse(f"Dossier conteneur {container_name} non trouvé", status=404)
    
    # Créer un ZIP du dossier conteneur
    import zipfile
    import tempfile
    
    try:
        # Créer un fichier ZIP temporaire
        with tempfile.NamedTemporaryFile(delete=False, suffix='.zip') as tmp_file:
            zip_path = tmp_file.name
        
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for root, dirs, files in os.walk(container_dir):
                for file in files:
                    file_path = os.path.join(root, file)
                    # Créer le chemin relatif dans le ZIP
                    arcname = os.path.relpath(file_path, container_dir)
                    zipf.write(file_path, f"conteneur_{container_name}/{arcname}")
        
        # Servir le fichier ZIP
        response = FileResponse(open(zip_path, 'rb'), content_type='application/zip')
        response['Content-Disposition'] = f'attachment; filename="conteneur_{container_name}.zip"'
        
        # Nettoyer le fichier ZIP temporaire après envoi
        import threading
        def cleanup_temp_file():
            import time
            time.sleep(1)  # Attendre que le fichier soit envoyé
            try:
                if os.path.exists(zip_path):
                    os.remove(zip_path)
            except:
                pass
        
        threading.Thread(target=cleanup_temp_file).start()
        
        return response
        
    except Exception as e:
        logger.error(f"Erreur création ZIP conteneur {container_name}: {e}")
        # Nettoyer en cas d'erreur
        if os.path.exists(zip_path):
            os.remove(zip_path)
        return HttpResponse("Erreur lors de la création du ZIP", status=500)