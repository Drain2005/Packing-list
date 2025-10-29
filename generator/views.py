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
import zipfile
import time

logger = logging.getLogger(__name__)

def home(request):
    """Vue principale — Upload, génération Excel/PDF et affichage des résultats"""

    # 🔹 Téléchargement direct via ?download=...&file_path=...
    file_type = request.GET.get('download')
    file_path = request.GET.get('file_path')
    if file_type and file_path and os.path.exists(file_path):
        return FileResponse(open(file_path, 'rb'), as_attachment=True)

    
    if request.method == 'POST':
        
        
        print(" === DÉBUT TRAITEMENT ===")
        logger.info(" DÉBUT TRAITEMENT UPLOAD")
        start_time = time.time()
        
        prep_file = request.FILES.get('preparation_pl')
        zzz_file = request.FILES.get('zzz_file')

        if not prep_file:
            messages.error(request, "Le fichier Preparation PL est obligatoire.")
            return render(request, 'upload.html')

        #  Générer toujours Excel et PDF
        generer_excel = True
        generer_pdf = True

        #  Supprimer les valeurs par défaut pour forcer les données utilisateur
        cariste = request.POST.get('cariste', '').strip()
        fournisseur = request.POST.get('fournisseur', '').strip()
        numero_dossier = request.POST.get('numero_dossier', '').strip()
        type_certification = request.POST.get('type_certification', '').strip()
        numero_certificat = request.POST.get('numero_certificat', '').strip()

        #  LOG des données du formulaire
        print("📋 DONNÉES FORMULAIRE:")
        print(f"  Cariste: '{cariste}'")
        print(f"  Fournisseur: '{fournisseur}'")
        print(f"  Numéro dossier: '{numero_dossier}'")
        logger.info("=== DONNÉES FORMULAIRE RECEUILLIES ===")
        logger.info(f"Cariste: '{cariste}'")
        logger.info(f"Fournisseur: '{fournisseur}'")
        logger.info(f"Numéro dossier: '{numero_dossier}'")
        logger.info(f"Type certification: '{type_certification}'")
        logger.info(f"Numéro certificat: '{numero_certificat}'")

        try:
            #  Sauvegarde fichiers
            print("1.  Sauvegarde des fichiers...")
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
            print("    Fichiers sauvegardés")

            processor = ExcelProcessor()
            pdf_generator = PDFGenerator()

            #  Configuration template
            print("2. ⚙️ Configuration template...")
            if zzz_file and zzz_obj:
                processor.set_template(zzz_obj.file.path)
                print(f"    Template défini: {zzz_obj.file.path}")
                logger.info(f"Template zzzz.xlsx défini : {zzz_obj.file.path}")
            else:
                print("    Aucun template uploadé")
                logger.warning("Aucun template zzzz.xlsx uploadé fourni")

            #  Lecture fichier principal
            print("3.  Lecture fichier Excel...")
            prep_data, columns = processor.read_excel_file(prep_obj.file.path)
            print(f"    Fichier lu: {len(prep_data)} lignes, {len(columns)} colonnes")
            
            # : Extraction conteneurs
            print("4.  Extraction conteneurs...")
            containers = processor.extract_containers(prep_data)
            print(f"    Conteneurs trouvés: {containers}")
            
            if not containers:
                messages.error(request, "Aucun conteneur trouvé dans le fichier.")
                return render(request, 'upload.html')

            # Dossier principal 
            base_output_dir = getattr(settings, 'CUSTOM_DOWNLOAD_DIR',
                                      os.path.join(settings.MEDIA_ROOT, 'generated'))
            os.makedirs(base_output_dir, exist_ok=True)

            #  sous-dossier session
            session_timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            session_dir = os.path.join(base_output_dir, f"session_{session_timestamp}")
            os.makedirs(session_dir, exist_ok=True)
            print(f"    Dossier session: {session_dir}")

            results = []

            # Traitement de chaque conteneur
            print(f"5. Traitement de {len(containers)} conteneurs...")
            for i, container in enumerate(containers):
                container_start = time.time()
                print(f"    Conteneur {i+1}/{len(containers)}: {container}")
                
                container_data = processor.filter_by_container(prep_data, container)
                print(f"       Données: {len(container_data)} bobines")

                excel_path = None
                pdf_path = None

                #  Génération Excel (TOUJOURS)
                print(f"       Génération Excel...")
                excel_path = processor.create_excel(
                    data=container_data,
                    container=container,
                    output_dir=session_dir,
                    cariste=cariste,
                    fournisseur=fournisseur,
                    numero_dossier=numero_dossier,
                    type_certification=type_certification,
                    numero_certificat=numero_certificat
                )
                if excel_path:
                    print(f"       Excel généré: {os.path.basename(excel_path)}")
                else:
                    print(f"       Erreur génération Excel")

                #  Génération PDF 
                if excel_path:
                    print(f"   Génération PDF...")
                    pdf_path = pdf_generator.convert_excel_to_pdf(
                        excel_path=excel_path,
                        output_dir=session_dir,
                        container_name=container
                    )
                    if pdf_path:
                        print(f"      PDF généré: {os.path.basename(pdf_path)}")
                    else:
                        print(f"       Erreur génération PDF")

                #  Enregistrement dans la base
                if excel_path:
                    GeneratedFile.objects.create(
                        file=os.path.relpath(excel_path, settings.MEDIA_ROOT),
                        file_type='excel',
                        container_name=container
                    )
                if pdf_path:
                    GeneratedFile.objects.create(
                        file=os.path.relpath(pdf_path, settings.MEDIA_ROOT),
                        file_type='pdf',
                        container_name=container
                    )

                #  Ajout des résultats
                results.append({
                    'container': container,
                    'excel_path': excel_path,
                    'pdf_path': pdf_path,
                    'excel_filename': os.path.basename(excel_path) if excel_path else 'Non généré',
                    'pdf_filename': os.path.basename(pdf_path) if pdf_path else 'Non généré',
                })
                
                container_time = time.time() - container_start
                print(f"        Temps conteneur: {container_time:.2f}s")

            # ⭐ ÉTAPE 6: Création ZIP
            print("6.  Création ZIP...")
            zip_path = create_session_zip(session_dir, session_timestamp)
            print(f"    ZIP créé: {zip_path}")

            total_time = time.time() - start_time
            print(f" TRAITEMENT TERMINÉ - Temps total: {total_time:.2f}s")
            logger.info(f" TRAITEMENT TERMINÉ - {len(containers)} conteneurs en {total_time:.2f}s")

            return render(request, 'upload.html', {
                'results': results,
                'show_results': True,
                'session_dir': session_dir,
                'zip_path': zip_path,
                'total_containers': len(containers),
                'cariste_utilise': cariste,
                'fournisseur_utilise': fournisseur,
                'numero_dossier_utilise': numero_dossier,
                'type_certification_utilise': type_certification,
                'numero_certificat_utilise': numero_certificat,
            })

        except Exception as e:
            error_time = time.time() - start_time
            print(f" ERREUR après {error_time:.2f}s: {str(e)}")
            logger.error(f"Erreur lors du traitement: {traceback.format_exc()}")
            messages.error(request, f"Erreur: {str(e)}")
            return render(request, 'upload.html')

    return render(request, 'upload.html')


def create_session_zip(session_dir, session_timestamp):
    """Crée un ZIP contenant tous les fichiers Excel/PDF de la session."""
    zip_filename = f"fichiers_conteneurs_{session_timestamp}.zip"
    zip_path = os.path.join(os.path.dirname(session_dir), zip_filename)

    print(f"  Création ZIP: {zip_path}")
    
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root, dirs, files in os.walk(session_dir):
            for file in files:
                file_path = os.path.join(root, file)
                arcname = os.path.basename(file_path)
                zipf.write(file_path, arcname)
                print(f"  Ajout: {file}")

    print(f"  ZIP créé avec {len(files)} fichiers")
    return zip_path


def download_file(request):
    """Télécharge un fichier individuel (Excel ou PDF)."""
    file_path = request.GET.get('file_path')

    if not file_path or not os.path.exists(file_path):
        messages.error(request, "Fichier non trouvé")
        return redirect('home')

    try:
        response = FileResponse(open(file_path, 'rb'), as_attachment=True)
        response['Content-Disposition'] = f'attachment; filename="{os.path.basename(file_path)}"'
        return response
    except Exception as e:
        logger.error(f"Erreur téléchargement fichier {file_path}: {e}")
        messages.error(request, f"Erreur lors du téléchargement: {str(e)}")
        return redirect('home')