
import os
import subprocess
import logging
from pathlib import Path
import pythoncom
import win32com.client

logger = logging.getLogger(__name__)

class PDFGenerator:
    def __init__(self):
        logger.info("PDFGenerator initialisé (Excel → PDF avec win32com)")

    def convert_excel_to_pdf(self, excel_path, output_dir, container_name=None):
        """
        Convertit Excel en PDF avec win32com (Windows seulement)
        """
        try:
            if not excel_path or not os.path.exists(excel_path):
                raise FileNotFoundError(f"Fichier Excel introuvable : {excel_path}")

            os.makedirs(output_dir, exist_ok=True)

            if not container_name:
                container_name = Path(excel_path).stem

            pdf_filename = f"{container_name}.pdf"
            pdf_path = os.path.join(output_dir, pdf_filename)

            logger.info(f"Conversion Excel → PDF : {excel_path} → {pdf_path}")

            #Initialisation COM
            
            pythoncom.CoInitialize()

            # Utilisation de win32com
            
            
            excel_app = None
            workbook = None
            
            try:
                excel_app = win32com.client.Dispatch("Excel.Application")
                excel_app.Visible = False
                excel_app.DisplayAlerts = False

                workbook = excel_app.Workbooks.Open(os.path.abspath(excel_path))
                
                # Export en PDF
                workbook.ExportAsFixedFormat(0, os.path.abspath(pdf_path))  # 0 = xlTypePDF
                
                workbook.Close(SaveChanges=False)
                
                if os.path.exists(pdf_path):
                    logger.info(f" PDF créé avec succès : {pdf_path}")
                    return pdf_path
                else:
                    logger.warning(f" PDF non généré : {pdf_path}")
                    return None
                    
            except Exception as e:
                logger.error(f"Erreur lors de la conversion Excel: {e}")
                return None
            finally:
                # Fermeture propre dans le bon ordre
                try:
                    if workbook:
                        workbook.Close(SaveChanges=False)
                except:
                    pass
                
                try:
                    if excel_app:
                        excel_app.Quit()
                except:
                    pass
                
                # Libération des objets COM
                try:
                    if workbook:
                        del workbook
                    if excel_app:
                        del excel_app
                except:
                    pass
                
                #  Désinitialisation COM
                pythoncom.CoUninitialize()

        except ImportError:
            logger.error("win32com non disponible - PDF non généré")
            return None
        except Exception as e:
            logger.error(f"Erreur lors de la conversion Excel → PDF : {e}")
            return None

def create_pdf_from_excel(excel_path, output_dir, container_name=None):
    """
    Fonction simplifiée pour conversion directe.
    """
    generator = PDFGenerator()
    return generator.convert_excel_to_pdf(excel_path, output_dir, container_name)