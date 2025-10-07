# pdf_generator.py - VERSION CORRIGÉE POUR CHEMINS AVEC CARACTÈRES SPÉCIAUX

import os
import pandas as pd
from datetime import datetime
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import logging
import barcode
from barcode.writer import ImageWriter
import tempfile
import openpyxl
from io import BytesIO
import base64

logger = logging.getLogger(__name__)

class PDFGenerator:
    def __init__(self):
        self.styles = getSampleStyleSheet()
        self._setup_custom_styles()
        logger.info("PDFGenerator initialisé avec support codes-barres")
    
    def _setup_custom_styles(self):
        """Configure les styles personnalisés"""
        try:
            pdfmetrics.registerFont(TTFont('Arial', 'arial.ttf'))
            self.default_font = 'Arial'
        except:
            self.default_font = 'Helvetica'
        
        self.title_style = ParagraphStyle(
            'CustomTitle',
            parent=self.styles['Heading1'],
            fontName=self.default_font,
            fontSize=16,
            spaceAfter=30,
            alignment=1
        )
        
        self.normal_style = ParagraphStyle(
            'CustomNormal',
            parent=self.styles['Normal'],
            fontName=self.default_font,
            fontSize=10,
            leading=12
        )
        
        self.small_style = ParagraphStyle(
            'CustomSmall',
            parent=self.styles['Normal'],
            fontName=self.default_font,
            fontSize=8,
            leading=10
        )
    
    def create_pdf_from_excel(self, excel_path, container, output_dir,
                             cariste="FIRN", fournisseur="FIRN", 
                             numero_dossier="DSF23044",
                             type_certification="FSC RECYCLED 100%", 
                             numero_certificat="CU-COC-903458"):
        """Transforme le fichier Excel traité en PDF avec codes-barres"""
        try:
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            pdf_filename = f"SEMBA_RECEPTION_{container}_{timestamp}.pdf"
            pdf_path = os.path.join(output_dir, pdf_filename)
            
            logger.info(f"Conversion Excel → PDF pour {container}")
            
            # Lire les données réelles de l'Excel
            data_rows, headers = self._read_excel_actual_data(excel_path)
            
            if not data_rows:
                logger.warning(f"Aucune donnée trouvée dans l'Excel pour {container}")
                return None
            
            logger.info(f"Données trouvées: {len(data_rows)} lignes")
            
            # Créer le PDF
            doc = SimpleDocTemplate(
                pdf_path,
                pagesize=A4,
                rightMargin=20,
                leftMargin=20,
                topMargin=30,
                bottomMargin=30
            )
            
            story = []
            
            # Titre principal
            story.extend(self._create_title_section())
            
            # Informations générales
            story.extend(self._create_info_section(container, cariste, numero_dossier))
            story.append(Spacer(1, 20))
            
            # Tableau des données avec codes-barres
            table_elements = self._create_structured_table_with_barcodes(data_rows, headers)
            story.extend(table_elements)
            
            # Générer le PDF
            doc.build(story)
            
            logger.info(f"✓ PDF créé avec succès: {pdf_path}")
            return pdf_path
            
        except Exception as e:
            logger.error(f"✗ Erreur conversion PDF {container}: {str(e)}", exc_info=True)
            raise

    def _read_excel_actual_data(self, excel_path):
        """Lit les données RÉELLES de l'Excel généré par ExcelProcessor"""
        try:
            # Charger le workbook avec openpyxl
            workbook = openpyxl.load_workbook(excel_path, data_only=True)
            sheet = workbook['FO57']
            
            data_rows = []
            headers = [
                'N°', 'N° Fournisseur', 'Fournisseur', 'Code barre N° bobine/N° SEMBA',
                'Référence', 'Diamètre', 'Poids (KG)', 'N° Certificat FSC',
                'Type de certification FSC', 'Observation MAGASIN par rapport au papier reçu'
            ]
            
            # Parcourir les lignes pour trouver les données
            for row in sheet.iter_rows(values_only=True):
                # Chercher les lignes qui commencent par un numéro (1, 2, 3...)
                if row and row[0] is not None and str(row[0]).strip().isdigit():
                    row_data = []
                    for i, cell in enumerate(row):
                        if i >= 10:  # Seulement les 10 premières colonnes
                            break
                        if cell is None:
                            row_data.append("")
                        else:
                            row_data.append(str(cell).strip())
                    
                    # Compléter avec des valeurs vides si nécessaire
                    while len(row_data) < 10:
                        row_data.append("")
                    
                    data_rows.append(row_data)
            
            workbook.close()
            logger.info(f"Données réelles extraites: {len(data_rows)} lignes")
            return data_rows, headers
            
        except Exception as e:
            logger.error(f"Erreur lecture données réelles Excel: {str(e)}")
            return [], []
    
    def _create_title_section(self):
        """Crée la section titre"""
        elements = []
        title = Paragraph("FICHE DE CONTRÔLE RECEPTION DES BOBINES", self.title_style)
        elements.append(title)
        elements.append(Spacer(1, 15))
        return elements
    
    def _create_info_section(self, container, cariste, numero_dossier):
        """Crée la section d'informations générales"""
        elements = []
        info_text = (
            f"<b>CARISTE :</b> {cariste} &nbsp;&nbsp;&nbsp; "
            f"<b>N° CT :</b> {container} &nbsp;&nbsp;&nbsp; "
            f"<b>DATE :</b> {datetime.now().strftime('%d/%m/%Y')} &nbsp;&nbsp;&nbsp; "
            f"<b>No. Dossier :</b> {numero_dossier}"
        )
        info_paragraph = Paragraph(info_text, self.normal_style)
        elements.append(info_paragraph)
        elements.append(Spacer(1, 15))
        return elements
    
    def _create_structured_table_with_barcodes(self, data_rows, headers):
        """Crée le tableau avec codes-barres intégrés"""
        elements = []
        
        if not data_rows:
            no_data = Paragraph("Aucune donnée disponible", self.normal_style)
            elements.append(no_data)
            return elements
        
        # Préparer les en-têtes
        formatted_headers = []
        for header in headers:
            if len(header) > 15:
                formatted_header = Paragraph(header.replace(' ', '<br/>'), self.small_style)
            else:
                formatted_header = Paragraph(header, self.small_style)
            formatted_headers.append(formatted_header)
        
        table_data = [formatted_headers]
        
        # Ajouter les données avec codes-barres
        barcode_success_count = 0
        
        for row_idx, row in enumerate(data_rows):
            try:
                formatted_row = []
                for col_idx, cell_value in enumerate(row):
                    if col_idx == 3:  # Colonne "Code barre"
                        # Utiliser le "N° Fournisseur" (colonne 1) pour générer le code-barres
                        barcode_value = row[1] if len(row) > 1 else cell_value
                        
                        barcode_image = self._generate_barcode_safe(barcode_value, row_idx)
                        if barcode_image:
                            formatted_cell = barcode_image
                            barcode_success_count += 1
                        else:
                            # Fallback : afficher le texte formaté
                            formatted_cell = Paragraph(f"<b>⎕ {barcode_value}</b>", self.small_style)
                    elif col_idx in [8, 9]:  # Colonnes avec texte long
                        if cell_value and len(cell_value) > 20:
                            formatted_cell = Paragraph(cell_value, self.small_style)
                        else:
                            formatted_cell = cell_value
                    else:
                        formatted_cell = cell_value
                    formatted_row.append(formatted_cell)
                
                table_data.append(formatted_row)
                
            except Exception as e:
                logger.error(f"Erreur formatage ligne {row_idx}: {e}")
                continue
        
        logger.info(f"Résumé codes-barres: {barcode_success_count}/{len(data_rows)} générés avec succès")
        
        # Largeurs de colonnes
        col_widths = [25, 70, 40, 120, 50, 40, 50, 70, 80, 80]
        
        try:
            table = Table(table_data, colWidths=col_widths, repeatRows=1)
            
            table_style = TableStyle([
                # En-têtes
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#2E4057')),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                ('FONTNAME', (0, 0), (-1, 0), self.default_font),
                ('FONTSIZE', (0, 0), (-1, 0), 7),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 8),
                ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
                ('VALIGN', (0, 0), (-1, 0), 'MIDDLE'),
                
                # Données
                ('BACKGROUND', (0, 1), (-1, -1), colors.white),
                ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
                ('FONTNAME', (0, 1), (-1, -1), self.default_font),
                ('FONTSIZE', (0, 1), (-1, -1), 7),
                ('ALIGN', (0, 1), (-1, -1), 'CENTER'),
                ('VALIGN', (0, 1), (-1, -1), 'MIDDLE'),
                
                # Grille
                ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
                ('LEFTPADDING', (0, 0), (-1, -1), 3),
                ('RIGHTPADDING', (0, 0), (-1, -1), 3),
                ('TOPPADDING', (0, 0), (-1, -1), 2),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
            ])
            
            table.setStyle(table_style)
            elements.append(table)
            
        except Exception as e:
            logger.error(f"Erreur création tableau: {e}")
            elements.extend(self._create_simple_fallback(data_rows))
        
        return elements
    
    def _generate_barcode_safe(self, code_value, row_index):
        """Génère un code-barres de manière sécurisée sans problèmes de chemin"""
        try:
            if not code_value or not isinstance(code_value, str):
                return None
            
            clean_code = str(code_value).strip()
            if not clean_code:
                return None
            
            # Nettoyer le code
            clean_code = ''.join(c for c in clean_code if c.isalnum() or c in '-_')
            
            # Méthode SÉCURISÉE : générer en mémoire sans fichiers temporaires
            barcode_image = self._generate_barcode_in_memory(clean_code)
            if barcode_image:
                return barcode_image
            
            # Fallback : méthode avec chemin sécurisé
            return self._generate_barcode_secure_path(clean_code)
            
        except Exception as e:
            logger.error(f"Erreur génération code-barres sécurisée: {e}")
            return None
    
    def _generate_barcode_in_memory(self, clean_code):
        """Génère le code-barres entièrement en mémoire (méthode recommandée)"""
        try:
            from PIL import Image as PILImage
            import barcode
            from barcode.writer import ImageWriter
            from io import BytesIO
            
            # Générer le code-barres en mémoire
            code128 = barcode.get_barcode_class('code128')
            
            writer_options = {
                'write_text': False,
                'quiet_zone': 2.0,
                'module_height': 15.0,
                'module_width': 0.3,
                'background': 'white',
                'foreground': 'black',
            }
            
            # Créer un buffer en mémoire
            buffer = BytesIO()
            
            # Générer et sauvegarder dans le buffer
            barcode_obj = code128(clean_code, writer=ImageWriter())
            barcode_obj.write(buffer, options=writer_options)
            
            # Réinitialiser le buffer
            buffer.seek(0)
            
            # Créer l'image ReportLab directement depuis le buffer
            barcode_img = Image(buffer, width=100, height=35)
            
            logger.info(f"✓ Code-barres généré en mémoire: {clean_code}")
            return barcode_img
            
        except Exception as e:
            logger.warning(f"Échec génération en mémoire: {e}")
            return None
    
    def _generate_barcode_secure_path(self, clean_code):
        """Génère le code-barres avec un chemin sécurisé (sans caractères spéciaux)"""
        try:
            import barcode
            from barcode.writer import ImageWriter
            
            # Créer un répertoire temporaire sécurisé
            safe_temp_dir = tempfile.mkdtemp(prefix='barcode_')
            temp_path = os.path.join(safe_temp_dir, f"barcode_{clean_code}.png")
            
            # Générer le code-barres
            code128 = barcode.get_barcode_class('code128')
            
            writer_options = {
                'write_text': False,
                'quiet_zone': 2.0,
                'module_height': 15.0,
                'module_width': 0.3,
                'background': 'white',
                'foreground': 'black',
            }
            
            barcode_obj = code128(clean_code, writer=ImageWriter())
            barcode_obj.save(temp_path, options=writer_options)
            
            # Vérifier que le fichier a été créé
            if os.path.exists(temp_path) and os.path.getsize(temp_path) > 100:
                # Lire le fichier et créer l'image
                with open(temp_path, 'rb') as f:
                    from io import BytesIO
                    buffer = BytesIO(f.read())
                
                barcode_img = Image(buffer, width=100, height=35)
                
                # Nettoyer
                os.unlink(temp_path)
                os.rmdir(safe_temp_dir)
                
                logger.info(f"✓ Code-barres généré avec chemin sécurisé: {clean_code}")
                return barcode_img
            else:
                # Nettoyer en cas d'échec
                if os.path.exists(temp_path):
                    os.unlink(temp_path)
                os.rmdir(safe_temp_dir)
                return None
                
        except Exception as e:
            logger.warning(f"Échec génération avec chemin sécurisé: {e}")
            # Nettoyer en cas d'erreur
            try:
                if 'temp_path' in locals() and os.path.exists(temp_path):
                    os.unlink(temp_path)
                if 'safe_temp_dir' in locals() and os.path.exists(safe_temp_dir):
                    os.rmdir(safe_temp_dir)
            except:
                pass
            return None

    def _create_simple_fallback(self, data_rows):
        """Version de secours simple"""
        elements = []
        try:
            simple_headers = ['N°', 'N° Fournisseur', 'Fournisseur', 'Code barre', 'Référence', 'Diamètre', 'Poids']
            table_data = [simple_headers]
            
            for row in data_rows:
                if len(row) >= 7:
                    simple_row = [row[0], row[1], row[2], row[1], row[4], row[5], row[6]]
                    table_data.append(simple_row)
            
            table = Table(table_data, colWidths=[30, 80, 50, 80, 60, 50, 50], repeatRows=1)
            table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ('FONTSIZE', (0, 0), (-1, -1), 8),
            ]))
            elements.append(table)
        except Exception as e:
            logger.error(f"Erreur fallback: {e}")
        
        return elements

def create_pdf_from_excel(excel_path, output_dir, container, **kwargs):
    generator = PDFGenerator()
    return generator.create_pdf_from_excel(
        excel_path=excel_path,
        output_dir=output_dir,
        container=container,
        **kwargs
    )