
import pandas as pd
import os
import re
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.drawing.image import Image
import barcode
from barcode.writer import ImageWriter
import logging
from django.conf import settings
import tempfile

logger = logging.getLogger(__name__)

class ExcelProcessor:
    def __init__(self):
        self.container_column = None
    
    def read_excel_file(self, file_path):
        """Lecture du fichier"""
        try:
            df = pd.read_excel(file_path, sheet_name=0)
            df.columns = df.columns.astype(str)
            df.columns = self._clean_column_names(df.columns)
            df = self._remove_duplicate_columns(df)
            
            if self.container_column is None:
                self.container_column = self._find_container_column(df)
            
            return df, list(df.columns)
            
        except Exception as e:
            logger.error(f"Erreur lecture Excel {file_path}: {str(e)}")
            raise
    
    def _clean_column_names(self, columns):
        cleaned_columns = []
        seen_columns = set()
        
        for col in columns:
            clean_col = str(col).strip().upper()
            
            if clean_col in ['REEL NO.', 'NO BOBINE', 'NUMERO BOBINE', 'REEL_NO', 'N° BOBINE']:
                clean_col = 'NO_BOBINE'
            elif clean_col in ['CONTAINER', 'CONTENEUR', 'CTN']:
                clean_col = 'CONTENEUR'
            elif clean_col in ['DIAM MM', 'DIAMÈTRE', 'DIAMETRE', 'DIAM', 'DIAMÉTRE']:
                clean_col = 'DIAMETRE'
            elif clean_col in ['MT', 'POIDS', 'WEIGHT', 'POIDS (KG)']:
                clean_col = 'POIDS'
            elif clean_col in ['PRODUCT', 'PRODUIT']:
                clean_col = 'PRODUIT'
            elif clean_col in ['REF PAPIER', 'REFERENCE PAPIER', 'REF_PAPIER', 'RÉFÉRENCE']:
                clean_col = 'REF_PAPIER'
            elif clean_col in ['METRAGE', 'LONGUEUR']:
                clean_col = 'METRAGE'
            
            if clean_col in seen_columns:
                counter = 1
                new_col = f"{clean_col}_{counter}"
                while new_col in seen_columns:
                    counter += 1
                    new_col = f"{clean_col}_{counter}"
                clean_col = new_col
            
            seen_columns.add(clean_col)
            cleaned_columns.append(clean_col)
        
        return cleaned_columns
    
    def _remove_duplicate_columns(self, df):
        columns_to_drop = []
        
        for i, col1 in enumerate(df.columns):
            for j, col2 in enumerate(df.columns):
                if i < j and df[col1].equals(df[col2]):
                    columns_to_drop.append(col2)
        
        columns_to_drop = list(set(columns_to_drop))
        
        if columns_to_drop:
            return df.drop(columns=columns_to_drop)
        
        return df
    
    def _find_container_column(self, df):
        """Trouve la colonne conteneur"""
        container_keywords = ['conteneur', 'container', 'ctn', 'cnt']
        
        for col in df.columns:
            col_lower = str(col).lower()
            if any(keyword in col_lower for keyword in container_keywords):
                return col
        
        return df.columns[0] if len(df.columns) > 0 else None
    
    def extract_containers(self, df):
        """Extrait les conteneurs uniques"""
        if self.container_column is None or self.container_column not in df.columns:
            return []
        
        containers = df[self.container_column].dropna().unique()
        return [str(container).strip() for container in containers if str(container).strip()]
    
    def filter_by_container(self, df, container):
        """Filtre les données pour un conteneur spécifique"""
        if self.container_column not in df.columns:
            return df
        
        return df[df[self.container_column] == container]
    
    def create_excel(self, data, container, output_dir,
                     cariste, fournisseur, numero_dossier,
                     type_certification, numero_certificat):
        """Crée un fichier Excel avec codes-barres"""
        try:
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            filename = f"SEMBA_RECEPTION_{container}_{timestamp}.xlsx"
            file_path = os.path.join(output_dir, filename)
            
            # Créer un nouveau workbook
            workbook = openpyxl.Workbook()
            sheet = workbook.active
            sheet.title = "FO57"
            
            # Création du template avec codes-barres
            self._create_template_with_barcodes(sheet, data, container, output_dir,
                                              cariste, fournisseur, numero_dossier,
                                              type_certification, numero_certificat)
            
            workbook.save(file_path)
            self._cleanup_temp_images(output_dir)
            
            logger.info(f"Fichier Excel créé: {file_path}")
            return file_path
            
        except Exception as e:
            logger.error(f"Erreur création Excel {container}: {str(e)}")
            raise

    def create_excel_from_zzz_template(self, data, container, output_dir, zzz_file_path,
                                      cariste, fournisseur, numero_dossier,
                                      type_certification, numero_certificat):
        """Crée un fichier Excel en utilisant le template ZZZZ complet"""
        try:
            return self.create_excel(
                data=data,
                container=container,
                output_dir=output_dir,
                cariste=cariste,
                fournisseur=fournisseur,
                numero_dossier=numero_dossier,
                type_certification=type_certification,
                numero_certificat=numero_certificat
            )
            
        except Exception as e:
            logger.error(f"Erreur création Excel depuis template {container}: {str(e)}")
            raise
    
    def _create_template_with_barcodes(self, sheet, data, container, output_dir,
                                     cariste, fournisseur, numero_dossier,
                                     type_certification, numero_certificat):
        """Crée le template SEMBA avec codes-barres"""
        
        # Styles
        title_font = Font(size=16, bold=True, name='Arial')
        header_font = Font(size=10, bold=True, name='Arial')
        normal_font = Font(size=10, name='Arial')
        small_font = Font(size=8, name='Arial')
        company_font = Font(size=12, bold=True, name='Arial')
        
        thin_border = Border(left=Side(style='thin'), right=Side(style='thin'),
                           top=Side(style='thin'), bottom=Side(style='thin'))
        header_fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
        
        # 1. LOGO SEMBA
        logo_path = self._get_logo_path()
        if logo_path and os.path.exists(logo_path):
            try:
                logo = Image(logo_path)
                logo.width = 120
                logo.height = 60
                sheet.add_image(logo, 'A1')
                sheet['D1'] = "SEMBA"
                sheet['D1'].font = company_font
                sheet['D1'].alignment = Alignment(vertical='center')
                sheet.row_dimensions[1].height = 45
            except Exception as e:
                logger.warning(f"Impossible d'ajouter le logo: {e}")
                sheet['A1'] = "SEMBA"
                sheet['A1'].font = company_font
        
        # 2. Titre principal
        sheet['C3'] = "FICHE DE CONTROLE RECEPTION DES BOBINES"
        sheet['C3'].font = title_font
        sheet['C3'].alignment = Alignment(horizontal='center')
        sheet.merge_cells('C3:H3')
        
        sheet['I2'] = f"Date de création : {datetime.now().strftime('%d/%m/%Y')}"
        sheet['I2'].font = small_font
        sheet['I2'].alignment = Alignment(horizontal='right')
        
        # 4. Informations générales
        info_row = 5
        
        sheet[f'A{info_row}'] = f"CARISTE : {cariste}"
        sheet[f'A{info_row}'].font = Font(size=12, bold=True)
        
        sheet[f'C{info_row}'] = f"N° CT : {container}"
        sheet[f'C{info_row}'].font = Font(size=12, bold=True)
        
        sheet[f'E{info_row}'] = f"DATE : {datetime.now().strftime('%d/%m/%Y')}"
        sheet[f'E{info_row}'].font = Font(size=12, bold=True)
        
        sheet[f'G{info_row}'] = f"No. Dossier : {numero_dossier}"
        sheet[f'G{info_row}'].font = Font(size=12, bold=True)
        sheet[f'G{info_row}'].alignment = Alignment(horizontal='right')
        
        # 5. En-têtes du tableau
        headers = [
            'N°', 'N° Fournisseur', 'Fournisseur', 'Code barre N° bobine/N° SEMBA', 
            'Référence', 'Diamètre', 'Poids (KG)', 'N° Certificat FSC', 
            'Type de certification FSC', 'Observation MAGASIN par rapport au papier reçu'
        ]
        
        start_row = info_row + 1
        
        for col, header in enumerate(headers, 1):
            cell = sheet.cell(row=start_row, column=col)
            cell.value = header
            cell.font = header_font
            cell.border = thin_border
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        
        # 6. Données des bobines avec codes-barres
        barcode_count = 0
        
        for idx, (_, row) in enumerate(data.iterrows(), 1):
            row_num = start_row + idx
            
            no_bobine = str(row.get('NO_BOBINE', '')) if 'NO_BOBINE' in row else f"Bobine_{idx}"
            
            # Colonne 1: N°
            sheet.cell(row=row_num, column=1, value=idx)
            
            # Colonne 2: N° Fournisseur
            sheet.cell(row=row_num, column=2, value=no_bobine)
            
            # Colonne 3: Fournisseur
            sheet.cell(row=row_num, column=3, value=fournisseur)
            
            # Colonne 4: Code barre
            barcode_success = self._insert_barcode(sheet, no_bobine, output_dir, idx, row_num)
            
            if barcode_success:
                barcode_count += 1
            else:
                # Fallback : afficher le numéro
                sheet.cell(row=row_num, column=4, value=f"Code: {no_bobine}")
            
            # Colonnes 5-10
            sheet.cell(row=row_num, column=5, value=str(row.get('REF_PAPIER', '')) if 'REF_PAPIER' in row else "")
            sheet.cell(row=row_num, column=6, value=row.get('DIAMETRE', '') if 'DIAMETRE' in row else '')
            sheet.cell(row=row_num, column=7, value=row.get('POIDS', '') if 'POIDS' in row else '')
            sheet.cell(row=row_num, column=8, value=numero_certificat)
            sheet.cell(row=row_num, column=9, value=type_certification)
            sheet.cell(row=row_num, column=10, value="")
            
            # Appliquer les styles
            for col in range(1, 11):
                if col != 4:
                    cell = sheet.cell(row=row_num, column=col)
                    cell.font = normal_font
                    cell.border = thin_border
                    cell.alignment = Alignment(vertical='center', horizontal='center')
        
        logger.info(f"Codes-barres générés avec succès: {barcode_count}/{len(data)}")
        
        # 7. Ajustements des dimensions
        column_widths = [5, 15, 12, 25, 15, 10, 12, 18, 20, 30]
        for col, width in enumerate(column_widths, 1):
            sheet.column_dimensions[openpyxl.utils.get_column_letter(col)].width = width
        
        sheet.column_dimensions['D'].width = 25
        sheet.row_dimensions[3].height = 25
        sheet.row_dimensions[start_row].height = 35
        
        for idx in range(1, len(data) + 1):
            sheet.row_dimensions[start_row + idx].height = 45
    
    def _insert_barcode(self, sheet, code_value, output_dir, index, row_num):
        """Insère un code-barres dans la feuille Excel"""
        try:
            clean_code = str(code_value).strip()
            if not clean_code:
                clean_code = f"BT{index:04d}"
            
            # Générer le code-barres
            barcode_path = self._generate_barcode(clean_code, output_dir, index)
            
            if barcode_path and os.path.exists(barcode_path):
                file_size = os.path.getsize(barcode_path)
                if file_size > 100:  # Vérifier que le fichier n'est pas vide
                    img = Image(barcode_path)
                    img.width = 180
                    img.height = 50
                    sheet.add_image(img, f'D{row_num}')
                    logger.info(f"Code-barres généré: {clean_code}")
                    return True
                else:
                    logger.warning(f"Fichier code-barres vide: {clean_code}")
                    return False
            else:
                logger.warning(f"Code-barres non généré: {clean_code}")
                return False
                
        except Exception as e:
            logger.error(f"Erreur code-barres {code_value}: {e}")
            return False
    
    def _generate_barcode(self, code_value, output_dir, index):
        """Génère un code-barres Code128"""
        try:
            # Créer le répertoire pour les codes-barres
            barcode_dir = os.path.join(output_dir, 'temp_barcodes')
            os.makedirs(barcode_dir, exist_ok=True)
            
            # Utiliser Code128
            code128 = barcode.get_barcode_class('code128')
            
            # Configuration optimale
            options = {
                'write_text': False,      # Pas de texte sous le code-barres
                'quiet_zone': 2.0,       # Zone tranquille
                'module_height': 15.0,   # Hauteur des barres
                'module_width': 0.2,     # Largeur des barres
                'font_size': 0,          # Taille police 0
                'text_distance': 0,      # Distance texte 0
                'background': 'white',   # Fond blanc
                'foreground': 'black',   # Barres noires
            }
            
            # Générer le code-barres
            barcode_obj = code128(code_value, writer=ImageWriter())
            
            # Sauvegarder dans le répertoire temporaire
            safe_name = re.sub(r'[^\w]', '_', code_value)
            filename = f"barcode_{safe_name}_{index}"
            full_path = os.path.join(barcode_dir, filename)
            
            # Sauvegarder l'image
            barcode_obj.save(full_path, options=options)
            
            # Retourner le chemin complet avec extension
            return f"{full_path}.png"
            
        except Exception as e:
            logger.error(f"Erreur génération code-barres {code_value}: {e}")
            return None
    
    def _get_logo_path(self):
        """Retourne le chemin vers le logo SEMBA"""
        base_dir = getattr(settings, 'BASE_DIR', None)
        if base_dir:
            static_paths = [
                os.path.join(base_dir, 'static', 'images', 'semba_logo.jpg'),
                os.path.join(base_dir, 'static', 'images', 'semba_logo.png'),
            ]
            
            for path in static_paths:
                if os.path.exists(path):
                    return path
        return None
    
    def _cleanup_temp_images(self, output_dir):
        """Nettoie les fichiers temporaires"""
        try:
            barcode_dir = os.path.join(output_dir, 'temp_barcodes')
            if os.path.exists(barcode_dir):
                for file in os.listdir(barcode_dir):
                    file_path = os.path.join(barcode_dir, file)
                    if os.path.isfile(file_path):
                        os.remove(file_path)
                os.rmdir(barcode_dir)
        except Exception as e:
            logger.warning(f"Erreur nettoyage images: {e}")