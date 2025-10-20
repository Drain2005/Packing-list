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
from pathlib import Path
import shutil

logger = logging.getLogger(__name__)

class ExcelProcessor:
    def __init__(self):
        self.container_column = None

    def read_excel_file(self, file_path):
        """Lit un fichier Excel et nettoie les noms de colonnes."""
        try:
            df = pd.read_excel(file_path, sheet_name=0)
            df.columns = df.columns.astype(str)
            df.columns = self._clean_column_names(df.columns)
            df = self._remove_duplicate_columns(df)
            self.container_column = self._find_container_column(df)
            return df, list(df.columns)
        except Exception as e:
            logger.error(f"Erreur lors de la lecture du fichier Excel : {e}")
            raise

    def _clean_column_names(self, columns):
        """Nettoie et uniformise les noms de colonnes."""
        cleaned_columns = []
        seen_columns = set()
        mapping = {
            'REEL NO.': 'NO_BOBINE', 'NO BOBINE': 'NO_BOBINE', 'NUMERO BOBINE': 'NO_BOBINE',
            'REEL_NO': 'NO_BOBINE', 'NÂ° BOBINE': 'NO_BOBINE', 'CONTAINER': 'CONTENEUR',
            'CONTENEUR': 'CONTENEUR', 'CTN': 'CONTENEUR', 'DIAM MM': 'DIAMETRE',
            'DIAMÈTRE': 'DIAMETRE', 'DIAMETRE': 'DIAMETRE', 'POIDS (KG)': 'POIDS',
            'POIDS': 'POIDS', 'PRODUCT': 'PRODUIT', 'PRODUIT': 'PRODUIT',
            'REF PAPIER': 'REF_PAPIER', 'REFERENCE PAPIER': 'REF_PAPIER',
            'RÉFÉRENCE': 'REF_PAPIER', 'METRAGE': 'METRAGE', 'LONGUEUR': 'METRAGE'
        }

        for col in columns:
            col_str = str(col).strip().upper() if col else ""
            clean_col = mapping.get(col_str, col_str)
            if clean_col in seen_columns:
                i = 1
                while f"{clean_col}_{i}" in seen_columns:
                    i += 1
                clean_col = f"{clean_col}_{i}"
            seen_columns.add(clean_col)
            cleaned_columns.append(clean_col)
        return cleaned_columns

    def _remove_duplicate_columns(self, df):
        """Supprime les colonnes dupliquées."""
        return df.loc[:, ~df.columns.duplicated()]

    def _find_container_column(self, df):
        """Trouve la colonne contenant les conteneurs."""
        for col in df.columns:
            if col in ['CONTENEUR', 'CONTAINER', 'CTN']:
                return col
        return None

    def create_excel(self, data, container, output_dir, cariste, fournisseur,
                     numero_dossier, type_certification, numero_certificat, template_path=None):
        """Crée un fichier Excel basé sur le template pour un conteneur."""
        try:
            required_columns = ['NO_BOBINE', 'REF_PAPIER', 'DIAMETRE', 'POIDS']
            missing_cols = [col for col in required_columns if col not in data.columns]
            if missing_cols:
                logger.warning(f"Colonnes manquantes : {missing_cols}")

            Path(output_dir).mkdir(parents=True, exist_ok=True)
            filename = f"{container}.xlsx"
            file_path = os.path.join(output_dir, filename)

            # Charger le template
            if template_path and os.path.exists(template_path):
                shutil.copy2(template_path, file_path)
                workbook = openpyxl.load_workbook(file_path)
                if 'Mode de remplisage' in workbook.sheetnames:
                    logger.info("✓ Feuille 'Mode de remplisage' supprimée")
                    del workbook['Mode de remplisage']
            else:
                workbook = openpyxl.Workbook()
                logger.warning("Template non trouvé, création d'un nouveau classeur")

            sheet = workbook['FO57'] if 'FO57' in workbook.sheetnames else workbook.active
            sheet.title = "FO57"

            # Remplir les données
            self._fill_template_complete(
                sheet, data, container, cariste, fournisseur, numero_dossier,
                type_certification, numero_certificat, output_dir
            )

            workbook.save(file_path)
            self._cleanup_temp_images(output_dir)
            logger.info(f"✓ Fichier Excel créé : {file_path}")
            return file_path

        except Exception as e:
            logger.error(f"Erreur lors de la création de l'Excel pour {container}: {e}")
            raise

    def _fill_template_complete(self, sheet, data, container, cariste, fournisseur, 
                                numero_dossier, type_certification, numero_certificat, output_dir):
        """Remplit complètement le template avec les données."""
        try:
            # Chercher la ligne du tableau (en-têtes)
            table_start_row = self._find_table_start_row(sheet)
            
            # Remplir les en-têtes du document (CARISTE, N° CT, DATE, No. Dossier)
            self._fill_document_headers(sheet, cariste, numero_dossier)
            
            # Remplir les données du tableau
            self._fill_table_data(
                sheet, table_start_row, data, cariste, fournisseur, 
                numero_certificat, type_certification, output_dir
            )
            
            logger.debug("Template rempli complètement")
        except Exception as e:
            logger.error(f"Erreur lors du remplissage du template : {e}")
            raise

    def _find_table_start_row(self, sheet):
        """Trouve la première ligne de données du tableau (après les en-têtes)."""
        for row in sheet.iter_rows(min_row=1, max_row=20, min_col=1, max_col=10):
            for cell in row:
                if cell.value and isinstance(cell.value, str):
                    if 'TABLEAU' in str(cell.value).upper() or str(cell.value).strip() == 'N°':
                        logger.debug(f"Ligne de tableau trouvée à la ligne {cell.row}")
                        return cell.row + 1
        # Par défaut, commencer après la ligne 6
        logger.debug("Ligne de tableau par défaut : 7")
        return 7

    def _fill_document_headers(self, sheet, cariste, numero_dossier):
        """Remplit les champs d'en-tête du document (CARISTE, N° DOSSIER, DATE, etc.)."""
        try:
            current_date = datetime.now().strftime("%d/%m/%Y")
            
            # Parcourir les premières lignes pour trouver et remplir les champs
            for row in sheet.iter_rows(min_row=1, max_row=6):
                for cell in row:
                    if not cell.value:
                        continue
                    
                    cell_value_upper = str(cell.value).upper().strip()
                    
                    if 'CARISTE' in cell_value_upper:
                        # Remplir la cellule suivante avec le cariste
                        next_col = cell.column + 1
                        self._set_cell_value_safe(sheet, cell.row, next_col, cariste)
                        logger.debug(f"Cariste rempli : {cariste}")
                    
                    elif 'N° CT' in cell_value_upper or 'N°CT' in cell_value_upper:
                        next_col = cell.column + 1
                        self._set_cell_value_safe(sheet, cell.row, next_col, "-")
                    
                    elif 'DATE' in cell_value_upper and 'DOSSIER' not in cell_value_upper:
                        next_col = cell.column + 1
                        self._set_cell_value_safe(sheet, cell.row, next_col, current_date)
                        logger.debug(f"Date remplie : {current_date}")
                    
                    elif 'DOSSIER' in cell_value_upper or 'N°. DOSSIER' in cell_value_upper:
                        next_col = cell.column + 1
                        self._set_cell_value_safe(sheet, cell.row, next_col, numero_dossier)
                        logger.debug(f"Numéro dossier rempli : {numero_dossier}")
        except Exception as e:
            logger.error(f"Erreur lors du remplissage des en-têtes du document : {e}")

    def _fill_table_data(self, sheet, start_row, data, cariste, fournisseur, 
                        numero_certificat, type_certification, output_dir):
        """Remplit le tableau des données."""
        try:
            current_row = start_row
            
            for idx, (_, row_data) in enumerate(data.iterrows()):
                # Remplir chaque colonne
                no_bobine = self._safe_get_value(row_data, 'NO_BOBINE', '')
                ref_papier = self._safe_get_value(row_data, 'REF_PAPIER', '')
                diametre = self._safe_get_value(row_data, 'DIAMETRE', '')
                poids = self._safe_get_value(row_data, 'POIDS', '')

                # Col 1: Numéro
                self._set_cell_value_safe(sheet, current_row, 1, idx + 1)
                
                # Col 2: Code-barres (inséré après)
                # Col 3: N° bobine
                self._set_cell_value_safe(sheet, current_row, 3, no_bobine)
                
                # Col 4: N° bobine entre parenthèses (format template)
                self._set_cell_value_safe(sheet, current_row, 4, f"({no_bobine})")
                
                # Col 5: Fournisseur
                self._set_cell_value_safe(sheet, current_row, 5, fournisseur)
                
                # Col 6: Référence papier
                self._set_cell_value_safe(sheet, current_row, 6, ref_papier)
                
                # Col 7: Diamètre
                self._set_cell_value_safe(sheet, current_row, 7, diametre)
                
                # Col 8: Poids
                self._set_cell_value_safe(sheet, current_row, 8, poids)
                
                # Col 9: N° Certificat
                self._set_cell_value_safe(sheet, current_row, 9, numero_certificat)
                
                # Col 10: Type certification
                self._set_cell_value_safe(sheet, current_row, 10, type_certification)
                
                # Col 11: Observation
                self._set_cell_value_safe(sheet, current_row, 11, "")
                
                # Insérer le code-barres en colonne 2
                self._insert_barcode_in_template(sheet, no_bobine, output_dir, idx, current_row, 2)
                
                current_row += 1
                
            logger.info(f"✓ {len(data)} lignes de tableau remplies")
        except Exception as e:
            logger.error(f"Erreur lors du remplissage du tableau : {e}")
            raise

    def _safe_get_value(self, row_data, column, default=''):
        """Récupère une valeur de manière sécurisée."""
        try:
            if column not in row_data.index:
                return default
            value = row_data[column]
            if pd.isna(value):
                return default
            if isinstance(value, (tuple, list)):
                logger.warning(f"Valeur tuple/liste détectée dans {column}: {value}")
                return str(value)
            if isinstance(value, (int, float)) and column in ['DIAMETRE', 'POIDS']:
                return float(value)
            return str(value).strip()
        except Exception as e:
            logger.error(f"Erreur lors de la récupération de {column}: {e}")
            return default

    def _set_cell_value_safe(self, sheet, row, col, value):
        """Définit la valeur d'une cellule en gérant les cellules fusionnées."""
        try:
            cell = sheet.cell(row=row, column=col)
            
            # Vérifier si c'est une cellule fusionnée
            if cell.coordinate in sheet.merged_cells:
                logger.debug(f"Cellule fusionnée ignorée : {cell.coordinate}")
                return
            
            if isinstance(value, (tuple, list)):
                logger.warning(f"Valeur tuple/liste à ({row},{col}): {value}")
                value = str(value)
            
            cell.value = value
        except Exception as e:
            logger.error(f"Erreur à la cellule ({row},{col}): {e}")

    def _insert_barcode_in_template(self, sheet, barcode_value, output_dir, idx, row, col):
        """Insère un code-barres dans le template."""
        try:
            if not barcode_value or not isinstance(barcode_value, str):
                logger.error(f"Code-barres invalide à la ligne {row}: {barcode_value}")
                return

            barcode_value = ''.join(c for c in barcode_value if c.isalnum())
            if not barcode_value:
                logger.warning(f"Code-barres vide après nettoyage à la ligne {row}")
                return

            temp_dir = os.path.join(output_dir, 'temp_barcodes')
            Path(temp_dir).mkdir(parents=True, exist_ok=True)
            
            # save() ajoute automatiquement .png
            barcode_filename = f"barcode_{idx}_{row}"
            barcode_path_without_ext = os.path.join(temp_dir, barcode_filename)
            
            logger.debug(f"Génération code-barres : {barcode_value}")
            code128 = barcode.get('code128', barcode_value, writer=ImageWriter())
            code128.save(barcode_path_without_ext)
            
            barcode_path = f"{barcode_path_without_ext}.png"
            if not os.path.exists(barcode_path):
                logger.warning(f"Fichier code-barres non créé : {barcode_path}")
                return

            img = Image(barcode_path)
            img.width, img.height = 100, 30
            cell_address = f"{chr(64 + col)}{row}"
            sheet.add_image(img, cell_address)
            logger.info(f"✓ Code-barres inséré : {cell_address}")
        except Exception as e:
            logger.warning(f"Erreur code-barres à ({row},{col}): {e}")

    def _cleanup_temp_images(self, output_dir):
        """Nettoie les images temporaires."""
        temp_dir = os.path.join(output_dir, 'temp_barcodes')
        try:
            if os.path.exists(temp_dir):
                shutil.rmtree(temp_dir)
                logger.info(f"✓ Dossier temporaire supprimé")
        except Exception as e:
            logger.warning(f"Erreur nettoyage dossier temporaire : {e}")

    def extract_containers(self, data):
        """Extrait la liste des conteneurs uniques."""
        if self.container_column and self.container_column in data.columns:
            containers = data[self.container_column].dropna().unique().tolist()
            return [str(c) for c in containers if c]
        return []

    def filter_by_container(self, data, container):
        """Filtre les données par conteneur."""
        if self.container_column and self.container_column in data.columns:
            return data[data[self.container_column] == container]
        return pd.DataFrame()