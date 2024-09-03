import os
from pdfminer.pdfparser import PDFParser
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdfinterp import resolve1
import pandas as pd

def resolve_object(obj, max_depth=10):
    """
    Fonction récursive pour résoudre les objets PDF.
    """
    if max_depth <= 0:
        return None

    if isinstance(obj, dict):
        resolved_dict = {}
        for key, value in obj.items():
            resolved_key = resolve_object(key, max_depth - 1)
            resolved_value = resolve_object(value, max_depth - 1)
            resolved_dict[resolved_key] = resolved_value
        return resolved_dict
    elif isinstance(obj, list):
        return [resolve_object(item, max_depth - 1) for item in obj]
    elif isinstance(obj, bytes):
        return obj.decode('utf-8', errors='ignore')
    else:
        return obj

def extract_form_fields(pdf_file):
    """
    Extraction des champs de formulaire d'un fichier PDF.
    """
    fields = {}
    # Ouvrir le fichier PDF
    with open(pdf_file, 'rb') as file:
        parser = PDFParser(file)
        doc = PDFDocument(parser)
        acro_form = resolve1(doc.catalog['AcroForm'])

        # Si la référence est valide, on fait la résolution
        if acro_form:
            fields_ref = resolve1(acro_form.get('Fields'))

            # Si les champs de formulaire sont trouvés, on fait l'itération
            if fields_ref:
                for field_ref in fields_ref:
                    field = resolve1(field_ref)
                    field_name = resolve1(field.get('T'))
                    field_value = resolve1(field.get('V'))
                    if field_value is not None:  # s'assurer si c'est pas vide
                        if isinstance(field_value, bytes):  # Vérifie si la valeur est encodée en bytes
                            field_value = field_value.decode('utf-8', errors='ignore')  # Décode la valeur
                            # Enlever le préfixe 'b' et le suffixe "'" des valeurs
                            if field_value.startswith("b'") and field_value.endswith("'"):
                                field_value = field_value[2:-1]
                        fields[field_name] = field_value

    return fields

import re

def clean_string(value):
    """
    Nettoyage de la chaîne de caractères en éliminant les caractères non imprimables.
    """
    return re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\xff]', '', value)

def export_to_excel(data, excel_file):
    """
    Exportation des données vers un fichier Excel.
    """
    # On applique la fonction pour un clean de data
    cleaned_data = [{k: clean_string(v) if isinstance(v, str) else v for k, v in entry.items()} for entry in data]
    
    df = pd.DataFrame(cleaned_data)
    df.to_excel(excel_file, index=False)



def process_folder(folder_path):
    """
    Fonction pour traiter tous les fichiers PDF dans un dossier.
    """
    data = []
    # Parcourir tous les fichiers dans le dossier
    for filename in os.listdir(folder_path):
        if filename.endswith(".pdf"):
            pdf_path = os.path.join(folder_path, filename)
            form_fields = extract_form_fields(pdf_path)
            data.append(form_fields)
    return data

# Dossier contenant les fichiers PDF
pdf_folder = "./"

# Extraire les données de tous les fichiers PDF dans le dossier
data = process_folder(pdf_folder)

# Export des données vers un fichier Excel
export_to_excel(data, "donnees_formulaire.xlsx")
