import os
import csv
import re
import time
import signal
import pythoncom
from win32com.client import Dispatch

# Chemin vers le dossier contenant les fichiers Word
folder_path = r'C:\Users\PC-Mahdi\Desktop\Adresse'

# Fichier CSV de sortie
output_csv = 'extracted_names.csv'

# Expression régulière pour correspondre au format Monsieur le Docteur Nom Prénom
pattern = re.compile(r'(?:D(?:octeur)?|docteur)\s+([A-Za-zÀ-ÖØ-öø-ÿ\s]+)')

# Fonction pour extraire les noms et prénoms d'un document .doc
def extract_name(doc_path):
    pythoncom.CoInitialize()
    word = Dispatch('Word.Application')
    full_text = []
    try:
        doc = word.Documents.Open(doc_path)
        for paragraph in doc.Paragraphs:
            full_text.append(paragraph.Range.Text)
        doc.Close(False)
    except Exception as e:
        print(f"Error processing file {doc_path}: {e}")
    finally:
        word.Quit()
    for paragraph in full_text:
        match = pattern.search(paragraph)
        if match:
            return match.group(1).strip()
    return None

# Gestion du signal d'interruption (Ctrl+C)
def signal_handler(signal, frame):
    print("\nExtraction interrompue par l'utilisateur.")
    generate_csv()
    exit(0)

# Fonction pour générer le fichier CSV
def generate_csv():
    with open(output_csv, 'w', newline='', encoding='utf-8') as csvfile:
        csvwriter = csv.writer(csvfile)
        csvwriter.writerow(['Nom et Prénom'])
        csvwriter.writerows(extracted_names)

# Enregistrement du gestionnaire de signal
signal.signal(signal.SIGINT, signal_handler)

# Parcourir tous les fichiers Word dans le dossier et extraire les noms et prénoms
extracted_names = []
for filename in os.listdir(folder_path):
    if filename.endswith('.doc'):
        file_path = os.path.join(folder_path, filename)
        print(f"Processing file: {file_path}")
        name = extract_name(file_path)
        if name:
            extracted_names.append([name])
        else:
            print(f"No match found in file: {file_path}")

# Appel de la fonction pour générer le fichier CSV
generate_csv()

print(f'Extraction terminée. Les noms et prénoms ont été enregistrés dans {output_csv}')
