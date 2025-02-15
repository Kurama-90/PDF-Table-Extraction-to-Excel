#Kuram-90 ( https://github.com/Kurama-90 )

import fitz  # PyMuPDF
import pandas as pd  # Pour créer un fichier Excel
import os

# Chemins des dossiers
UPLOADS_DIR = 'uploads'
OUTPUTS_DIR = 'outputs'

def extract_tables_from_pdf(pdf_path):
    # Ouvrir le PDF
    pdf_document = fitz.open(pdf_path)
    
    # Liste pour stocker les tables extraites
    tables = []

    # Parcourir chaque page du PDF
    for page_num in range(len(pdf_document)):
        print(f"Traitement de la page {page_num + 1}...")
        page = pdf_document[page_num]
        
        # Extraire les tables avec leurs coordonnées
        tables_in_page = page.find_tables()  # Détecter les tables dans la page
        
        for table in tables_in_page:
            # Extraire les données de la table
            table_data = table.extract()
            tables.append(table_data)  # Ajouter la table à la liste
    
    # Fermer le PDF
    pdf_document.close()
    
    return tables

def export_tables_to_excel(tables, output_excel_path):
    # Créer un fichier Excel avec plusieurs feuilles (une par table)
    with pd.ExcelWriter(output_excel_path, engine='xlsxwriter') as writer:
        for i, table in enumerate(tables):
            # Convertir la table en DataFrame
            df = pd.DataFrame(table[1:], columns=table[0])  # La première ligne est l'en-tête
            # Exporter la table dans une feuille Excel
            df.to_excel(writer, sheet_name=f"Table_{i + 1}", index=False)
    
    print(f"Fichier Excel généré : {output_excel_path}")

def process_pdf(pdf_path, output_excel_path):
    # Extraire les tables du PDF
    tables = extract_tables_from_pdf(pdf_path)
    
    # Exporter les tables dans un fichier Excel
    export_tables_to_excel(tables, output_excel_path)

if __name__ == '__main__':
    # Demander le nom du fichier PDF
    pdf_name = input("Veuillez entrer le nom du fichier PDF (avec l'extension .pdf) : ")
    
    # Chemin du fichier PDF dans le dossier uploads
    pdf_file = os.path.join(UPLOADS_DIR, pdf_name)
    
    # Vérifier si le fichier existe
    if not os.path.exists(pdf_file):
        print(f"Erreur : Le fichier {pdf_file} n'existe pas dans le dossier 'uploads'.")
        exit()
    
    # Chemin du fichier Excel de sortie dans le dossier outputs
    output_excel = os.path.join(OUTPUTS_DIR, 'output.xlsx')
    
    # Vérifier si le dossier outputs existe, sinon le créer
    if not os.path.exists(OUTPUTS_DIR):
        print(f"Création du dossier {OUTPUTS_DIR}...")
        os.makedirs(OUTPUTS_DIR)
    
    # Traiter le PDF et générer le fichier Excel
    process_pdf(pdf_file, output_excel)
    print(f"Le fichier Excel a été généré avec succès.")