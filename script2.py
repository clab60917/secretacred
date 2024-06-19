import pandas as pd
from tqdm import tqdm
from tqdm.auto import tqdm as auto_tqdm

# Initialiser tqdm pour pandas
tqdm.pandas()

def read_excel_with_progress(file_path):
    # Charger le DataFrame avec une barre de progression pour chaque chunk de données lu
    iter_csv = pd.read_excel(file_path, chunksize=1000)  # Ajustez le chunksize selon la mémoire disponible
    df = pd.concat([chunk.progress_apply(lambda x: x) for chunk in tqdm(iter_csv, desc=f"Lecture de {file_path}")], ignore_index=True)
    return df

# Lire les fichiers Excel avec barre de chargement
people = read_excel_with_progress('people.xlsx')
custom = read_excel_with_progress('custom.xlsx')
departements_c3 = read_excel_with_progress('departements.xlsx', header=None)

# Créer des copies des DataFrames pour les manipulations
people_copy = people.copy()
custom_copy = custom.copy()

# Renommer les colonnes dans 'custom_copy' pour correspondre à celles de 'people_copy'
custom_copy.rename(columns={'GGI': 'IGG', 'Email': 'GROUP_MAIL', 'Department': 'LIB_SERVICE'}, inplace=True)

# Fusionner people_copy avec custom_copy pour compléter les informations manquantes
merged_data = pd.merge(people_copy, custom_copy[['IGG', 'GROUP_MAIL', 'LIB_SERVICE']], on='IGG', how='left', suffixes=('', '_custom')).progress_apply(lambda x: x, desc="Fusion des données")

# Utiliser fillna pour combler les informations manquantes
merged_data['GROUP_MAIL'] = merged_data['GROUP_MAIL'].fillna(merged_data['GROUP_MAIL_custom'])
merged_data['LIB_SERVICE'] = merged_data['LIB_SERVICE'].fillna(merged_data['LIB_SERVICE_custom'])

# Supprimer les colonnes inutiles après la fusion
merged_data.drop(columns=['GROUP_MAIL_custom', 'LIB_SERVICE_custom'], inplace=True)

# Filtrer les départements C3 avec une barre de progression
c3_departments = set(departements_c3[0])
filtered_data = merged_data[merged_data['LIB_SERVICE'].isin(c3_departments)].progress_apply(lambda x: x, desc="Filtrage des départements C3")

# Sélectionner les colonnes nécessaires pour le fichier final
final_data = filtered_data[['IGG', 'GROUP_MAIL', 'LIB_SERVICE']]

# Sauvegarder le fichier final
final_data.to_excel('C3_accredited_users.xlsx', index=False)

print("Le fichier 'C3_accredited_users.xlsx' a été créé avec succès.")
