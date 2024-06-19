import pandas as pd
from tqdm import tqdm, tqdm_notebook

print("Initialisation du script...")

# Fonction pour lire les fichiers Excel avec une barre de progression
def read_excel_with_progress(filename):
    print(f"Chargement du fichier '{filename}'...")
    iterator = pd.read_excel(filename, chunksize=500)  # Réglage de chunksize selon la taille du fichier
    df = pd.concat([chunk for chunk in tqdm(iterator, total=100, desc=f"Lecture de {filename}")])  # Estimer 'total' selon la taille/division attendue
    return df

# Lire les fichiers Excel avec barres de progression
people = read_excel_with_progress('people.xlsx')
custom = read_excel_with_progress('custom.xlsx')
departements_c3 = read_excel_with_progress('departements.xlsx', header=None)

# Créer des copies des DataFrames pour les manipulations
print("Création de copies des DataFrames...")
people_copy = people.copy()
custom_copy = custom.copy()

# Renommer les colonnes dans 'custom_copy' pour correspondre à celles de 'people_copy'
print("Renommage des colonnes dans le DataFrame 'custom'...")
custom_copy.rename(columns={'GGI': 'IGG', 'Email': 'GROUP_MAIL', 'Department': 'LIB_SERVICE'}, inplace=True)

# Fusionner people_copy avec custom_copy pour compléter les informations manquantes
print("Fusion des DataFrames...")
merged_data = pd.merge(people_copy, custom_copy[['IGG', 'GROUP_MAIL', 'LIB_SERVICE']], on='IGG', how='left', suffixes=('', '_custom'))

# Utiliser fillna avec une barre de progression pour combler les informations manquantes
print("Comblement des informations manquantes dans 'GROUP_MAIL' et 'LIB_SERVICE'...")
merged_data['GROUP_MAIL'] = merged_data['GROUP_MAIL'].progress_apply(lambda x: x.fillna(merged_data['GROUP_MAIL_custom']))
merged_data['LIB_SERVICE'] = merged_data['LIB_SERVICE'].progress_apply(lambda x: x.fillna(merged_data['LIB_SERVICE_custom']))

# Supprimer les colonnes inutiles après la fusion
print("Suppression des colonnes temporaires après la fusion...")
merged_data.drop(columns=['GROUP_MAIL_custom', 'LIB_SERVICE_custom'], inplace=True)

# Filtrer les départements C3 avec une barre de progression
print("Filtration des utilisateurs selon leur département...")
c3_departments = set(departements_c3[0])
filtered_data = merged_data[merged_data['LIB_SERVICE'].isin(c3_departments)]

# Affichage de la progression pour la filtration
for _ in tqdm(range(filtered_data.shape[0]), desc="Filtrage des départements C3"):
    pass

# Assurer que la colonne LIB_SERVICE est toujours remplie
print("Assurance que la colonne 'LIB_SERVICE' est toujours remplie...")
filtered_data['LIB_SERVICE'] = filtered_data['LIB_SERVICE'].fillna(method='bfill')

# Sélectionner les colonnes nécessaires pour le fichier final
print("Sélection des colonnes finales pour l'export...")
final_data = filtered_data[['IGG', 'GROUP_MAIL', 'LIB_SERVICE']]

# Sauvegarder le fichier final
print("Sauvegarde du fichier final 'C3_accredited_users.xlsx'...")
final_data.to_excel('C3_accredited_users.xlsx', index=False)

print("Le fichier 'C3_accredited_users.xlsx' a été créé avec succès. Le processus est terminé.")
