""" Étapes pour automatisation des users C3 :
1. Lire les fichiers people et custom en utilisant pandas.
2. Synchroniser les données basées sur les colonnes IGG dans people et GGI dans custom pour combler les manques en GROUP_MAIL et Email.
3. Ajouter et synchroniser le département si manquant dans people.
4. Lire le fichier departements pour obtenir la liste des départements C3.
5. Identifier et extraire les utilisateurs C3 en se basant sur leur appartenance aux départements listés dans le fichier departements.
6. Générer un nouveau fichier Excel contenant les colonnes IGG, GROUP_MAIL, et LIB_SERVICE pour les utilisateurs C3.
"""
import pandas as pd
from tqdm import tqdm

# Initialiser tqdm pour pandas
tqdm.pandas()

# Lire les fichiers Excel
people = pd.read_excel('people.xlsx')
custom = pd.read_excel('custom.xlsx')
departements_c3 = pd.read_excel('departements.xlsx', header=None)

# Créer des copies des DataFrames pour les manipulations
people_copy = people.copy()
custom_copy = custom.copy()

# Renommer les colonnes dans 'custom_copy' pour correspondre à celles de 'people_copy'
custom_copy.rename(columns={'GGI': 'IGG', 'Email': 'GROUP_MAIL', 'Department': 'LIB_SERVICE'}, inplace=True)

# Fusionner people_copy avec custom_copy pour compléter les informations manquantes
merged_data = pd.merge(people_copy, custom_copy[['IGG', 'GROUP_MAIL', 'LIB_SERVICE']], on='IGG', how='left', suffixes=('', '_custom'))

# Utiliser fillna pour combler les informations manquantes
merged_data['GROUP_MAIL'] = merged_data['GROUP_MAIL'].fillna(merged_data['GROUP_MAIL_custom'])
merged_data['LIB_SERVICE'] = merged_data['LIB_SERVICE'].fillna(merged_data['LIB_SERVICE_custom'])

# Supprimer les colonnes inutiles après la fusion
merged_data.drop(columns=['GROUP_MAIL_custom', 'LIB_SERVICE_custom'], inplace=True)

# Filtrer les départements C3 avec une barre de progression
c3_departments = set(departements_c3[0])
filtered_data = merged_data[merged_data['LIB_SERVICE'].isin(c3_departments)].progress_apply(lambda x: x)

# Assurer que la colonne LIB_SERVICE est toujours remplie
filtered_data['LIB_SERVICE'] = filtered_data['LIB_SERVICE'].fillna(method='bfill')

# Sélectionner les colonnes nécessaires pour le fichier final
final_data = filtered_data[['IGG', 'GROUP_MAIL', 'LIB_SERVICE']]

# Sauvegarder le fichier final
final_data.to_excel('C3_accredited_users.xlsx', index=False)

print("Le fichier 'C3_accredited_users.xlsx' a été créé avec succès.")
