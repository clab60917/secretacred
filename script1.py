""" Étapes pour automatisation des users C3 :
1. Lire les fichiers people et custom en utilisant pandas.
2. Synchroniser les données basées sur les colonnes IGG dans people et GGI dans custom pour combler les manques en GROUP_MAIL et Email.
3. Ajouter et synchroniser le département si manquant dans people.
4. Lire le fichier departements pour obtenir la liste des départements C3.
5. Identifier et extraire les utilisateurs C3 en se basant sur leur appartenance aux départements listés dans le fichier departements.
6. Générer un nouveau fichier Excel contenant les colonnes IGG, GROUP_MAIL, et LIB_SERVICE pour les utilisateurs C3.
"""
import pandas as pd

# Lire les fichiers Excel
people = pd.read_excel('people.xlsx')
custom = pd.read_excel('custom.xlsx')
departements_c3 = pd.read_excel('departements.xlsx', header=None)

# Renommer les colonnes pour la correspondance
custom.rename(columns={'GGI': 'IGG', 'Email': 'GROUP_MAIL'}, inplace=True)

# Fusionner people avec custom pour compléter les informations manquantes
people_updated = pd.merge(people, custom[['IGG', 'GROUP_MAIL', 'Department']], on='IGG', how='left')
people['GROUP_MAIL'] = people['GROUP_MAIL'].fillna(people_updated['GROUP_MAIL'])
people['Department'] = people['Department'].fillna(people_updated['Department'])

# Filtrer les départements C3
c3_departments = set(departements_c3[0])
people_c3 = people[people['Department'].isin(c3_departments)]

# Sélectionner les colonnes nécessaires pour le fichier final
final_data = people_c3[['IGG', 'GROUP_MAIL', 'Department', 'LIB_SERVICE']]

# Sauvegarder le fichier final
final_data.to_excel('C3_accredited_users.xlsx', index=False)

print("Le fichier 'C3_accredited_users.xlsx' a été créé avec succès.")
