import pandas as pd
from tqdm import tqdm
from openpyxl import load_workbook

def read_excel_with_progress(file_path, header='infer'):
    # Charger le workbook avec openpyxl pour lire le nombre de lignes
    wb = load_workbook(filename=file_path, read_only=True)
    ws = wb.active
    total_rows = ws.max_row
    
    # Lire les noms des colonnes depuis la première ligne si header est 0
    if header == 0:
        columns = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]
        data_start_row = 2
    else:
        columns = [f'Column_{i+1}' for i in range(len(next(ws.iter_rows(min_row=1, max_row=1))))]
        data_start_row = 1

    data = {column: [] for column in columns}
    
    # Lire les données avec progression
    pbar = tqdm(total=total_rows, desc=f"Lecture de {file_path}")
    for row in ws.iter_rows(min_row=data_start_row, max_row=ws.max_row):
        for key, cell in zip(columns, row):
            data[key].append(cell.value)
        pbar.update(1)
    pbar.close()

    wb.close()
    
    # Créer un DataFrame à partir du dictionnaire
    return pd.DataFrame(data)

print("Initialisation du script...")

# Lire les fichiers Excel avec barres de progression
print("\n---------------\nLecture des fichiers Excel...\n---------------")
people = read_excel_with_progress('people.xlsx', header=0)
custom = read_excel_with_progress('custom.xlsx', header=0)
departements_c3 = read_excel_with_progress('departements.xlsx', header=None)

# Vérification des colonnes des DataFrames
print("\n---------------\nVérification des colonnes des DataFrames...\n---------------")
print("Colonnes de 'people':", people.columns.tolist())
print("Colonnes de 'custom':", custom.columns.tolist())

# Créer des copies des DataFrames pour les manipulations
print("\n---------------\nCréation de copies des DataFrames pour manipulation...\n---------------")
people_copy = people.copy()
custom_copy = custom.copy()

# Renommer les colonnes dans 'custom_copy'
print("\n---------------\nRenommage des colonnes dans 'custom' pour correspondre à celles de 'people'...\n---------------")
if 'GGI' in custom_copy.columns and 'Email' in custom_copy.columns and 'Department' in custom_copy.columns:
    custom_copy.rename(columns={'GGI': 'IGG', 'Email': 'GROUP_MAIL', 'Department': 'LIB_SERVICE'}, inplace=True)
    print("Colonnes après renommage dans 'custom':", custom_copy.columns.tolist())
else:
    raise KeyError("Les colonnes attendues 'GGI', 'Email' et 'Department' ne sont pas présentes dans 'custom'.")

# Fusionner people_copy avec custom_copy pour compléter les informations manquantes
print("\n---------------\nFusion des DataFrames...\n---------------")
merged_data = pd.merge(people_copy, custom_copy[['IGG', 'GROUP_MAIL', 'LIB_SERVICE']], on='IGG', how='left', suffixes=('', '_custom'))

# Utiliser fillna pour combler les informations manquantes
print("\n---------------\nComblement des informations manquantes...\n---------------")
merged_data['GROUP_MAIL'] = merged_data['GROUP_MAIL'].fillna(merged_data['GROUP_MAIL_custom'])
merged_data['LIB_SERVICE'] = merged_data['LIB_SERVICE'].fillna(merged_data['LIB_SERVICE_custom'])

# Supprimer les colonnes inutiles après la fusion
print("\n---------------\nSuppression des colonnes temporaires après la fusion...\n---------------")
merged_data.drop(columns=['GROUP_MAIL_custom', 'LIB_SERVICE_custom'], inplace=True)

# Filtrer les départements C3
print("\n---------------\nFiltration des utilisateurs selon les départements C3...\n---------------")
c3_departments = set(departements_c3.iloc[:, 0])
filtered_data = merged_data[merged_data['LIB_SERVICE'].isin(c3_departments)].copy()

# Assurer que la colonne LIB_SERVICE est toujours remplie
print("\n---------------\nAssurer que la colonne 'LIB_SERVICE' est toujours remplie...\n---------------")
filtered_data.loc[:, 'LIB_SERVICE'] = filtered_data['LIB_SERVICE'].ffill()

# Sélectionner les colonnes nécessaires pour le fichier final
print("\n---------------\nSélection des colonnes finales pour l'export...\n---------------")
final_data = filtered_data[['IGG', 'GROUP_MAIL', 'LIB_SERVICE']]

# Sauvegarder le fichier final
print("\n---------------\nSauvegarde du fichier final 'C3_accredited_users.xlsx'...\n---------------")
final_data.to_excel('C3_accredited_users.xlsx', index=False)

print("\n---------------\nLe fichier 'C3_accredited_users.xlsx' a été créé avec succès. Le processus est terminé.\n---------------")
