import pandas as pd
from tqdm import tqdm
from openpyxl import load_workbook

def read_excel_with_progress(file_path, header=None):
    wb = load_workbook(filename=file_path, read_only=True)
    ws = wb.active
    total_rows = ws.max_row
    wb.close()

    with tqdm(total=total_rows, desc=f"Chargement du fichier '{file_path}'") as pbar:
        data = pd.read_excel(file_path, header=header)
        pbar.update(total_rows)
    return data

print("Initialisation du script...")

people = read_excel_with_progress('people.xlsx')
custom = read_excel_with_progress('custom.xlsx')
departements_c3 = read_excel_with_progress('departements.xlsx', header=None)

print("Colonne de 'people':", people.columns.tolist())
print("Colonne de 'custom':", custom.columns.tolist())

people_copy = people.copy()
custom_copy = custom.copy()

# Vérifier que les noms des colonnes correspondent
custom_copy.rename(columns={'GGI': 'IGG', 'Email': 'GROUP_MAIL', 'Department': 'LIB_SERVICE'}, inplace=True)

print("Colonnes après renommage dans 'custom':", custom_copy.columns.tolist())

merged_data = pd.merge(people_copy, custom_copy[['IGG', 'GROUP_MAIL', 'LIB_SERVICE']], on='IGG', how='left', suffixes=('', '_custom'))

merged_data['GROUP_MAIL'] = merged_data['GROUP_MAIL'].fillna(merged_data['GROUP_MAIL_custom'])
merged_data['LIB_SERVICE'] = merged_data['LIB_SERVICE'].fillna(merged_data['LIB_SERVICE_custom'])

merged_data.drop(columns=['GROUP_MAIL_custom', 'LIB_SERVICE_custom'], inplace=True)

c3_departments = set(departements_c3.iloc[:, 0])
filtered_data = merged_data[merged_data['LIB_SERVICE'].isin(c3_departments)]

filtered_data['LIB_SERVICE'] = filtered_data['LIB_SERVICE'].fillna(method='bfill')

final_data = filtered_data[['IGG', 'GROUP_MAIL', 'LIB_SERVICE']]

final_data.to_excel('C3_accredited_users.xlsx', index=False)

print("Le fichier 'C3_accredited_users.xlsx' a été créé avec succès. Le processus est terminé.")
