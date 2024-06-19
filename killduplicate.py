import pandas as pd

# Lire le fichier Excel de sortie
output_file = 'C3_accredited_users.xlsx'
data = pd.read_excel(output_file)

# Éliminer les doublons en ne gardant qu'une seule occurrence de chaque IGG
unique_data = data.drop_duplicates(subset='IGG')

# Sauvegarder le fichier nettoyé
cleaned_output_file = 'C3_accredited_users_cleaned.xlsx'
unique_data.to_excel(cleaned_output_file, index=False)

print(f"Le fichier '{cleaned_output_file}' a été créé avec succès, sans doublons.")
