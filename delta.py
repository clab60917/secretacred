import pandas as pd

# Charger les fichiers Excel
old_c3 = pd.read_excel('jde.xlsx')
new_c3 = pd.read_excel('C3_accredited_users.xlsx')

# Extraire les emails
old_c3_emails = set(old_c3['GROUP_MAIL'].dropna().str.lower())
new_c3_emails = set(new_c3['GROUP_MAIL'].dropna().str.lower())

# Identifier les nouvelles personnes accréditées C3
newly_accredited_c3 = new_c3_emails - old_c3_emails

# Filtrer les nouvelles personnes accréditées C3 dans le nouveau fichier
new_c3_accredited_df = new_c3[new_c3['GROUP_MAIL'].str.lower().isin(newly_accredited_c3)]

# Sauvegarder le fichier avec les nouvelles personnes accréditées C3
new_c3_accredited_df.to_excel('new_c3_accredited_users.xlsx', index=False)

print("Le fichier 'new_c3_accredited_users.xlsx' a été créé avec succès.")
