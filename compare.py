import pandas as pd

# Charger les fichiers Excel
votre_output = pd.read_excel('C3_accredited_users_cleaned.xlsx')
collegue_output = pd.read_excel('collegue.xlsx')

# Extraire les mails des utilisateurs C3 de chaque fichier
vos_mails = set(votre_output['GROUP_MAIL'])
collegue_mails = set(collegue_output.iloc[:, 1])  # Colonne B correspond à l'index 1

# Trouver les différences
mails_uniques_vos = vos_mails - collegue_mails
mails_uniques_collegue = collegue_mails - vos_mails
mails_communs = vos_mails & collegue_mails

# Créer des DataFrames pour les résultats
unique_vos_df = votre_output[votre_output['GROUP_MAIL'].isin(mails_uniques_vos)]
unique_collegue_df = collegue_output[collegue_output.iloc[:, 1].isin(mails_uniques_collegue)]
communs_df = votre_output[votre_output['GROUP_MAIL'].isin(mails_communs)]

# Sauvegarder les résultats dans un fichier Excel avec des feuilles séparées
with pd.ExcelWriter('comparison_C3_users.xlsx') as writer:
    unique_vos_df.to_excel(writer, sheet_name='Vos utilisateurs uniques', index=False)
    unique_collegue_df.to_excel(writer, sheet_name='Utilisateurs collègues uniques', index=False)
    communs_df.to_excel(writer, sheet_name='Utilisateurs communs', index=False)

print("\n---------------\nLe script a été exécuté avec succès. Les différences ont été sauvegardées dans 'comparison_C3_users.xlsx'.\n---------------")
