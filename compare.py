import pandas as pd

# Charger les fichiers Excel
your_output = pd.read_excel('C3_accredited_users.xlsx')
colleague_output = pd.read_excel('colleague_output.xlsx', header=None, names=['GROUP_MAIL'])  # Lire le fichier de votre collègue sans en-tête

# Extraire les mails
your_emails = set(your_output['GROUP_MAIL'].dropna())
colleague_emails = set(colleague_output['GROUP_MAIL'].dropna())

# Trouver les différences et les mails communs
emails_in_your_not_in_colleague = your_emails - colleague_emails
emails_in_colleague_not_in_your = colleague_emails - your_emails
common_emails = your_emails & colleague_emails

# Créer des DataFrames pour les résultats
df_emails_in_your_not_in_colleague = pd.DataFrame(list(emails_in_your_not_in_colleague), columns=["Clement"])
df_emails_in_colleague_not_in_your = pd.DataFrame(list(emails_in_colleague_not_in_your), columns=["Jerome"])
df_common_emails = pd.DataFrame(list(common_emails), columns=['Commun'])

# Enregistrer les résultats dans des fichiers Excel
with pd.ExcelWriter('comparison_results.xlsx') as writer:
    df_emails_in_your_not_in_colleague.to_excel(writer, sheet_name='Clement', index=False)
    df_emails_in_colleague_not_in_your.to_excel(writer, sheet_name='Jerome', index=False)
    df_common_emails.to_excel(writer, sheet_name='Commun', index=False)

print("La comparaison des mails est terminée. Les résultats ont été enregistrés dans 'comparison_results.xlsx'.")
