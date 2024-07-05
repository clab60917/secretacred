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
df_emails_in_your_not_in_colleague = pd.DataFrame(list(emails_in_your_not_in_colleague), columns=["Emails in your output but not in colleague's"])
df_emails_in_colleague_not_in_your = pd.DataFrame(list(emails_in_colleague_not_in_your), columns=["Emails in colleague's output but not in yours"])
df_common_emails = pd.DataFrame(list(common_emails), columns=['Common Emails'])

# Enregistrer les résultats dans des fichiers Excel
with pd.ExcelWriter('comparison_results.xlsx') as writer:
    df_emails_in_your_not_in_colleague.to_excel(writer, sheet_name='In your not in colleague', index=False)
    df_emails_in_colleague_not_in_your.to_excel(writer, sheet_name='In colleague not in your', index=False)
    df_common_emails.to_excel(writer, sheet_name='Common Emails', index=False)

print("La comparaison des mails est terminée. Les résultats ont été enregistrés dans 'comparison_results.xlsx'.")
