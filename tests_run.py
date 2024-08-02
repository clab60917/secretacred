import pandas as pd
from colorama import Fore, Style, init

# Initialiser colorama pour les couleurs dans le terminal
init(autoreset=True)

# Chemins des fichiers
c3_output_file = 'C3_accredited_users.xlsx'
colleague_output_file = 'colleague_output.xlsx'
jde_file = 'JDE.xlsx'
comparison_file = 'comparison_data.xlsx'

# Charger les fichiers Excel nécessaires
c3_data = pd.read_excel(c3_output_file)
colleague_data = pd.read_excel(colleague_output_file, header=None, names=['GROUP_MAIL'])
jde_data = pd.read_excel(jde_file)
comparison_data = pd.read_excel(comparison_file, sheet_name=None)

# Extraire les onglets 'jerome' et 'clement'
jerome_data = comparison_data['jerome']
clement_data = comparison_data['clement']

# Tests
def test_nominative_users():
    print(Fore.BLUE + "\n--- Test des utilisateurs nominatifs ---")
    nominative_mails = set(jde_data['GROUP_MAIL'].dropna())
    c3_mails = set(c3_data['GROUP_MAIL'].dropna())

    missing_mails = nominative_mails - c3_mails

    if not missing_mails:
        print(Fore.GREEN + "[SUCCÈS] Tous les utilisateurs nominatifs sont inclus.")
    else:
        print(Fore.RED + "[ÉCHEC] Les utilisateurs nominatifs suivants ne sont pas inclus :")
        for mail in missing_mails:
            print(Fore.RED + mail)

def test_excluded_users():
    print(Fore.BLUE + "\n--- Test des utilisateurs à exclure ---")
    excluded_emails = set(comparison_data['DPTS-USER TO BE EXCLUDED']['Mail'].dropna())
    
    c3_emails = set(c3_data['GROUP_MAIL'].dropna())

    included_excluded_mails = c3_emails & excluded_emails

    if not included_excluded_mails:
        print(Fore.GREEN + "[SUCCÈS] Aucun utilisateur exclu ne figure dans la liste C3.")
    else:
        print(Fore.RED + "[ÉCHEC] Les utilisateurs suivants, qui devaient être exclus, sont toujours présents :")
        for mail in included_excluded_mails:
            print(Fore.RED + mail)

def test_comparison_with_colleague():
    print(Fore.BLUE + "\n--- Test de comparaison avec la liste du collègue ---")
    your_emails = set(c3_data['GROUP_MAIL'].dropna())
    colleague_emails = set(colleague_data['GROUP_MAIL'].dropna())

    emails_in_your_not_in_colleague = your_emails - colleague_emails
    emails_in_colleague_not_in_your = colleague_emails - your_emails
    common_emails = your_emails & colleague_emails

    if not emails_in_your_not_in_colleague:
        print(Fore.GREEN + "[SUCCÈS] Tous les utilisateurs de votre liste se trouvent également dans celle du collègue.")
    else:
        print(Fore.RED + "[ÉCHEC] Les utilisateurs suivants sont dans votre liste mais pas dans celle du collègue :")
        for email in emails_in_your_not_in_colleague:
            print(Fore.RED + email)

    if not emails_in_colleague_not_in_your:
        print(Fore.GREEN + "[SUCCÈS] Tous les utilisateurs de la liste du collègue se trouvent également dans la vôtre.")
    else:
        print(Fore.RED + "[ÉCHEC] Les utilisateurs suivants sont dans la liste du collègue mais pas dans la vôtre :")
        for email in emails_in_colleague_not_in_your:
            print(Fore.RED + email)

def test_comparison_with_provided_lists():
    print(Fore.BLUE + "\n--- Test de comparaison avec les listes fournies ---")
    c3_mails_depts = c3_data[['GROUP_MAIL', 'DEPARTEMENT']].dropna()
    c3_mails_depts_set = set(tuple(x) for x in c3_mails_depts.to_numpy())

    jerome_missing = set(tuple(x) for x in jerome_data.to_numpy())
    missing_from_c3 = jerome_missing - c3_mails_depts_set

    clement_extra = set(tuple(x) for x in clement_data.to_numpy())
    expected_in_c3 = clement_extra & c3_mails_depts_set

    def print_comparison_result(label, missing_from_c3, expected_in_c3):
        if missing_from_c3:
            print(Fore.RED + f"[ÉCHEC] {label} - Utilisateurs manquants dans le C3:")
            for email, dept in missing_from_c3:
                print(Fore.RED + f"Email: {email}, Département: {dept}")
        else:
            print(Fore.GREEN + f"[SUCCÈS] {label} - Tous les utilisateurs sont correctement inclus.")

        if expected_in_c3:
            print(Fore.YELLOW + f"[ATTENTION] {label} - Utilisateurs trouvés mais inattendus dans le C3:")
            for email, dept in expected_in_c3:
                print(Fore.YELLOW + f"Email: {email}, Département: {dept}")

    print_comparison_result("Comparaison avec les utilisateurs manquants (Jerome)", missing_from_c3, expected_in_c3)
    print_comparison_result("Comparaison avec les utilisateurs inattendus (Clement)", clement_extra - c3_mails_depts_set, None)

def main():
    test_nominative_users()
    test_excluded_users()
    test_comparison_with_colleague()
    test_comparison_with_provided_lists()

if __name__ == "__main__":
    main()
