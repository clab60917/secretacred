import pandas as pd
from colorama import Fore, Style, init

# Initialiser colorama pour les couleurs dans le terminal
init(autoreset=True)

# Chemins des fichiers
c3_output_file = 'C3_accredited_users.xlsx'
jde_file = 'JDE.xlsx'
jerome_clement_file = 'comparison_data.xlsx'  # Le fichier contenant les onglets 'jerome' et 'clement'

# Charger les fichiers Excel nécessaires
c3_data = pd.read_excel(c3_output_file)
jde_data = pd.read_excel(jde_file)
jerome_clement_data = pd.read_excel(jerome_clement_file, sheet_name=None)

# Extraire les onglets 'jerome' et 'clement'
jerome_data = jerome_clement_data['jerome']
clement_data = jerome_clement_data['clement']

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

def test_comparison_with_jerome_clement():
    print(Fore.BLUE + "\n--- Test de comparaison avec les listes 'jerome' et 'clement' ---")
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
    test_comparison_with_jerome_clement()

if __name__ == "__main__":
    main()
