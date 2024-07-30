import pandas as pd
from colorama import init, Fore, Style

# Initialisation de colorama
init(autoreset=True)

# Chargement des données de référence et de l'output du script
synchro_data = pd.read_excel('synchronized_people_custom.xlsx')
departements_data = pd.read_excel('departements.xlsx', sheet_name=None)
output_data = pd.read_excel('C3_accredited_users.xlsx')

def print_result(success, message):
    if success:
        print(Fore.GREEN + Style.BRIGHT + "Test réussi: " + message)
    else:
        print(Fore.RED + Style.BRIGHT + "Échec du test: " + message)

# Test 1: Vérification des utilisateurs des départements IGAD et DFIN
def test_departments_igad_dfin():
    print("\n" + Style.BRIGHT + "Test 1: Vérification des utilisateurs des départements IGAD et DFIN")
    igad_dfin_users = synchro_data[
        synchro_data['LIB_SERVICE'].str.contains('IGAD|DFIN', na=False)
    ]
    mismatches = igad_dfin_users[~igad_dfin_users['GROUP_MAIL'].isin(output_data['GROUP_MAIL'])]
    if mismatches.empty:
        print_result(True, "Tous les utilisateurs des départements IGAD et DFIN sont inclus dans la liste des utilisateurs C3.")
    else:
        print_result(False, f"Les utilisateurs suivants des départements IGAD et DFIN ne sont pas inclus dans la liste des utilisateurs C3:\n{mismatches}")

# Test 2: Comparaison entre synchronized_people_custom.xlsx et l'onglet "C3 dpt only"
def test_c3_dpt_only():
    print("\n" + Style.BRIGHT + "Test 2: Comparaison entre synchronized_people_custom.xlsx et l'onglet 'C3 dpt only'")
    c3_dpt_only = departements_data['LIST C3 DPT ONLY INTERNALS']
    c3_dpt_only_departments = set(c3_dpt_only.iloc[5:, 0].apply(lambda x: str(x).split('.')[0].strip()))

    potential_c3_users = synchro_data[
        synchro_data['LIB_SERVICE'].apply(lambda x: any(dept in str(x) for dept in c3_dpt_only_departments))
    ]
    potential_c3_users.to_excel('Potential_C3_Users_Comparison.xlsx', index=False)
    print_result(True, "Le fichier 'Potential_C3_Users_Comparison.xlsx' a été généré pour comparaison manuelle.")

# Test 3: Vérification des utilisateurs nominatifs
def test_nominative_users():
    print("\n" + Style.BRIGHT + "Test 3: Vérification des utilisateurs nominatifs")
    nominative_users = departements_data['NOMINATIVE USERS']
    nominative_emails = set(nominative_users['Mail'].dropna())

    missing_users = nominative_emails - set(output_data['GROUP_MAIL'])
    if not missing_users:
        print_result(True, "Tous les utilisateurs nominatifs sont inclus dans la liste des utilisateurs C3.")
    else:
        print_result(False, f"Les utilisateurs nominatifs suivants ne sont pas inclus dans la liste des utilisateurs C3:\n{missing_users}")

# Test 4: Comparaison avec la liste des utilisateurs à exclure
def test_excluded_users():
    print("\n" + Style.BRIGHT + "Test 4: Comparaison avec la liste des utilisateurs à exclure")
    excluded_sheet = departements_data['DPTS-USER TO BE EXCLUDED']
    email_column = excluded_sheet.columns[4]  # Colonne E est l'index 4 (0-indexed)

    excluded_users = set(excluded_sheet[email_column].dropna())
    remaining_excluded_users = excluded_users & set(output_data['GROUP_MAIL'])
    
    if not remaining_excluded_users:
        print_result(True, "Aucun utilisateur devant être exclu n'est présent dans la liste finale.")
    else:
        print_result(False, f"Les utilisateurs suivants, qui devraient être exclus, sont encore présents dans la liste finale:\n{remaining_excluded_users}")

# Test 5: Création d'une liste de noms à vérifier
def test_users_to_verify():
    print("\n" + Style.BRIGHT + "Test 5: Création d'une liste de noms à vérifier")
    uncertain_users = []  # Liste à remplir manuellement ou par des critères spécifiques

    # Placeholder pour une méthode d'identification automatique ou manuelle
    # Ajouter des critères pour identifier les utilisateurs à vérifier ou utiliser une liste pré-établie

    # Export de la liste des noms à vérifier
    pd.DataFrame(uncertain_users, columns=['GROUP_MAIL']).to_excel('Users_to_Verify.xlsx', index=False)
    print_result(True, "Le fichier 'Users_to_Verify.xlsx' a été généré pour vérification manuelle.")

# Exécution des tests
test_departments_igad_dfin()
test_c3_dpt_only()
test_nominative_users()
test_excluded_users()
test_users_to_verify()
