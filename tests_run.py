import pandas as pd
from colorama import init, Fore, Style

# Initialisation de colorama
init(autoreset=True)

# Chargement des données de référence et de l'output du script
synchro_data = pd.read_excel('synchronized_people_custom.xlsx')
departements_data = pd.read_excel('departements.xlsx', sheet_name=None)
output_data = pd.read_excel('C3_accredited_users.xlsx')
collegue_data = pd.read_excel('collegue.xlsx', header=None, names=['GROUP_MAIL'])

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
    uncertain_users = []

    # Critère 1: Utilisateurs avec des domaines d'e-mail externes
    external_emails = output_data[~output_data['GROUP_MAIL'].str.endswith('@votreentreprise.com')]
    uncertain_users.extend(external_emails['GROUP_MAIL'].tolist())

    # Critère 2: Utilisateurs avec des titres de poste spécifiques (ajustez selon vos besoins)
    specific_titles = synchro_data[synchro_data['CONTRACT_GROUP_TYPE_LABEL'].str.contains('Trainee|Apprenticeship|Extern-temp|International Business Volunte', na=False)]
    uncertain_users.extend(specific_titles['GROUP_MAIL'].tolist())

    # Supprimer les doublons
    uncertain_users = list(set(uncertain_users))

    # Export de la liste des noms à vérifier
    pd.DataFrame(uncertain_users, columns=['GROUP_MAIL']).to_excel('Users_to_Verify.xlsx', index=False)
    print_result(True, "Le fichier 'Users_to_Verify.xlsx' a été généré pour vérification manuelle.")

# Test 6: Compter les utilisateurs IGAD et DFIN
def test_count_igad_dfin():
    print("\n" + Style.BRIGHT + "Test 6: Compter les utilisateurs IGAD et DFIN")
    igad_users = synchro_data[synchro_data['LIB_SERVICE'].str.contains('IGAD', na=False)]
    dfin_users = synchro_data[synchro_data['LIB_SERVICE'].str.contains('DFIN', na=False)]
    
    print(Fore.CYAN + f"Nombre d'utilisateurs dans IGAD: {len(igad_users)}")
    print(Fore.CYAN + f"Nombre d'utilisateurs dans DFIN: {len(dfin_users)}")

# Test 7: Comparaison des listes C3 potentielles
def test_compare_potential_c3():
    print("\n" + Style.BRIGHT + "Test 7: Comparaison des listes C3 potentielles")
    # Chargement des mails des utilisateurs potentiellement C3
    potential_c3_users = pd.read_excel('Potential_C3_Users_Comparison.xlsx')['GROUP_MAIL']
    collegue_c3_users = set(collegue_data['GROUP_MAIL'].dropna())
    script_output_c3_users = set(output_data['GROUP_MAIL'].dropna())

    # Utilisateurs communs entre les trois listes
    common_users = set(potential_c3_users).intersection(collegue_c3_users, script_output_c3_users)
    # Utilisateurs spécifiques à chaque source
    only_in_potential = set(potential_c3_users) - (collegue_c3_users | script_output_c3_users)
    only_in_collegue = collegue_c3_users - (set(potential_c3_users) | script_output_c3_users)
    only_in_script = script_output_c3_users - (set(potential_c3_users) | collegue_c3_users)

    # Affichage des résultats
    print(Fore.CYAN + f"Utilisateurs communs: {len(common_users)}")
    print(Fore.YELLOW + f"Uniquement dans Potential_C3_Users_Comparison.xlsx: {len(only_in_potential)}")
    print(Fore.YELLOW + f"Uniquement dans collegue.xlsx: {len(only_in_collegue)}")
    print(Fore.YELLOW + f"Uniquement dans C3_accredited_users.xlsx: {len(only_in_script)}")

    # Export des résultats pour vérification manuelle
    pd.DataFrame({'Common Users': list(common_users)}).to_excel('Common_C3_Users.xlsx', index=False)
    pd.DataFrame({'Only in Potential': list(only_in_potential)}).to_excel('Only_in_Potential_C3_Users.xlsx', index=False)
    pd.DataFrame({'Only in Collegue': list(only_in_collegue)}).to_excel('Only_in_Collegue_C3_Users.xlsx', index=False)
    pd.DataFrame({'Only in Script Output': list(only_in_script)}).to_excel('Only_in_Script_Output_C3_Users.xlsx', index=False)

    print_result(True, "Les fichiers de comparaison des listes C3 ont été générés.")

# Test 8: Vérification des utilisateurs de JDE.xlsx dans l'output final
def test_users_in_jde():
    print("\n" + Style.BRIGHT + "Test 8: Vérification des utilisateurs de JDE.xlsx dans l'output final")
    jde_data = pd.read_excel('JDE.xlsx')
    test_emails = set(jde_data['GROUP_MAIL'].dropna())

    # Vérification de la présence dans l'output
    present_users = test_emails & set(output_data['GROUP_MAIL'])
    absent_users = test_emails - set(output_data['GROUP_MAIL'])

    print(Fore.CYAN + f"Utilisateurs trouvés dans l'output: {len(present_users)}")
    print(Fore.RED + f"Utilisateurs manquants dans l'output: {len(absent_users)}")
    
    if absent_users:
        print_result(False, f"Les utilisateurs suivants de JDE.xlsx ne sont pas présents dans la liste des utilisateurs C3:\n{absent_users}")
    else:
        print_result(True, "Tous les utilisateurs de JDE.xlsx sont présents dans la liste des utilisateurs C3.")

# Exécution des tests
test_departments_igad_dfin()
test_c3_dpt_only()
test_nominative_users()
test_excluded_users()
test_users_to_verify()
test_count_igad_dfin()
test_compare_potential_c3()
test_users_in_jde()
