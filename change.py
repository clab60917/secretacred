import pandas as pd

def detect_changes(old_file, new_file, key_column, column_to_check):
    # Charger les anciens et nouveaux fichiers
    old_data = pd.read_excel(old_file)
    new_data = pd.read_excel(new_file)
    
    # Fusionner les deux fichiers sur la colonne clé
    merged_data = pd.merge(old_data, new_data, on=key_column, suffixes=('_old', '_new'))
    
    # Détecter les changements
    changes = merged_data[merged_data[f'{column_to_check}_old'] != merged_data[f'{column_to_check}_new']]
    
    return changes[[key_column, f'{column_to_check}_old', f'{column_to_check}_new']]

def main():
    # Fichiers à comparer
    old_people_file = 'people_old.xlsx'
    new_people_file = 'people_new.xlsx'
    old_syncro_file = 'synchronized_people_custom_old.xlsx'
    new_syncro_file = 'synchronized_people_custom_new.xlsx'
    
    # Colonne à vérifier
    column_to_check = 'LIB_SERVICE'
    
    # Détection des changements dans le fichier people
    print("Détection des changements dans 'people'...")
    people_changes = detect_changes(old_people_file, new_people_file, 'IGG', column_to_check)
    people_changes.to_excel('people_changes.xlsx', index=False)
    print("Les changements dans 'people' ont été enregistrés dans 'people_changes.xlsx'.")
    
    # Détection des changements dans le fichier syncro
    print("Détection des changements dans 'syncro'...")
    syncro_changes = detect_changes(old_syncro_file, new_syncro_file, 'IGG', column_to_check)
    syncro_changes.to_excel('syncro_changes.xlsx', index=False)
    print("Les changements dans 'syncro' ont été enregistrés dans 'syncro_changes.xlsx'.")

if __name__ == "__main__":
    main()
