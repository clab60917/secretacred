# README - Projet de Vérification des Utilisateurs C3 - secretacred

## Objectif
Ce projet a pour objectif d'identifier et de vérifier les utilisateurs accrédités C3 en se basant sur plusieurs sources de données. Les scripts fournis permettent de synchroniser les données, d'appliquer des filtres spécifiques pour extraire les utilisateurs C3, et de vérifier les résultats en les comparant à des listes fournies.

## Contenu des Scripts

### 1. `scriptV18.py`
**Fonction :** Ce script principal identifie les utilisateurs C3 en se basant sur plusieurs fichiers d'entrée. Il synchronise les données entre différents fichiers, applique des filtres pour exclure certains utilisateurs, et génère un fichier final `C3_accredited_users.xlsx` contenant la liste des utilisateurs C3.

**Étapes principales :**
1. **Synchronisation des données :** Fusion des fichiers `people.xlsx` et `custom.xlsx` pour créer `synchronized_people_custom.xlsx`.
2. **Filtrage selon `LIST C3 DPT ONLY INTERNALS` :** Identification des utilisateurs C3 selon les départements.
3. **Filtrage selon `LIST OF ELR` :** Identification des utilisateurs C3 selon les habilitations ELR.
4. **Application des règles d'exclusion :** Exclusion des utilisateurs selon des critères spécifiques (ex. contrats temporaires, emails externes).
5. **Filtrage selon `LIST C3 DPT ALL TYPES` :** Identification des utilisateurs supplémentaires selon tous types de départements.
6. **Filtrage des utilisateurs nominatifs :** Ajout des utilisateurs dont l'email est dans `NOMINATIVE USERS`.
7. **Filtrage final :** Exclusion des utilisateurs absents et suppression des doublons avant de générer le fichier `C3_accredited_users.xlsx`.

### 2. `tests_run.py`
**Fonction :** Ce script exécute une série de tests pour vérifier l'exactitude des résultats générés par `scriptV18.py`. Il compare le fichier `C3_accredited_users.xlsx` avec diverses listes de référence et génère des rapports de test.

**Tests inclus :**
- **Comparaison avec la liste de mon collègue :** Vérifie les différences entre les utilisateurs C3 identifiés par le script et ceux identifiés par un collègue.
- **Comparaison avec les listes 'jerome' et 'clement' :** Compare les résultats du script avec deux listes fournies par 'jerome' et 'clement', en vérifiant les utilisateurs manquants ou en trop.
- **Vérification des utilisateurs dans `JDE.xlsx` :** Assure que les utilisateurs de `JDE.xlsx` sont tous inclus dans le fichier final.

## Instructions d'Utilisation

### Prérequis
- **Python 3.x** installé.
- Installer les bibliothèques nécessaires :
  ```
  pip install pandas openpyxl alive-progress
  ```

### Exécution des Scripts

1. **Script Principal (`scriptV18.py`)**
   - Placez les fichiers `people.xlsx`, `custom.xlsx`, et `departements.xlsx` dans le même dossier que le script.
   - Exécutez le script :
     ```
     python scriptV18.py
     ```
   - Le fichier final `C3_accredited_users.xlsx` sera généré dans le même dossier.

2. **Script de Tests (`tests_run.py`)**
   - Assurez-vous que le fichier `C3_accredited_users.xlsx` a été généré par le script principal.
   - Placez les fichiers `colleague_output.xlsx`, `comparison_data.xlsx`, et `JDE.xlsx` dans le même dossier que le script.
   - Exécutez le script :
     ```
     python tests_run.py
     ```
   - Les résultats des tests seront affichés dans le terminal et des fichiers de comparaison seront générés, comme `comparison_results.xlsx`.

## Structure des Fichiers

- **people.xlsx** : Données des utilisateurs.
- **custom.xlsx** : Données supplémentaires des utilisateurs.
- **departements.xlsx** : Contient les informations sur les départements, les utilisateurs nominatifs, les habilitations ELR, et les utilisateurs à exclure.
- **colleague_output.xlsx** : Liste des utilisateurs C3 identifiés par un collègue.
- **comparison_data.xlsx** : Contient les listes de mails fournies par 'jerome' et 'clement'.
- **JDE.xlsx** : Liste des utilisateurs à vérifier.

## Résolution des Problèmes

- **Erreurs lors de l'exécution** : Assurez-vous que tous les fichiers requis sont bien placés dans le même dossier que les scripts.
- **Résultats inattendus** : Vérifiez les fichiers d'entrée pour vous assurer qu'ils contiennent les bonnes données et qu'ils sont bien formatés.
