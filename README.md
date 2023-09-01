# Projet SAAFI

Ce script a été conçu pour automatiser le processus de téléchargement de fichiers à partir d'une liste d'URL donnée et pour les organiser dans des dossiers sur votre bureau.
Il cible spécifiquement l'extraction de données à partir d'un fichier Excel avec certaines colonnes.

## Prérequis 
Avant d'exécuter le script, assurez-vous d'avoir installé les dépendances suivantes :
- openpyxl
- requests

## Utilisation 
1. Modifiez la variable 'DATA' pour spécifier le chemin de votre fichier Excel.
2. Définissez la variable 'RESULT_LOCATION' avec le répertoire où vous souhaitez créer des dossiers et enregistrer les fichiers téléchargés. Mettez à jour cette variable avec le chemin de votre dossier souhaité.

### Remarque : 
- Le script générera des noms de dossiers en fonction des valeurs de certaines colonnes (prénom, nom de famille et un identifiant) de votre fichier Excel. Assurez-vous que les colonnes de votre fichier Excel correspondent aux indices de colonne définis dans le code.
- Vous devrez peut-être personnaliser davantage le script si la structure ou les données de votre fichier Excel diffèrent des hypothèses faites dans ce script.