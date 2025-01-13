# Génération de Planning

Ce programme Python permet de générer un planning à partir de fichiers texte contenant des informations de calendrier, de binômes et de PSP. Le planning généré est exporté sous forme de fichier Excel.

## Prérequis

Avant d'utiliser ce programme, assurez-vous d'avoir installé les bibliothèques suivantes :

- `tkinter`
- `openpyxl`
- `pytz`
- `Pillow`

Vous pouvez les installer en utilisant pip :

```bash
pip install openpyxl pytz Pillow
```

# Génération de Planning

Ce programme Python permet de générer un planning à partir de fichiers texte contenant des informations de calendrier, de binômes et de PSP. Le planning généré est exporté sous forme de fichier Excel.

## Prérequis

Avant d'utiliser ce programme, assurez-vous d'avoir installé les bibliothèques suivantes :

- `tkinter`
- `openpyxl`
- `pytz`
- `Pillow`

Vous pouvez les installer en utilisant pip :

```bash
pip install openpyxl pytz Pillow
```

## Installation de Python sur macOS

Pour utiliser ce programme, vous devez avoir Python installé sur votre machine. Voici les étapes pour installer Python sur macOS :

1. **Télécharger l'installateur Python** :
   - Rendez-vous sur le site officiel de Python : [python.org](https://www.python.org/).
   - Cliquez sur "Downloads" dans le menu principal.
   - Téléchargez l'installateur pour macOS.

2. **Installer Python** :
   - Ouvrez le fichier téléchargé (généralement un fichier `.pkg`).
   - Suivez les instructions de l'installateur.
   - Assurez-vous de cocher l'option "Install Certificates.command" pour installer les certificats SSL nécessaires.

3. **Vérifier l'installation** :
   - Ouvrez le Terminal.
   - Tapez la commande suivante pour vérifier que Python est correctement installé :
     ```bash
     python3 --version
     ```
   - Vous devriez voir la version de Python installée.

4. **Installer pip (si nécessaire)** :
   - `pip` est généralement inclus avec l'installation de Python. Pour vérifier, tapez :
     ```bash
     pip3 --version
     ```
   - Si `pip` n'est pas installé, vous pouvez l'installer en utilisant la commande suivante :
     ```bash
     sudo easy_install pip
     ```


## Utilisation

1. **Lancer le programme** :
   Exécutez le script Python. Une interface graphique s'ouvrira.

2. **Sélectionner les fichiers** :
   - **Fichier calendrier (.txt)** : Cliquez sur "Parcourir" pour sélectionner le fichier texte contenant les informations de calendrier.
   - **Fichier binômes PS (.txt)** : Cliquez sur "Parcourir" pour sélectionner le fichier texte contenant les informations des binômes.
   - **Fichier PSP (.txt)** : Cliquez sur "Parcourir" pour sélectionner le fichier texte contenant les informations des PSP.
   - **Dossier de sortie** : Cliquez sur "Parcourir" pour sélectionner le dossier où le fichier Excel généré sera enregistré.

3. **Générer le planning** :
   Cliquez sur le bouton "Générer le planning". Le programme traitera les fichiers sélectionnés et générera un fichier Excel contenant le planning.

4. **Vérifier le statut** :
   Le statut de l'opération sera affiché dans la zone de statut en bas de l'interface. En cas de succès, un message de confirmation s'affichera avec le chemin du fichier généré. En cas d'erreur, un message d'erreur s'affichera.

## Fonctionnalités

- **Interface graphique** : Utilisation de `tkinter` pour une interface utilisateur conviviale.
- **Gestion des fichiers** : Sélection des fichiers texte et du dossier de sortie via des boîtes de dialogue.
- **Génération de fichier Excel** : Création d'un fichier Excel avec les informations de planning, triées et formatées.
- **Fusion des cellules** : Fusion des cellules similaires pour une meilleure lisibilité.
- **Statistiques** : Ajout d'une feuille de statistiques dans le fichier Excel généré.

## Exemple de fichiers d'entrée

- **Fichier calendrier (.txt)** : Contient les événements au format iCalendar (BEGIN:VEVENT, DTSTART, DTEND, SUMMARY, END:VEVENT).
- **Fichier binômes PS (.txt)** : Contient la liste des binômes, un par ligne.
- **Fichier PSP (.txt)** : Contient la liste des PSP, un par ligne.

## Exemple de fichier de sortie

Le fichier Excel généré contiendra plusieurs feuilles :
- Une feuille par semaine avec les informations de planning.
- Une feuille "Statistiques" avec le nombre de passages pour chaque binôme et PSP.

## Personnalisation

Vous pouvez personnaliser l'icône de l'application en modifiant la variable `icon_path` dans la fonction `main()`. L'icône peut être au format `.ico` ou `.png`.
