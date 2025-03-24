# Générateur de codes barres et de documents Word

## Description
Ce script PowerShell permet de générer un certain nombre de codes-barres (de type Code128) en utilisant l'API TEC-IT, puis de les organiser dans un document Word. Les utilisateurs peuvent spécifier un préfixe, un numéro de machine, et un dossier de destination pour enregistrer les codes-barres générés sous forme d'images PNG.

## Fonctionnalités principales
1. **Interface graphique utilisateur (GUI)** : Le script utilise une interface graphique pour collecter les informations nécessaires (numéro de machine, nombre de codes à générer, dossier de sauvegarde).
2. **Sélection de dossier** : Permet à l'utilisateur de parcourir et de choisir un dossier pour enregistrer les codes-barres.
3. **Génération de codes-barres** : Crée un nombre spécifié de codes-barres en concaténant un préfixe et le numéro de la machine avec des numéros incrémentaux.
4. **Exportation vers Word** : Insère les codes-barres générés dans un document Word, en les alternant entre deux colonnes dans un tableau.
5. **Gestion des erreurs** : Affiche des messages d'erreur dans l'interface en cas de problème, par exemple si le dossier n'existe pas.

## Prérequis
- **PowerShell** (version 5.1 ou supérieure)
- **Microsoft Word** installé pour l'exportation vers Word.
- **Connexion Internet** pour télécharger les images des codes-barres à partir de l'API TEC-IT.

## Instructions d'utilisation

### Étape 1 : Activer l'exécution de scripts
Si l'exécution de scripts est désactivée sur votre machine, vous devrez d'abord activer l'exécution de scripts PowerShell en exécutant la commande suivante :

```powershell
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force
```

### Étape 2 : Exécuter le script
1. Téléchargez le script PowerShell.
2. Ouvrez une fenêtre PowerShell et naviguez jusqu'au dossier contenant le script.
3. Exécutez le script avec la commande suivante :

```powershell
.\vide_de_lignes.ps1
```

### Étape 3 : Utiliser l'interface graphique
Une fois le script lancé, une fenêtre apparaîtra avec les champs suivants :
- **Numéro de machine** : Entrez le numéro de la machine (par exemple, "Machine001").
- **Nombre de codes** : Spécifiez le nombre de codes-barres à générer.
- **Dossier d'enregistrement** : Sélectionnez un dossier où les images PNG des codes-barres seront sauvegardées.
- **Parcourir** : Utilisez ce bouton pour choisir un dossier de sauvegarde.
- **Générer les codes** : Cliquez sur ce bouton pour lancer la génération des codes-barres.

Le script générera les codes-barres et les enregistrera dans le dossier spécifié. Ensuite, un document Word sera créé avec les codes-barres insérés dans un tableau.

### Étape 4 : Visualiser les résultats
Une fois la génération terminée, vous trouverez le document Word créé dans le dossier spécifié. Le document contiendra les images des codes-barres insérées dans un tableau à deux colonnes.

## Détails techniques

### Structure de l'interface
- **Formulaire principal** : Le formulaire a une taille fixe, avec un `TableLayoutPanel` pour la disposition des éléments (textes, champs de saisie et boutons).
- **Champs de saisie** : Chaque champ (Numéro de machine, Nombre de codes, Dossier de sauvegarde) est associé à un `RichTextBox` permettant de saisir les informations.
- **État** : Un `RichTextBox` supplémentaire affiche l'état d'avancement des opérations et des erreurs.
  
### Génération des codes-barres
Les codes-barres sont générés à l'aide de l'API TEC-IT, avec un format `Code128` et une résolution de 96 DPI. Chaque code-barre est enregistré en tant qu'image PNG dans le dossier choisi.

### Exportation vers Word
Après la génération des codes-barres, un document Word est créé. Les images des codes-barres sont insérées dans un tableau à deux colonnes. Le document est ensuite sauvegardé sous le nom `vide_de_lignes.docx`.

## Problèmes courants et solutions

- **Le dossier de sauvegarde n'existe pas** : Si le dossier spécifié n'existe pas, le script le crée automatiquement.
- **Erreur de connexion à l'API TEC-IT** : Si la connexion à l'API échoue, le script affichera un message d'erreur pour chaque code-barre qui n'a pas pu être généré.


## Licence
Ce script est mis à disposition pour le Groupe Mayr-Melnhof gratuitement.
