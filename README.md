# Générateur de Code QR (Code128) pour Machines

Ce script PowerShell permet de générer une série de codes QR (Code128) personnalisés pour une machine spécifique. Vous pouvez spécifier un nom de machine, un nombre de codes à générer, et le dossier de destination pour enregistrer les fichiers images des codes QR générés.

## Prérequis

- PowerShell 5.0 ou supérieur
- Une connexion Internet pour télécharger les images des codes QR via l'API TEC-IT.
- Les bibliothèques .NET suivantes :
  - `System.Windows.Forms`
  - `System.Drawing`

## Fonctionnalités

1. **Interface graphique (GUI)** pour la saisie des informations nécessaires :
   - Numéro de la machine
   - Nombre de codes à générer
   - Dossier de destination pour l'enregistrement des fichiers image.

2. **Génération des codes QR** :
   - Les codes QR sont créés à l'aide de l'API TEC-IT.
   - Les codes QR sont enregistrés sous forme d'images PNG dans le dossier spécifié.

3. **Validation des entrées** :
   - Vérifie que tous les champs sont remplis.
   - Vérifie que le nombre de codes est un entier valide et positif.
   - Crée le dossier de destination si celui-ci n'existe pas.

## Utilisation

### Étape 1 : Lancer le script

1. Ouvrez PowerShell.
2. Exécutez le script fourni.

### Étape 2 : Remplir le formulaire

1. **Numéro de machine** : Entrez un nom ou un identifiant pour la machine.
2. **Nombre de codes** : Indiquez combien de codes QR vous souhaitez générer.
3. **Dossier d'enregistrement** : Spécifiez le dossier où vous souhaitez enregistrer les fichiers PNG des codes QR. Vous pouvez également cliquer sur le bouton "Parcourir" pour sélectionner un dossier via l'explorateur de fichiers.

### Étape 3 : Générer les codes QR

Cliquez sur le bouton **"Générer les codes"** pour démarrer la génération des codes QR. Le script affichera des messages de progression dans la zone de texte de statut. À la fin du processus, vous trouverez les fichiers PNG dans le dossier spécifié.

## Exemple d'utilisation

1. Numéro de machine : **Machine001**
2. Nombre de codes : **5**
3. Dossier d'enregistrement : **C:\Users\Utilisateur\Documents\CodesQR**

Après avoir cliqué sur **"Générer les codes"**, le script générera cinq codes QR pour **Machine001** et les enregistrera dans le dossier **C:\Users\Utilisateur\Documents\CodesQR**.

## Notes

- Le script utilise un préfixe fixe (`Faf@|Q7G8TH7bGxg#e!6Jqj_`) avant d'ajouter le nom de la machine et un numéro incrémental pour chaque code généré.
- Les codes sont créés en utilisant l'API de TEC-IT et sont en format **Code128**.
- Le script gère également les erreurs, affichant un message dans la zone de statut si un problème survient lors de la génération d'un code QR.

## Dépannage

- **Si le dossier d'enregistrement n'existe pas** : Il sera créé automatiquement.
- **Si un code QR échoue à se générer** : Un message d'erreur s'affichera dans la zone de statut.
