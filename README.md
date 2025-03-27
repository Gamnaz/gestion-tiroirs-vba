# gestion-tiroirs-vba
Code VBA classant des données par tiroirs et change l'unité de valeur (W/kW)

# Classement automatique par tiroirs dans Excel (VBA)

Ce projet propose un outil automatisé dans Excel permettant de classer des données horaires dans trois catégories appelées "tiroirs", en fonction du jour de la semaine et de l’heure. L’ensemble est géré par des macros VBA, avec un bouton interactif intégré à la feuille Excel.

## Objectif

Faciliter l’analyse de séries de mesures en ajoutant automatiquement une colonne "Tiroir" en fonction des règles suivantes :

- Tiroir 1 : Du lundi au vendredi, entre 08h00 et 20h00
- Tiroir 2 : Le samedi, entre 08h00 et 20h00
- Tiroir 3 : Tous les autres cas

Le fichier traite également la conversion automatique des données de puissance, initialement en watts, en kilowatts, avec la possibilité de basculer d'une unité à l'autre via un bouton.

## Fichiers fournis

- `TiroirsAutomatiques.xlsm` : Fichier Excel principal contenant les macros prêtes à l’emploi.
- `ModuleTiroirs.bas` : Export du module VBA si besoin d’importation dans un autre classeur.

## Fonctionnalités des macros

### AjouterBoutonTiroir

Crée un bouton "Appliquer Tiroirs" positionné à l'avance (colonne J, ligne 10) si celui-ci n'existe pas. Ce bouton déclenche l’ensemble du traitement.

### ModifierValeurEtTiroir

Appel central déclenché par le bouton. Enchaîne trois actions :
1. Conversion des valeurs de W vers kW (si non encore converties)
2. Classement automatique en tiroirs
3. Ajout d’un second bouton permettant de rebasculer les unités

### ModifierValeur

Identifie la colonne intitulée "Valeur" (non convertie), divise les données par 1000, et renomme l’en-tête en "Valeur (kW)".

### AppliquerTiroir

Détecte les colonnes "Date de la mesure", "Heure de la mesure", et "Tiroir".
Classe les lignes dans le bon tiroir selon la logique définie.
Crée la colonne "Tiroir" si elle n’existe pas.
Applique une couleur de fond à la cellule de chaque ligne correspondant à son tiroir.

### AjouterBoutonToggleUnite

Ajoute un second bouton "Repasser en W" juste en dessous du bouton principal, qui permet d’alterner dynamiquement les unités.

### ToggleUniteValeur

Convertit les données en sens inverse selon l’unité courante (kW → W ou W → kW) et adapte l’intitulé du bouton et de la colonne en conséquence.

## Utilisation

1. Ouvrir le fichier Excel 'TiroirsAutomatiques.xlsm'
2. Activer les macros
3. Cliquer sur le bouton "Appliquer Tiroirs"
4. Le traitement s’effectue automatiquement :
   - Conversion des unités (si nécessaire)
   - Classification en tiroirs
   - Mise en forme
   - Ajout du bouton secondaire pour changer d’unité

## Compatibilité

- Excel 2016 ou version ultérieure (Windows)
- Macros activées (VBA)

## Auteur

Ce projet a été conçu dans le cadre d’un exercice test pour un entretien d'embauche.
