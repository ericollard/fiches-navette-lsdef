# Générateur de fiches navette UE, LSDEF NOF 27

Outil en ligne qui produit les fiches navette des unités d'enseignement de la licence
Sciences de l'éducation et de la formation, nouvelle offre de formation 2027.

**Adresse de l'outil : `https://ericollard.github.io/fiches-navette-lsdef/`**

Cette adresse s'affiche dans `Settings` puis `Pages` une fois la publication activée.
C'est celle-là qu'il faut communiquer aux collègues.

## Le principe

L'outil est une page web unique. Vous ouvrez la page, vous choisissez le classeur de base
sur votre ordinateur, vous obtenez une fiche Word par UE, téléchargeables une par une ou
en une seule archive.

Le traitement se fait entièrement dans le navigateur, en JavaScript. Le classeur n'est pas
envoyé, pas déposé sur un serveur, pas enregistré ailleurs que sur votre poste. Une fois la
page chargée, elle fonctionne même sans connexion.

## Ce que contient ce dépôt

| Fichier | Rôle |
|---|---|
| `index.html` | la page, son interface et sa logique d'affichage |
| `generateur.js` | le cœur du générateur, commun à cette page et au programme Python de référence |
| `modele/Fiche_navette_modele_balise.docx` | le modèle de fiche, balisé, modifications suivies acceptées |
| `vendor/` | les deux bibliothèques employées, SheetJS pour lire le classeur et JSZip pour écrire les fichiers Word |

## Ce que ce dépôt ne contient pas, et ne doit jamais contenir

**Le classeur de base n'est pas ici et ne doit pas y être déposé.** Un site GitHub Pages est
public sur l'internet, y compris lorsqu'il est servi depuis un dépôt privé : la publication
de Pages rend les fichiers accessibles à toute personne qui connaît l'adresse. Or la base
porte les noms et les courriels des collègues ainsi que leurs contraintes d'emploi du temps.
Elle reste donc sur les postes et circule par les canaux habituels de l'équipe.

Le fichier `.gitignore` écarte par précaution les classeurs et les fiches produites, afin
qu'un `git add .` ne les envoie pas par mégarde.

## Mise en service

1. Déposer le contenu de ce dossier dans le dépôt, à la racine. Le fichier `index.html` doit
   se trouver à la racine et non dans un sous-dossier, sans quoi l'adresse renverra une
   erreur 404.
2. Dans le dépôt, ouvrir `Settings`, puis `Pages`.
3. Sous `Build and deployment`, choisir la source `Deploy from a branch`, la branche `main`
   et le dossier `/ (root)`, puis `Save`.
4. Attendre une à deux minutes. L'adresse publiée s'affiche en haut de la même page.

Pages réclame un dépôt public, sauf abonnement payant. Puisque rien de nominatif ne figure
ici, un dépôt public convient.

## Usage, côté collègue

1. Ouvrir l'adresse de l'outil dans un navigateur récent, Safari, Firefox, Chrome ou Edge.
2. Déposer le classeur de base dans le cadre prévu, ou cliquer pour le choisir.
3. Vérifier le récapitulatif affiché, qui donne le nombre d'UE, la répartition par semestre,
   les régimes de préremplissage et les manques éventuels.
4. Cliquer sur le bouton de production, puis télécharger l'archive.

## Où sont les données

Aucun renseignement sur la formation ne figure dans le programme, ni dans la page. Tout ce
qui s'écrit dans les fiches vient du classeur, y compris les phrases fixes :

- `Ref_Textes` : toutes les phrases de la fiche et leurs gabarits. Les parties entre
  accolades, par exemple `{libelle}`, sont remplacées par les valeurs lues dans les autres
  feuilles. Changer une formulation se fait dans cette feuille.
- `Ref_Niveaux` : le libellé du niveau et son qualificatif, employés dans la phrase d'amorce
  de la section 2.
- `Ref_Blocs_maquette` : le libellé exact à écrire dans la case Bloc de maquette.
- `Ref_Gabarit_fiche` : la correspondance entre les régimes d'évaluation et les colonnes de
  la grille de la section 4.

Si une clé de texte manque ou si un gabarit emploie une variable inconnue, l'outil s'arrête
en nommant la clé en cause plutôt que de produire des fiches incomplètes.

## Rapport avec le programme Python

Le programme `generer_fiches.py`, employé par le comité, et cette page partagent la même
logique et la même source de données. Les deux ont été comparés fiche par fiche sur la base
de novembre 2026 : les 49 fiches produites sont identiques quant au contenu. La page
s'adresse aux collègues qui ne souhaitent pas installer Python, le programme reste l'outil
de travail du comité.

## Mettre à jour le modèle de fiche

Remplacer `modele/Fiche_navette_modele_balise.docx` par la nouvelle version balisée, puis
pousser la modification. Un modèle balisé porte les jetons de la forme `§NOM§` aux
emplacements à remplir. Tant que cette opération n'est pas faite, le bouton
« Utiliser un autre modèle » permet d'en charger un depuis son poste pour une production
ponctuelle.
