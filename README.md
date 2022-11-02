# Badges Forum Atlantique

Projet de génération automatique de badges pour le staff et les intervenants
entreprises pour le Forum Atlantique

## Installation

* Installez d'abord Python3 si ce n'est pas déjà fait (pour Mac c'est
    installé par défaut, pour Linux utilisez apt, et pour Windows utilisez
    le Windows Store)
* Installez les libraries (cf le fichier `requirements.txt`)

> Pour installer les libraries, ouvrez un terminal dans ce dossier et
> exécutez la commande suivante :
> ```bash
> python3 -m pip install -r requirements.txt
> ```
> ou pour installer les libraries une par une :
> ```bash
> python3 -m pip install <nom-de-la-library>
> ```

## Utilisation

### Générer les images PNG

* Dans le fichier `create_badges.py`, modifiez le nom du fichier excel d'entrée.
* Exécutez la commande :
    ```bash
    python3 create_badges.py
    ```
* Suivez les instructions dans le terminal
* Les images au format PNG apparaissent dans le dossier `badges/` (par défaut)

### Générer le PDF

* Dans le fichier `create_badges.py`, modifiez le nom du fichier excel d'entrée.
* Exécutez la commande :
    ```bash
    python3 create_pdf.py
    ```
* Le PDF est créé avec le nom `badges_FA.pdf` (par défaut)

## Reste à faire

Ci-dessous une liste d'améliorations possibles pour le projet :
* écrire le texte "Equipe Organisatrice" depuis le script python plutôt que sur
    l'image de fond (ça permettra à l'avenir de changer plus facilement le
    background s'ils en ont envie)
* réajuster les espacements du texte pour les entreprises (et peut-être les
    autres types aussi mais j'ai noté que c'était pas ouf pour les entreprises)
* ajouter le nom de l'entreprise qu'ils pilotent pour les élèves pilotes
* super-bonus : générer tout le texte directement sur le PDF et ne plus utiliser
    d'images intermédiaires (ça permettra d'avoir un format vectoriel pour le
    PDF et donc une bien meilleure qualité)

## Contact

Vous faites partie de l'équipe du Forum Atlantique et vous souhaiteriez contribuer au projet ? Contactez directement les contributeurs du projet :
* [@hydrielax](https://github.com/hydrielax)
* [@LeRatDuFA](https://github.com/LeRatDuFA)
