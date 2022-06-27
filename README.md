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

# Utilisation

* Pour générer les images au format PNG depuis le fichier excel :
    ```bash
    python3 create_badges.py
    ```
