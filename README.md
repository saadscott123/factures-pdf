# Générateur de Factures PDF

Cette application web permet de générer des factures PDF à partir d'un fichier Excel (.xlsx).

## Fonctionnalités

- Upload de fichiers Excel (.xlsx)
- Génération automatique de factures PDF
- Interface drag & drop
- Téléchargement des factures générées

## Format du fichier Excel

Le fichier Excel doit contenir les colonnes suivantes dans cet ordre exact :

1. Facture Numero
2. Date de facture
3. Client
4. Date de Depart
5. Date de Retour
6. Marque du Vehicule
7. Matricule
8. Prix Total Ht
9. TVA
10. Prix TTC

**Important** : 
- L'ordre des colonnes doit être strictement respecté
- Toutes les colonnes sont obligatoires
- Les données manquantes seront ignorées

## Prérequis

- Python 3.8 ou supérieur
- pip (gestionnaire de paquets Python)

## Installation

1. Cloner le repository :
```bash
git clone [votre-repo]
cd facturespdf
```

2. Créer un environnement virtuel et l'activer :
```bash
python -m venv venv
source venv/bin/activate  # Sur macOS/Linux
```

3. Installer les dépendances :
```bash
pip install -r requirements.txt
```

## Utilisation

1. Démarrer l'application :
```bash
python app.py
```

2. Ouvrir votre navigateur et accéder à `http://localhost:5000`

3. Préparer votre fichier Excel avec les colonnes requises dans l'ordre spécifié

4. Glisser-déposer votre fichier Excel ou cliquer pour le sélectionner

5. Les factures PDF seront générées automatiquement et disponibles pour téléchargement

## Format des Factures PDF

Chaque facture générée comprendra :

- Numéro de facture
- Informations client
- Dates de départ et de retour
- Informations sur le véhicule (marque et matricule)
- Détails du prix (HT, TVA, TTC)

## Support

Pour toute question ou problème, veuillez créer une issue dans le repository.
