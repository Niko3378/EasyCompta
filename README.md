# EasyCompta v2.0.0

Système de gestion comptable automatisé pour TPE/indépendants — Excel + VBA + Python.

## Fonctionnalités

- **Import automatique de factures PDF** : extraction des données (numéro, date, montants HT/TVA/TTC, fournisseur) via `pdf_extractor.exe`
- **Base de données Fournisseurs & Clients** : saisie manuelle ou import PDF avec listes déroulantes
- **Calcul de TVA mensuelle** : TVA collectée, TVA déductible, solde à payer
- **Tableaux de bord analytiques** : graphiques et pivots sur `Pivot_Analytics`
- **Clé de licence** : activation par don PayPal, vérification intégrée dans le classeur

## Installation

Téléchargez et exécutez `EasyCompta_v2.0.0.msi` (aucune installation de Python requise).

L'installateur crée :
- Un raccourci Bureau et Menu Démarrer
- `EasyCompta.xlsm` — le classeur principal
- `pdf_extractor.exe` — l'extracteur PDF autonome
- `Manuel_EasyCompta.pdf` — le guide utilisateur

## Utilisation

1. Ouvrez `EasyCompta.xlsm` dans Excel et activez les macros
2. Feuille **PDF_Import** : cliquez sur "Importer PDF" pour extraire une facture
3. Feuilles **Fournisseurs_DB** / **Clients_DB** : gérez vos factures
4. Feuille **TVA_Mensuelle** : calculez votre TVA (sélectionnez le mois en B3)
5. Feuille **Pivot_Analytics** : rafraîchissez les tableaux pour visualiser vos données

## Soutenir le projet

Ce logiciel est gratuit. Un don PayPal vous permet d'obtenir une clé d'activation :

**[paypal.me/NLaurent878](https://paypal.me/NLaurent878)**

## Build (développeurs)

### Prérequis

- Python 3.10+
- Microsoft Excel
- [WiX Toolset v3](https://github.com/wixtoolset/wix3/releases)

```
pip install openpyxl pdfplumber pypdf pywin32 fpdf2 pyinstaller
```

### Lancer le build

```bat
build_installer.bat
```

Produit `installer\EasyCompta_v2.0.0.msi` en 5 étapes :
1. Vérification des dépendances Python
2. Intégration VBA dans `EasyCompta.xlsm` via Excel COM
3. Génération du manuel PDF
4. Compilation de `pdf_extractor.exe` (PyInstaller)
5. Compilation du MSI (WiX Toolset)

## Structure du projet

```
EasyCompta/
├── EasyCompta.xlsm          # Classeur principal (VBA intégré)
├── EasyCompta.xlsx          # Classeur source (sans macros)
├── vba_Comptabilite.bas     # Code source VBA
├── pdf_extractor.py         # Extracteur PDF
├── generer_manuel.py        # Génération du manuel PDF
├── integrer_vba.py          # Intégration VBA via Excel COM
├── create_workbook.py       # Génération du classeur
├── generer_cle.py           # Génération de clé de licence
├── build_installer.bat      # Script de build complet
├── pdf_extractor.spec       # Config PyInstaller
├── version.txt              # Version courante
├── dist/                    # pdf_extractor.exe (Git LFS)
└── installer/               # MSI + sources WiX (Git LFS)
```

## Licence

Logiciel gratuit — tous droits réservés © 2026 NLaurent878
