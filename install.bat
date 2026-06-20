@echo off
chcp 65001 >nul
echo ============================================================
echo  INSTALLATION - EasyCompta
echo ============================================================
echo.

:: Vérifier Python
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERREUR] Python n'est pas installé ou pas dans le PATH.
    echo Téléchargez Python depuis : https://www.python.org/downloads/
    echo Cochez "Add Python to PATH" lors de l'installation.
    pause
    exit /b 1
)

echo [OK] Python détecté :
python --version
echo.

:: Installer les dépendances
echo Installation des bibliothèques Python...
pip install openpyxl pdfplumber pypdf
if errorlevel 1 (
    echo [ERREUR] Echec de l'installation. Essayez : pip install --upgrade pip
    pause
    exit /b 1
)

echo.
echo [OK] Bibliothèques installées avec succès.
echo.

:: Créer le dossier pdfs
if not exist "pdfs\" (
    mkdir pdfs
    echo [OK] Dossier pdfs\ créé.
)

echo.
echo ============================================================
echo  INSTALLATION TERMINÉE
echo ============================================================
echo.
echo Pour démarrer :
echo   Ouvrez EasyCompta.xlsm dans Excel et activez les macros.
echo.
pause
