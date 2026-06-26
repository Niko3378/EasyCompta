@echo off
setlocal EnableDelayedExpansion
cd /d "%~dp0"

:: -- Version du produit -------------------------------------------------------
set VERSION=1.2.0
set VERSION_WIX=%VERSION%.0

echo.
echo ============================================================
echo   BUILD  ^|  EasyCompta v%VERSION% - Installateur .msi
echo ============================================================
echo.

:: -- Etape 0 : verifier les fichiers sources --------------------------------
if not exist "EasyCompta.xlsx" (
    echo [ERREUR] EasyCompta.xlsx introuvable dans %~dp0
    pause & exit /b 1
)
if not exist "vba_Comptabilite.bas" (
    echo [ERREUR] vba_Comptabilite.bas introuvable dans %~dp0
    pause & exit /b 1
)
echo [OK] Fichiers sources trouves.

:: -- Etape 1 : verifier Python + dependances -------------------------------
echo.
echo [1/5] Verification de Python et des dependances...

python --version >nul 2>&1
if errorlevel 1 (
    echo [ERREUR] Python introuvable dans le PATH.
    echo Installez Python 3.10+ depuis python.org
    pause & exit /b 1
)

pip install pywin32 fpdf2 pyinstaller --quiet 2>&1
if errorlevel 1 (
    echo [AVERTISSEMENT] Certaines dependances n'ont pas pu etre installees.
    echo Verifiez votre connexion internet et relancez.
)
echo [OK] Dependances Python verifiees.

:: -- Etape 2 : integration VBA automatique -> cree EasyCompta.xlsm --
echo.
echo [2/5] Integration VBA automatique ^(via Excel COM^)...
echo       Excel va s'ouvrir brievement en arriere-plan.
echo.
python integrer_vba.py
if errorlevel 1 (
    echo.
    echo [ERREUR] Echec de l'integration VBA.
    echo.
    echo  Causes possibles :
    echo   - Excel n'est pas installe sur cette machine
    echo   - L'acces au projet VBA est refuse par Excel :
    echo     Ouvrez Excel ^> Fichier ^> Options ^> Centre de gestion de la confidentialite
    echo     ^> Parametres du centre de gestion de la confidentialite
    echo     ^> Parametres des macros ^> cochez "Approbation acces au modele d'objet VBA"
    echo.
    pause & exit /b 1
)
if not exist "EasyCompta.xlsm" (
    echo [ERREUR] EasyCompta.xlsm non cree apres integration VBA.
    pause & exit /b 1
)
echo [OK] EasyCompta.xlsm cree avec VBA integre.

:: -- Etape 3 : generation du manuel PDF ------------------------------------
echo.
echo [3/5] Generation du manuel utilisateur PDF...
python generer_manuel.py
if errorlevel 1 (
    echo [ERREUR] Echec de la generation du PDF.
    pause & exit /b 1
)
if not exist "Manuel_EasyCompta.pdf" (
    echo [ERREUR] Manuel_EasyCompta.pdf non cree.
    pause & exit /b 1
)
echo [OK] Manuel_EasyCompta.pdf genere.

:: -- Etape 4 : compiler pdf_extractor.exe ----------------------------------
echo.
echo [4/5] Compilation de pdf_extractor.exe ^(PyInstaller --onefile^)...
echo       Cette etape peut prendre 2 a 5 minutes - patientez.
echo.
pyinstaller --clean pdf_extractor.spec
if errorlevel 1 (
    echo [ERREUR] Echec de PyInstaller.
    pause & exit /b 1
)
if not exist "dist\pdf_extractor.exe" (
    echo [ERREUR] dist\pdf_extractor.exe introuvable apres compilation.
    pause & exit /b 1
)
echo [OK] dist\pdf_extractor.exe cree.

:: -- Etape 5 : compiler le MSI ---------------------------------------------
echo.
echo [5/5] Compilation du MSI ^(WiX Toolset^)...

set WIX_BIN=
for %%D in (
    "C:\Program Files (x86)\WiX Toolset v3.14\bin"
    "C:\Program Files (x86)\WiX Toolset v3.11\bin"
    "C:\Program Files (x86)\WiX Toolset v3.10\bin"
    "C:\Program Files\WiX Toolset v3.14\bin"
    "C:\Program Files\WiX Toolset v3.11\bin"
) do (
    if exist "%%~D\candle.exe" (
        set WIX_BIN=%%~D
        goto :wix_found
    )
)
where candle.exe >nul 2>&1
if not errorlevel 1 goto :wix_found

echo [ERREUR] WiX Toolset v3 introuvable.
echo Telechargez wix314.exe sur : https://github.com/wixtoolset/wix3/releases
pause & exit /b 1

:wix_found
if defined WIX_BIN (
    set CANDLE="%WIX_BIN%\candle.exe"
    set LIGHT="%WIX_BIN%\light.exe"
) else (
    set CANDLE=candle.exe
    set LIGHT=light.exe
)

cd installer

%CANDLE% EasyCompta.wxs -ext WixUtilExtension -dVersion=%VERSION_WIX% -out EasyCompta.wixobj -arch x64
if errorlevel 1 (
    echo [ERREUR] Echec candle.exe
    cd .. & pause & exit /b 1
)

set MSI_OUT=EasyCompta_v%VERSION%.msi
if exist "%MSI_OUT%" del /q "%MSI_OUT%" 2>nul

%LIGHT% EasyCompta.wixobj -ext WixUIExtension -ext WixUtilExtension -cultures:fr-fr -sice:ICE38 -out "%MSI_OUT%"
if errorlevel 1 (
    echo [ERREUR] Echec light.exe
    cd .. & pause & exit /b 1
)

del /q EasyCompta.wixobj 2>nul
del /q EasyCompta.wixpdb 2>nul
cd ..

echo.
echo ============================================================
echo   SUCCES !
echo   Installateur cree : installer\EasyCompta_v%VERSION%.msi
echo ============================================================
echo.
echo  Contenu de l'installateur :
echo   - EasyCompta.xlsm  ^(VBA integre, popup PayPal avec cle licence^)
echo   - pdf_extractor.exe        ^(extraction PDF autonome^)
echo   - Manuel_EasyCompta.pdf ^(guide utilisateur complet^)
echo   - Raccourci Bureau + Menu Demarrer
echo   - Python non requis sur la machine cible
echo.
pause
