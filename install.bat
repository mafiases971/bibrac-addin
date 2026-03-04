@echo off
chcp 65001 >nul
echo ============================================
echo   Installation BIBRAC Add-in pour Excel
echo ============================================
echo.

:: Dossier d'installation local
set INSTALL_DIR=%USERPROFILE%\Office-Addins

:: Créer le dossier si nécessaire
if not exist "%INSTALL_DIR%" (
    mkdir "%INSTALL_DIR%"
    echo [OK] Dossier créé : %INSTALL_DIR%
) else (
    echo [OK] Dossier existant : %INSTALL_DIR%
)

:: Télécharger le manifest depuis GitHub
echo.
echo Téléchargement du manifest...
powershell -Command "Invoke-WebRequest -Uri 'https://raw.githubusercontent.com/mafiases971/bibrac-addin/main/manifest.xml' -OutFile '%INSTALL_DIR%\manifest-bibrac.xml'" 2>nul
if %errorlevel% neq 0 (
    echo [ERREUR] Téléchargement échoué. Vérifiez votre connexion internet.
    pause
    exit /b 1
)
echo [OK] Manifest téléchargé.

:: Ajouter le dossier comme catalogue de confiance dans le registre Office
echo.
echo Configuration du registre Office...
set REG_PATH=HKCU\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{BIBRAC-ADDIN-CATALOG}
reg add "%REG_PATH%" /v "Url" /t REG_SZ /d "%INSTALL_DIR%" /f >nul
reg add "%REG_PATH%" /v "Flags" /t REG_DWORD /d 1 /f >nul
echo [OK] Registre configuré.

:: Terminé
echo.
echo ============================================
echo   Installation terminée !
echo ============================================
echo.
echo Prochaines étapes dans Excel :
echo  1. Fermer et rouvrir Excel
echo  2. Accueil ^> Compléments ^> Obtenir des compléments
echo  3. Onglet "Dossier partagé"
echo  4. Cliquer sur "BIBRAC Add-in" ^> Ajouter
echo.
pause
