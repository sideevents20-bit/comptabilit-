@echo off
rem ============================================================
rem  Installateur de l'application « Traitement des factures »
rem  Double-cliquez sur ce fichier : il installe tout ce qu'il
rem  faut et cree un raccourci sur le Bureau.
rem  Prerequis : Python 3.10+ (https://www.python.org/downloads/
rem  en cochant "Add python.exe to PATH" pendant l'installation).
rem ============================================================
cd /d "%~dp0"
echo.
echo === Installation de l'application Traitement des factures ===
echo.

where python >nul 2>nul
if errorlevel 1 (
    echo [ERREUR] Python n'est pas installe ou pas dans le PATH.
    echo Telechargez-le sur https://www.python.org/downloads/
    echo et cochez "Add python.exe to PATH" pendant l'installation.
    echo.
    pause
    exit /b 1
)

echo [1/4] Creation de l'environnement Python isole (.venv)...
python -m venv .venv || (echo [ERREUR] Creation du venv impossible. & pause & exit /b 1)

echo [2/4] Installation des dependances (quelques minutes)...
".venv\Scripts\python.exe" -m pip install --quiet --upgrade pip
".venv\Scripts\python.exe" -m pip install --quiet -r requirements.txt || (
    echo [ERREUR] Installation des dependances impossible. & pause & exit /b 1)

echo [3/4] Creation du raccourci sur le Bureau...
powershell -NoProfile -Command ^
  "$ws = New-Object -ComObject WScript.Shell; $s = $ws.CreateShortcut([Environment]::GetFolderPath('Desktop') + '\Traitement des factures.lnk'); $s.TargetPath = '%~dp0Lancer_Factures.bat'; $s.WorkingDirectory = '%~dp0'; $s.Description = 'Traitement automatique des factures'; $s.Save()"

echo [4/4] Verification de Tesseract (lecture des factures scannees)...
where tesseract >nul 2>nul
if errorlevel 1 (
    if not exist "C:\Program Files\Tesseract-OCR\tesseract.exe" (
        echo.
        echo [A FAIRE] Tesseract OCR n'est pas installe : les factures
        echo scannees ne pourront pas etre lues automatiquement.
        echo Telechargez-le ici ^(cochez le pack "French"^) :
        echo   https://github.com/UB-Mannheim/tesseract/wiki
    )
)

echo.
echo === Installation terminee ! ===
echo Un raccourci "Traitement des factures" a ete cree sur le Bureau.
echo Au premier lancement, l'ecran de reglages s'ouvrira pour saisir
echo le mot de passe d'application de la boite mtgsud.compta@gmail.com.
echo.
pause
