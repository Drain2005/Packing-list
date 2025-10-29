@echo off
chcp 65001
title Générateur DE FICHIER
mode con: cols=80 lines=20
color 0A

echo.
echo    ========================================
echo         GÉNÉRATEUR DE FICHIER
echo    ========================================
echo.
echo        Initialisation en cours...
echo        Veuillez patienter...
echo.
echo    ========================================
echo.

python --version >nul 2>&1
if errorlevel 1 (
    echo ERREUR : Python non installé
    echo.
    echo Téléchargez Python depuis python.org
    echo.
    pause
    exit /b 1
)

if not exist "venv\Scripts\python.exe" (
    echo Creation de l'environnement...
    python -m venv venv
    call venv\Scripts\activate.bat
    echo Installation des composants...
    pip install --upgrade pip >nul 2>&1
    pip install Django==4.2.7 pandas==2.3.3 openpyxl python-barcode reportlab pywin32 Pillow whitenoise python-decouple dj-database-url >nul 2>&1
) else (
    call venv\Scripts\activate.bat
)

python manage.py migrate >nul 2>&1

cls
echo.
echo    ========================================
echo         GÉNÉRATEUR DE FICHIER - PRET
echo    ========================================
echo.
echo    ✅ Application demarree avec succes !
echo.
echo    Ouverture du navigateur...
echo    Adresse : http://127.0.0.1:8000
echo.
echo    Pour quitter : fermez cette fenetre
echo    ========================================
echo.

timeout /t 3 /nobreak >nul
start "" "http://127.0.0.1:8000" >nul 2>&1

python manage.py runserver

echo.
echo Application fermee.
timeout /t 3