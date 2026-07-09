@echo off
rem Lanceur Windows de l'application de traitement des factures.
rem Double-cliquez sur ce fichier pour ouvrir l'application.
cd /d "%~dp0"

rem Utilise l'environnement virtuel s'il existe, sinon le Python du système.
if exist ".venv\Scripts\pythonw.exe" (
    start "" ".venv\Scripts\pythonw.exe" application.py
) else (
    start "" pythonw application.py
)
