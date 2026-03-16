@echo off
title Tilstra KPI Watcher — actief
color 0A

echo.
echo  ================================================
echo   Tilstra KPI Dashboard — Automatische Watcher
echo  ================================================
echo.
echo  Dit venster bijhouden (minimaliseren is OK).
echo  Sluit dit venster om de watcher te stoppen.
echo.

REM Ga naar de map van dit bestand
cd /d "%~dp0"

REM Controleer Python
where python >nul 2>&1
if %errorlevel% == 0 (
    python watcher.py
    goto einde
)

where python3 >nul 2>&1
if %errorlevel% == 0 (
    python3 watcher.py
    goto einde
)

echo FOUT: Python niet gevonden!
echo.
echo Download Python via: https://www.python.org/downloads/
echo Vink "Add Python to PATH" aan tijdens de installatie.
echo.

:einde
pause
