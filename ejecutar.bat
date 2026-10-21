@echo off
title NeXo Software - Automatizacion
cd /d "%~dp0"
python app_gui.py
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo Ocurrio un error al iniciar la aplicacion.
    pause
)
