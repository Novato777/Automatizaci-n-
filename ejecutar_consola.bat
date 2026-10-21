@echo off
title NeXo Prospeccion - Consola
chcp 65001 >nul
cd /d "%~dp0"
python enviar_campana.py
pause
