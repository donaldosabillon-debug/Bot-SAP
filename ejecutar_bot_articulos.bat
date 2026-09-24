@echo off
chcp 65001 > nul
title Bot Actualizador de Articulos SAP B1
cd /d "%~dp0"
echo ====================================================================
echo      BOT ACTUALIZADOR DE ARTICULOS: SAP B1 HANA (PLANIFICACION)
echo                     Lazarus & Lazarus
echo ====================================================================
echo.
echo Iniciando interfaz grafica...
start "" pythonw app_articulos_gui.py
if %ERRORLEVEL% NEQ 0 (
    python app_articulos_gui.py
)
exit
