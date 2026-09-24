@echo off
chcp 65001 > nul
title Bot Sincronizador Clientes SAP-Zoho
cd /d "%~dp0"
echo ====================================================================
echo        BOT SINCRONIZADOR DE CLIENTES: SAP B1 HANA -> ZOHO CRM
echo                     Lazarus & Lazarus
echo ====================================================================
echo.
echo Iniciando interfaz grafica...
start "" pythonw app_gui.py
if %ERRORLEVEL% NEQ 0 (
    python app_gui.py
)
exit
