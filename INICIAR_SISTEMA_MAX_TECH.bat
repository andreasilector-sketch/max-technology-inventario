@echo off
title MAX TECHNOLOGY STORE - Panel de Control ^& Sincronizador
color 0C
echo ========================================================
echo   MAX TECHNOLOGY STORE - SISTEMA DE INVENTARIO Y WEB
echo ========================================================
echo.
echo Iniciando panel de control en tu navegador...
timeout /t 1 >nul
start http://localhost:8000/admin
python server.py
pause