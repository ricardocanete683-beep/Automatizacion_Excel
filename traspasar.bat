@echo off
title Proceso de Staging - Auditoria Produccion
color 0B

echo ============================================
echo   INICIANDO ORGANIZACION DE PRODUCCION
echo ============================================
echo.

:: Verificamos si Python esta instalado
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python no esta instalado o no esta en el PATH.
    pause
    exit
)

:: Ejecutamos el script de Python
python copiar.py

echo.
echo ============================================
echo   PROCESO TERMINADO
echo ============================================
pause