@echo off
title Actualizador de Keystores SSL
color 0A
setlocal EnableDelayedExpansion

:: Ejecutar desde la carpeta del script (paths relativos)
cd /d "%~dp0"

echo ============================================
echo   AUDITORIA SSL v5.0
echo   Actualizacion de Keystores
echo ============================================
echo.

:: Verificar que Python este instalado
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python no esta instalado o no esta en el PATH.
    echo         Instala Python desde https://python.python.org
    echo.
    pause
    exit /b 1
)

:: Mostrar version de Python
for /f "tokens=*" %%v in ('python --version 2^>^&1') do echo   Python: %%v
echo.

:: Ejecutar el script principal
echo   Iniciando procesamiento...
echo.
python "procesar.py" %*
set RESULTADO=%errorlevel%

echo.
if %RESULTADO% equ 0 (
    echo ============================================
    echo   PROCESO FINALIZADO EXITOSAMENTE
    echo ============================================
) else (
    echo ============================================
    echo   [ADVERTENCIA] El proceso termino con
    echo   codigo de error: %RESULTADO%
    echo ============================================
)

echo.
echo   Log:     LOG_PROCESAMIENTO.txt
echo   Alertas: LOG_VENCIMIENTOS.txt
echo   Reporte: REPORTE_AUDITORIA.html
echo.
pause
exit /b %RESULTADO%