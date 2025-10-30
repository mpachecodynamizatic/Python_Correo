@echo off
echo ============================================================
echo DESCARGADOR DE CORREOS DE MICROSOFT
echo ============================================================
echo.

REM Verificar si el entorno virtual existe
if not exist ".venv" (
    echo [ERROR] Entorno virtual no encontrado
    echo.
    echo [INFO] Ejecuta primero install.bat para configurar el entorno
    echo.
    pause
    exit /b 1
)

echo [INFO] Activando entorno virtual...
call .venv\Scripts\activate.bat

echo [INFO] Iniciando aplicacion...
echo.
python main_alternative.py

echo.
echo [INFO] Gracias por usar el Descargador de Correos!
pause