@echo off
setlocal enabledelayedexpansion
echo ============================================================
echo AUTENTICACION RAPIDA GITHUB CLI
echo ============================================================
echo.

REM Agregar GitHub CLI al PATH si no está disponible
set "PATH=%PATH%;C:\Program Files\GitHub CLI"

echo [INFO] Verificando GitHub CLI...
gh --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] GitHub CLI no encontrado
    echo [INFO] Instala desde: https://cli.github.com/
    pause
    exit /b 1
)

echo [OK] GitHub CLI encontrado
gh --version

echo.
echo [INFO] Iniciando autenticacion con GitHub...
echo [INFO] Se abrira el navegador automaticamente
echo [INFO] Autoriza la aplicacion en el navegador y regresa aqui
echo.
pause

REM Autenticar con GitHub usando el método más simple
gh auth login --web

if errorlevel 1 (
    echo [ERROR] Error en autenticacion
    echo [INFO] Intenta con opciones adicionales:
    echo [INFO] gh auth login --hostname github.com --git-protocol https --web
    pause
    exit /b 1
) else (
    echo [OK] Autenticacion exitosa!
)

echo.
echo [INFO] Verificando autenticacion...
gh auth status

echo.
echo [INFO] Usuario autenticado:
gh api user --jq .login

echo.
echo ============================================================
echo [EXITO] GITHUB CLI AUTENTICADO CORRECTAMENTE!
echo ============================================================
echo.
echo Ahora puedes usar:
echo   - install.bat (para configuracion completa)
echo   - sync.bat (para sincronizacion)
echo   - gh repo create Python_Correo --public
echo.
pause