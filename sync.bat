@echo off
setlocal enabledelayedexpansion
echo ============================================================
echo SINCRONIZADOR AUTOMATICO - GitHub
echo Usuario: mpacheco@dynamizatic.com
echo ============================================================
echo.

REM Obtener el nombre de la carpeta actual
for %%I in (.) do set "PROJECT_NAME=%%~nxI"
set "GITHUB_USER=mpacheco@dynamizatic.com"

echo [INFO] Sincronizando proyecto: !PROJECT_NAME!
echo [INFO] Usuario GitHub: !GITHUB_USER!
echo.

REM Verificar si Git está instalado
git --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Git no encontrado
    echo [INFO] Instala Git desde: https://git-scm.com/
    pause
    exit /b 1
)

REM Verificar si existe repositorio Git
if not exist ".git" (
    echo [ERROR] No se encontro repositorio Git en esta carpeta
    echo [INFO] Ejecuta install.bat primero para inicializar el repositorio
    pause
    exit /b 1
)

echo [OK] Git encontrado - iniciando sincronizacion...

REM Verificar estado del repositorio
git status --porcelain >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Error verificando estado del repositorio
    pause
    exit /b 1
)

echo [INFO] Agregando todos los cambios...
git add .

REM Verificar si hay cambios para commit
git diff --staged --quiet
if errorlevel 1 (
    REM Hay cambios - hacer commit
    set /p commit_msg="Mensaje del commit (o presiona Enter para mensaje automatico): "
    
    if "!commit_msg!"=="" (
        REM Generar mensaje automático con fecha y hora
        for /f "tokens=1-4 delims=/ " %%a in ('date /t') do set fecha=%%a-%%b-%%c
        for /f "tokens=1-2 delims=: " %%a in ('time /t') do set hora=%%a:%%b
        set "commit_msg=Actualizacion automatica !fecha! !hora!"
    )
    
    echo [INFO] Haciendo commit: !commit_msg!
    git commit -m "!commit_msg!"
    
    if errorlevel 1 (
        echo [ERROR] Error haciendo commit
        pause
        exit /b 1
    )
) else (
    echo [INFO] No hay cambios nuevos para commit
)

REM Verificar remote origin
git remote get-url origin >nul 2>&1
if errorlevel 1 (
    echo [INFO] Configurando remote origin...
    git remote add origin https://github.com/!GITHUB_USER!/!PROJECT_NAME!.git
)

echo [INFO] Sincronizando con GitHub...

REM Verificar si la rama main existe en remoto
git ls-remote --heads origin main >nul 2>&1
if errorlevel 1 (
    echo [INFO] Primera subida al repositorio remoto...
    git branch -M main
    git push -u origin main
) else (
    echo [INFO] Actualizando repositorio remoto...
    REM Primero hacer pull por si hay cambios remotos
    git pull origin main --allow-unrelated-histories >nul 2>&1
    REM Luego push
    git push origin main
)

if errorlevel 1 (
    echo [WARNING] Error sincronizando con GitHub
    echo [INFO] Posibles soluciones:
    echo [INFO] 1. Verifica tu conexion a internet
    echo [INFO] 2. Verifica que el repositorio existe: https://github.com/!GITHUB_USER!/!PROJECT_NAME!
    echo [INFO] 3. Ejecuta: gh auth login (si usas GitHub CLI)
    echo.
    echo [INFO] Intenta sincronizar manualmente:
    echo [INFO]   git push origin main
    pause
    exit /b 1
) else (
    echo [OK] Sincronizacion completada exitosamente!
    echo [INFO] Repositorio actualizado: https://github.com/!GITHUB_USER!/!PROJECT_NAME!
)

echo.
echo ============================================================
echo [EXITO] SINCRONIZACION COMPLETADA!
echo ============================================================
echo.
echo El proyecto esta sincronizado con GitHub
echo URL: https://github.com/!GITHUB_USER!/!PROJECT_NAME!
echo.
echo ============================================================
pause