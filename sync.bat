@echo off
setlocal enabledelayedexpansion
echo ============================================================
echo SINCRONIZADOR AUTOMATICO - GitHub
echo Usuario: mpacheco@dynamizatic.com
echo ============================================================
echo.

REM Agregar GitHub CLI al PATH si está instalado
set "PATH=%PATH%;C:\Program Files\GitHub CLI"

REM Obtener el nombre de la carpeta actual
for %%I in (.) do set "PROJECT_NAME=%%~nxI"

REM Detectar usuario de GitHub automáticamente si está autenticado
for /f "tokens=*" %%i in ('gh api user --jq .login 2^>nul') do set "GITHUB_USER=%%i"
if "!GITHUB_USER!"=="" set "GITHUB_USER=mpachecodynamizatic"

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

REM Verificar GitHub CLI
gh --version >nul 2>&1
if errorlevel 1 (
    echo [WARNING] GitHub CLI no encontrado
    echo [INFO] Usando Git tradicional para sincronizacion...
    goto :git_sync_only
)

echo [OK] GitHub CLI encontrado

REM Verificar autenticación
gh auth status >nul 2>&1
if errorlevel 1 (
    echo [WARNING] No estas autenticado con GitHub CLI
    echo [INFO] Para autenticarte automaticamente ejecuta: auth_github.bat
    echo [INFO] O ejecuta manualmente: gh auth login --web
    echo.
    set /p auth_now="Quieres autenticarte ahora? (s/n): "
    if /i "!auth_now!"=="s" (
        echo [INFO] Abriendo autenticacion...
        gh auth login --web
        if errorlevel 1 (
            echo [WARNING] Error en autenticacion - usando Git tradicional
            goto :git_sync_only
        )
    ) else (
        echo [INFO] Usando Git tradicional para sincronizacion...
        goto :git_sync_only
    )
)

echo [OK] Autenticado con GitHub CLI - continuando...

:git_sync_only
echo.
echo [INFO] Verificando estado del repositorio...
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
    echo [INFO] 3. Ejecuta: gh auth login --web (si usas GitHub CLI)
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