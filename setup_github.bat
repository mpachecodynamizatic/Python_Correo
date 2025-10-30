@echo off
setlocal enabledelayedexpansion
echo ============================================================
echo INSTALADOR DE GITHUB CLI + CONFIGURACION AUTOMATICA
echo Usuario GitHub: mpacheco@dynamizatic.com
echo ============================================================
echo.

REM Verificar si GitHub CLI ya está instalado
echo [INFO] Verificando GitHub CLI...
gh --version >nul 2>&1
if errorlevel 1 (
    echo [INFO] GitHub CLI no encontrado - procediendo con instalacion...
    echo.
    echo [INFO] Descargando e instalando GitHub CLI...
    
    REM Verificar si winget está disponible
    winget --version >nul 2>&1
    if errorlevel 1 (
        echo [WARNING] winget no encontrado
        echo [INFO] Instalacion manual requerida:
        echo [INFO] 1. Ve a: https://cli.github.com/
        echo [INFO] 2. Descarga GitHub CLI para Windows
        echo [INFO] 3. Instala el archivo descargado
        echo [INFO] 4. Reinicia PowerShell/CMD
        echo [INFO] 5. Ejecuta este script nuevamente
        pause
        exit /b 1
    ) else (
        echo [OK] winget encontrado - instalando GitHub CLI...
        winget install --id GitHub.cli
        
        if errorlevel 1 (
            echo [WARNING] Error con winget - instrucciones manuales:
            echo [INFO] 1. Ve a: https://cli.github.com/
            echo [INFO] 2. Descarga GitHub CLI para Windows
            echo [INFO] 3. Instala el archivo descargado
            echo [INFO] 4. Reinicia PowerShell/CMD
            echo [INFO] 5. Ejecuta este script nuevamente
            pause
            exit /b 1
        ) else (
            echo [OK] GitHub CLI instalado exitosamente!
            echo [INFO] Reinicia PowerShell/CMD y ejecuta este script nuevamente
            pause
            exit /b 0
        )
    )
) else (
    echo [OK] GitHub CLI encontrado!
    gh --version
)

echo.
echo [INFO] Verificando autenticacion con GitHub...

REM Verificar autenticación
gh auth status >nul 2>&1
if errorlevel 1 (
    echo [INFO] No estas autenticado - iniciando proceso de autenticacion...
    echo.
    echo [INFO] Se abrira el navegador para autenticacion
    echo [INFO] 1. Autoriza la aplicacion en el navegador
    echo [INFO] 2. Regresa aqui cuando termine
    echo.
    pause
    
    gh auth login --hostname github.com --git-protocol https --web
    
    if errorlevel 1 (
        echo [ERROR] Error en autenticacion
        echo [INFO] Intenta manualmente: gh auth login
        pause
        exit /b 1
    ) else (
        echo [OK] Autenticacion exitosa!
    )
) else (
    echo [OK] Ya estas autenticado en GitHub
    gh auth status
)

echo.
echo [INFO] Verificando usuario autenticado...
for /f "tokens=*" %%i in ('gh api user --jq .login 2^>nul') do set "CURRENT_USER=%%i"

if "!CURRENT_USER!"=="" (
    echo [WARNING] No se pudo obtener el usuario actual
    echo [INFO] Verifica tu autenticacion: gh auth status
) else (
    echo [OK] Usuario autenticado: !CURRENT_USER!
)

REM Obtener nombre del proyecto actual
for %%I in (.) do set "PROJECT_NAME=%%~nxI"
echo [INFO] Proyecto actual: !PROJECT_NAME!

echo.
echo [INFO] Verificando si el repositorio existe en GitHub...

REM Verificar si el repositorio ya existe
gh repo view !PROJECT_NAME! >nul 2>&1
if errorlevel 1 (
    echo [INFO] Repositorio '!PROJECT_NAME!' no existe - creando...
    
    echo [INFO] Creando repositorio publico en GitHub...
    gh repo create !PROJECT_NAME! --public --description "Descargador de Correos de Microsoft - Aplicacion Python con autenticacion Graph API e IMAP"
    
    if errorlevel 1 (
        echo [ERROR] Error creando repositorio
        echo [INFO] Intenta manualmente: gh repo create !PROJECT_NAME! --public
        pause
        exit /b 1
    ) else (
        echo [OK] Repositorio creado exitosamente!
        echo [INFO] URL: https://github.com/!CURRENT_USER!/!PROJECT_NAME!
    )
) else (
    echo [OK] Repositorio '!PROJECT_NAME!' ya existe
    echo [INFO] URL: https://github.com/!CURRENT_USER!/!PROJECT_NAME!
)

echo.
echo [INFO] Configurando Git local si es necesario...

REM Verificar si estamos en un repositorio Git
if not exist ".git" (
    echo [INFO] Inicializando repositorio Git...
    git init
    
    REM Crear .gitignore si no existe
    if not exist ".gitignore" (
        echo [INFO] Creando .gitignore...
        (
            echo # Entorno virtual
            echo .venv/
            echo __pycache__/
            echo *.pyc
            echo.
            echo # Archivos de salida
            echo *.xlsx
            echo *.csv
            echo.
            echo # Archivos del sistema
            echo .DS_Store
            echo Thumbs.db
            echo desktop.ini
        ) > .gitignore
    )
    
    echo [INFO] Agregando archivos al repositorio...
    git add .
    git commit -m "Initial commit: Descargador de Correos de Microsoft"
)

REM Configurar remote origin
git remote get-url origin >nul 2>&1
if errorlevel 1 (
    echo [INFO] Configurando remote origin...
    git remote add origin https://github.com/!CURRENT_USER!/!PROJECT_NAME!.git
) else (
    echo [INFO] Remote origin ya configurado
)

echo [INFO] Configurando rama main...
git branch -M main

echo.
echo [INFO] Sincronizando con GitHub...

REM Intentar push
git push -u origin main >nul 2>&1
if errorlevel 1 (
    echo [INFO] Primera sincronizacion o conflictos - resolviendo...
    git pull origin main --allow-unrelated-histories >nul 2>&1
    git push -u origin main >nul 2>&1
    
    if errorlevel 1 (
        echo [WARNING] Error sincronizando - podrian existir conflictos
        echo [INFO] Resuelve manualmente con: git status
    ) else (
        echo [OK] Sincronizacion completada!
    )
) else (
    echo [OK] Repositorio sincronizado exitosamente!
)

echo.
echo ============================================================
echo [EXITO] GITHUB CLI CONFIGURADO EXITOSAMENTE!
echo ============================================================
echo.
echo Repositorio: https://github.com/!CURRENT_USER!/!PROJECT_NAME!
echo.
echo Ahora puedes usar:
echo   - gh repo view !PROJECT_NAME!
echo   - git push origin main
echo   - sync.bat (para sincronizacion rapida)
echo.
echo ============================================================
pause