@echo off
setlocal enabledelayedexpansion
echo ============================================================
echo INSTALADOR AUTOMATIZADO - Proyecto Python + GitHub
echo Usuario GitHub: mpacheco@dynamizatic.com
echo ============================================================
echo.

REM Verificar si Python está instalado
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python no esta instalado o no esta en el PATH
    echo.
    echo Por favor instala Python desde: https://python.org
    echo Asegurate de marcar "Add Python to PATH" durante la instalacion
    pause
    exit /b 1
)

echo [OK] Python encontrado
python --version
echo.

REM Verificar si Git está instalado
git --version >nul 2>&1
if errorlevel 1 (
    echo [INFO] Git no encontrado - no es obligatorio para el funcionamiento
) else (
    echo [OK] Git encontrado
    
    REM Verificar si ya existe un repositorio Git
    if exist ".git" (
        echo [OK] Repositorio Git encontrado
    ) else (
        echo.
        echo [INFO] No se encontro repositorio Git en este proyecto
        set /p crear_repo="Quieres crear un repositorio Git? (s/n): "
        
        if /i "!crear_repo!"=="s" (
            echo [INFO] Inicializando repositorio Git...
            git init
            if errorlevel 1 (
                echo [ERROR] Error inicializando repositorio Git
            ) else (
                echo [OK] Repositorio Git creado exitosamente
                
                REM Crear .gitignore si no existe
                if not exist ".gitignore" (
                    echo [INFO] Creando archivo .gitignore...
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
                        echo.
                        echo # Archivos temporales
                        echo *.tmp
                        echo *.temp
                        echo ~*
                    ) > .gitignore
                    echo [OK] Archivo .gitignore creado
                )
                
                echo [INFO] Agregando archivos al repositorio...
                git add .
                git commit -m "Initial commit: Descargador de Correos de Microsoft"
                if errorlevel 1 (
                    echo [WARNING] Error haciendo commit inicial
                ) else (
                    echo [OK] Commit inicial realizado
                )
                
                REM Obtener el nombre de la carpeta actual
                for %%I in (.) do set "PROJECT_NAME=%%~nxI"
                set "GITHUB_USER=mpacheco@dynamizatic.com"
                
                echo.
                echo [INFO] Configurando repositorio de GitHub automaticamente...
                echo [INFO] Usuario GitHub: !GITHUB_USER!
                echo [INFO] Nombre del repositorio: !PROJECT_NAME!
                
                REM Verificar si GitHub CLI está instalado
                gh --version >nul 2>&1
                if errorlevel 1 (
                    echo [WARNING] GitHub CLI no encontrado - instalando configuracion manual...
                    
                    REM Configurar remote origin automáticamente
                    echo [INFO] Configurando repositorio remoto...
                    git remote remove origin >nul 2>&1
                    git remote add origin https://github.com/!GITHUB_USER!/!PROJECT_NAME!.git
                    git branch -M main
                    
                    echo [INFO] Repositorio configurado para sincronizacion automatica
                    echo [INFO] Para completar la configuracion en GitHub:
                    echo [INFO] 1. Instala GitHub CLI: https://cli.github.com/
                    echo [INFO] 2. Ejecuta: gh auth login
                    echo [INFO] 3. Crea el repositorio: gh repo create !PROJECT_NAME! --public
                    echo [INFO] 4. Sincroniza: git push -u origin main
                    echo.
                    echo [INFO] O crea manualmente en: https://github.com/new
                    echo [INFO] Nombre: !PROJECT_NAME!
                    echo [INFO] Luego ejecuta: git push -u origin main
                ) else (
                    echo [OK] GitHub CLI encontrado - configurando automaticamente...
                    
                    REM Verificar si el repositorio ya existe en GitHub
                    gh repo view !GITHUB_USER!/!PROJECT_NAME! >nul 2>&1
                    if errorlevel 1 (
                        echo [INFO] Repositorio no existe en GitHub - creando...
                        
                        REM Verificar autenticación
                        gh auth status >nul 2>&1
                        if errorlevel 1 (
                            echo [INFO] Autenticando con GitHub CLI...
                            gh auth login --hostname github.com --protocol https --web
                        )
                        
                        REM Crear repositorio público automáticamente
                        echo [INFO] Creando repositorio publico en GitHub...
                        gh repo create !PROJECT_NAME! --public --source=. --remote=origin --push
                        
                        if errorlevel 1 (
                            echo [WARNING] Error creando repositorio automaticamente
                            echo [INFO] Configurando manualmente...
                            git remote remove origin >nul 2>&1
                            git remote add origin https://github.com/!GITHUB_USER!/!PROJECT_NAME!.git
                            git branch -M main
                            echo [INFO] Crea el repositorio manualmente en: https://github.com/new
                            echo [INFO] Nombre: !PROJECT_NAME!
                            echo [INFO] Luego ejecuta: git push -u origin main
                        ) else (
                            echo [OK] Repositorio creado y sincronizado exitosamente!
                            echo [INFO] URL: https://github.com/!GITHUB_USER!/!PROJECT_NAME!
                        )
                    ) else (
                        echo [OK] Repositorio ya existe en GitHub - sincronizando...
                        
                        REM Configurar remote si no existe
                        git remote get-url origin >nul 2>&1
                        if errorlevel 1 (
                            git remote add origin https://github.com/!GITHUB_USER!/!PROJECT_NAME!.git
                        )
                        
                        REM Sincronizar con GitHub
                        git branch -M main
                        git push -u origin main
                        
                        if errorlevel 1 (
                            echo [WARNING] Error sincronizando - podria ser por cambios remotos
                            echo [INFO] Sincronizando con pull primero...
                            git pull origin main --allow-unrelated-histories
                            git push -u origin main
                        )
                        
                        echo [OK] Repositorio sincronizado con GitHub!
                        echo [INFO] URL: https://github.com/!GITHUB_USER!/!PROJECT_NAME!
                    )
                )
            )
        ) else (
            echo [INFO] Repositorio Git no creado
        )
    )
)

echo.

REM Verificar si el entorno virtual ya existe
if exist ".venv" (
    echo [OK] Entorno virtual encontrado, activando...
) else (
    echo [INFO] Creando entorno virtual...
    python -m venv .venv
    if errorlevel 1 (
        echo [ERROR] Error creando entorno virtual
        pause
        exit /b 1
    )
    echo [OK] Entorno virtual creado exitosamente
)

echo.
echo [INFO] Activando entorno virtual...
call .venv\Scripts\activate.bat

echo.
echo [INFO] Actualizando pip...
python -m pip install --upgrade pip

echo.
REM Verificar si existe requirements.txt
if exist "requirements.txt" (
    echo [INFO] Instalando dependencias desde requirements.txt...
    pip install -r requirements.txt
    
    if errorlevel 1 (
        echo [ERROR] Error instalando dependencias
        pause
        exit /b 1
    )
    echo [OK] Dependencias instaladas exitosamente
) else (
    echo [INFO] No se encontro archivo requirements.txt
    echo [INFO] Saltando instalacion de dependencias...
)

echo.
echo ============================================================
echo [EXITO] INSTALACION COMPLETADA EXITOSAMENTE!
echo ============================================================
echo.
echo Para ejecutar la aplicacion:
echo    1. Activa el entorno virtual: .venv\Scripts\activate.bat
echo    2. Ejecuta la aplicacion: python main_alternative.py
echo.
echo Tip: Tambien puedes ejecutar directamente:
echo    .venv\Scripts\python.exe main_alternative.py
echo.
echo O usa el script rapido: run.bat
echo.
echo ============================================================