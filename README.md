# Descargador de Correos de Microsoft

Una aplicación en Python que se conecta a tu correo de Microsoft usando dos métodos alternativos: IMAP directo o Microsoft Graph API con Device Code Flow. Descarga correos en un rango de fechas específico y los exporta a un archivo Excel con formato profesional.

## Características

- ✅ **Doble método de autenticación:**
  - IMAP directo (rápido y eficiente)
  - Microsoft Graph API con Device Code Flow (compatible con todas las cuentas)
- ✅ **Acceso a múltiples carpetas** (INBOX y "1 - JIRA")
- ✅ Descarga de correos en un rango de fechas personalizable
- ✅ Extrae fecha, asunto, remitente, dominio y carpeta
- ✅ **Exporta a Excel** con formato profesional y estilos
- ✅ Interfaz completamente en español
- ✅ Compatible con todas las cuentas Microsoft sin excepción

## Requisitos

- Python 3.7 o superior
- Cuenta de Microsoft (Outlook, Hotmail, Office 365)
- Para IMAP: IMAP habilitado en tu cuenta
- Para Graph API: No requiere configuraciones especiales

## Instalación Rápida

### Opción 1: Instalación completamente automatizada (Recomendada)
```bash
install.bat
```
Este script realiza **TODO AUTOMÁTICAMENTE**:
- ✅ Verifica que Python esté instalado
- ✅ **Detecta GitHub CLI** (aunque no esté en PATH)
- ✅ **Autentica automáticamente** con GitHub (abre navegador)
- ✅ **Crea repositorio en GitHub** usando tu usuario `mpacheco@dynamizatic.com`
- ✅ **Sincroniza automáticamente** todo el código con GitHub
- ✅ Detecta e inicializa repositorio Git si es necesario
- ✅ Crea automáticamente el entorno virtual (si no existe)
- ✅ Activa el entorno virtual
- ✅ Instala todas las dependencias necesarias (si existe requirements.txt)

### Scripts adicionales:
```bash
setup_github.bat    # Configuración avanzada de GitHub CLI
sync.bat           # Sincronización rápida posterior
```

### Opción 2: Instalación manual
```bash
python -m venv .venv
.venv\Scripts\activate.bat
pip install -r requirements.txt
```

## Dependencias

- `requests>=2.31.0` - Para Microsoft Graph API
- `openpyxl>=3.1.0` - Para exportación a Excel

## Uso

### Opción 1: Ejecución rápida (Recomendada)
```bash
run.bat
```
Este script activa automáticamente el entorno virtual y ejecuta la aplicación.

### Opción 2: Ejecución manual
1. **Activa el entorno virtual:**
   ```bash
   .venv\Scripts\activate.bat
   ```

2. **Ejecuta la aplicación:**
   ```bash
   python main_alternative.py
   ```

3. **Selecciona método de autenticación:**
   - **Método 1 (IMAP)**: Más rápido, requiere IMAP habilitado
   - **Método 2 (Graph API)**: Compatible con todas las cuentas, usa navegador web

4. **Configura rango de fechas** y revisa los resultados

## Ejemplo de Uso - Graph API

```
============================================================
DESCARGADOR DE CORREOS DE MICROSOFT
============================================================

¿Qué método prefieres? (1=IMAP, 2=Graph API): 2

🌐 Iniciando autenticación Graph API (Device Code)...
============================================================
📋 INSTRUCCIONES DE AUTENTICACIÓN
============================================================
🌐 Ve a: https://microsoft.com/devicelogin
🔑 Ingresa este código: ABC123
👤 Inicia sesión con tu cuenta de Microsoft
✅ Autoriza la aplicación
============================================================

🎉 ¡Autenticación exitosa!
👤 Usuario: Juan Pérez (juan@empresa.com)

==================================================
📅 CONFIGURACIÓN DE FECHAS
==================================================
Fecha de inicio [2025-10-23]: 2025-10-01
Fecha de fin [2025-10-30]: 2025-10-29

📥 Buscando correos desde 2025-10-01 hasta 2025-10-29...
📂 Carpetas a revisar: inbox, 1 - JIRA

📁 Procesando carpeta: inbox
   📧 Total de correos obtenidos: 312

📁 Procesando carpeta: 1 - JIRA
   📧 Total de correos obtenidos: 77

==================================================
📊 RESUMEN DE RESULTADOS
==================================================
📧 Total de correos encontrados: 389

📂 Distribución por carpetas:
   inbox: 312 correos
   1 - JIRA: 77 correos

🏢 Dominios más frecuentes:
   empresa.com: 185 correos
   cliente.es: 107 correos

💾 ¿Exportar a Excel? (s/n): s
💾 Correos exportados a: correos_20251030_070432.xlsx
```

## Estructura del Proyecto

```
Python_Correo/
├── .venv/                   # Entorno virtual (creado automáticamente)
├── .git/                    # Repositorio Git (creado automáticamente)
├── main_alternative.py     # Aplicación principal
├── device_auth.py          # Autenticación Graph API
├── graph_email_manager.py  # Gestión de correos Graph API
├── email_manager.py        # Gestión de correos IMAP
├── requirements.txt        # Dependencias de Python
├── install.bat            # 🚀 Instalación COMPLETAMENTE automatizada
├── setup_github.bat       # 🔧 Configuración avanzada GitHub CLI
├── run.bat                # Script de ejecución rápida
├── sync.bat               # Script de sincronización rápida con GitHub
└── README.md              # Documentación completa
```

## Características Técnicas

### Integración con Git/GitHub
- **Configuración automática**: Usuario GitHub preconfigurado (`mpacheco@dynamizatic.com`)
- **Inicialización automática**: El script detecta si el proyecto está en un repositorio Git y lo inicializa si es necesario
- **Creación de repositorio remoto**: Integración con GitHub CLI para crear automáticamente el repositorio en GitHub
- **Sincronización inteligente**: Detecta si el repositorio ya existe y sincroniza automáticamente
- **Script de sincronización**: `sync.bat` para actualizaciones rápidas posteriores
- **Configuración inteligente**: Detección automática del nombre del proyecto basado en la carpeta actual
- **Instrucciones de respaldo**: Si GitHub CLI no está disponible, proporciona instrucciones detalladas para configuración manual

## Formato del Archivo Excel Exportado

El archivo Excel contiene:

| Columna | Descripción | Ejemplo |
|---------|-------------|---------|
| Fecha | Fecha en formato DD/MM/YYYY | 29/10/2025 |
| Fecha Completa | Fecha y hora completa | 2025-10-29 14:30:22 |
| Asunto | Asunto del correo | Reunión de proyecto |
| Remitente | Email completo del remitente | juan@empresa.com |
| Dominio | Dominio del remitente | empresa.com |
| Carpeta | Carpeta del correo | inbox / 1 - JIRA |

### Exportación a Excel Profesional
- **Formato automatizado**: Headers estilizados en azul con texto blanco
- **Ajuste inteligente**: Columnas ajustadas automáticamente al contenido
- **Formato de fechas**: DD/MM/YYYY para análisis en español
- **Información completa**: Fecha, asunto, remitente, dominio y carpeta de origen

### Entorno Virtual y Dependencias
- **Gestión automática**: Creación y activación automática del entorno virtual
- **Instalación inteligente**: Detecta requirements.txt y instala dependencias automáticamente
- **Verificación Python**: Valida instalación de Python antes de proceder

**Características del Excel:**
- ✅ Encabezados con formato azul y texto blanco
- ✅ Columnas ajustadas automáticamente
- ✅ Fecha en formato DD/MM/YYYY para fácil análisis

## Solución de Problemas

### Método IMAP
- **Error IMAP**: Habilita IMAP en Outlook.com → Configuración → Sincronización
- **2FA**: Usa contraseña de aplicación desde tu cuenta Microsoft

### Método Graph API
- **Error de autorización**: Repite el proceso de Device Code
- **Navegador no abre**: Copia la URL manualmente
- **Código expirado**: La aplicación generará uno nuevo automáticamente

## Ventajas por Método

### IMAP (Método 1)
- ✅ Más rápido
- ✅ No requiere navegador
- ❌ Requiere IMAP habilitado

### Graph API (Método 2)
- ✅ Compatible con TODAS las cuentas Microsoft
- ✅ Funciona aunque IMAP esté deshabilitado
- ✅ Método oficial de Microsoft
- ❌ Requiere autorización en navegador

## Seguridad

- ✅ **Graph API**: Usa Device Code Flow oficial de Microsoft
- ✅ **IMAP**: Conexión cifrada SSL/TLS
- ✅ No almacena credenciales en archivos
- ✅ Tokens temporales con expiración automática
- ✅ Solo acceso de lectura a correos

## Limitaciones

- Graph API procesa hasta 500 correos por carpeta por consulta
- IMAP procesa hasta 100 correos por consulta
- Ambos métodos optimizados para rendimiento

## Licencia

Este proyecto está bajo la Licencia MIT.