# main_alternative.py
"""
Aplicación principal con métodos alternativos de autenticación
- IMAP (si está disponible)
- Microsoft Graph API (alternativa sin IMAP)
"""

import sys
from datetime import datetime, timedelta

def choose_authentication_method():
    """Permite al usuario elegir el método de autenticación"""
    print("="*60)
    print("DESCARGADOR DE CORREOS DE MICROSOFT")
    print("="*60)
    print("Esta aplicación puede acceder a tu correo de dos formas:\n")
    
    print("🔹 MÉTODO 1: IMAP (Conexión directa)")
    print("   ✅ Más rápido y eficiente")
    print("   ❌ Requiere habilitar IMAP en tu cuenta")
    print("   📧 Funciona con: Outlook.com, Hotmail, cuentas corporativas con IMAP\n")
    
    print("🔹 MÉTODO 2: Microsoft Graph API (Device Code)")
    print("   ✅ Funciona incluso si IMAP está deshabilitado")
    print("   ✅ Compatible con TODAS las cuentas Microsoft")
    print("   ✅ Método muy seguro y confiable")
    print("   ❌ Requiere autorización en navegador web")
    print("   📧 Funciona con: Todas las cuentas Microsoft sin excepción\n")
    
    while True:
        choice = input("¿Qué método prefieres? (1=IMAP, 2=Graph API, h=ayuda): ").strip().lower()
        
        if choice == '1':
            return 'imap'
        elif choice == '2':
            return 'graph'
        elif choice == 'h' or choice == 'help' or choice == 'ayuda':
            show_method_help()
        else:
            print("❌ Opción inválida. Escribe 1, 2 o h para ayuda.")

def show_method_help():
    """Muestra ayuda detallada sobre los métodos"""
    print("\n" + "="*50)
    print("📚 AYUDA DETALLADA")
    print("="*50)
    
    print("\n🔹 ¿Cuándo usar IMAP (Método 1)?")
    print("   • Tienes una cuenta personal de Outlook/Hotmail")
    print("   • Puedes acceder a la configuración de tu cuenta")
    print("   • Quieres máxima velocidad de descarga")
    print("   • No te molesta configurar IMAP")
    
    print("\n🔹 ¿Cuándo usar Graph API (Método 2)?")
    print("   • Tu cuenta corporativa no permite IMAP")
    print("   • No puedes cambiar configuraciones de cuenta")
    print("   • Prefieres autenticación web segura")
    print("   • El método IMAP te da errores")
    
    print("\n🔹 ¿No estás seguro?")
    print("   • Prueba primero el Método 2 (Graph API)")
    print("   • Es más compatible y fácil de usar")
    print("   • Si no funciona, prueba el Método 1")
    
    print("\n" + "="*50 + "\n")

def run_imap_method():
    """Ejecuta la aplicación usando IMAP"""
    try:
        from auth import MicrosoftAuthenticator
        from email_manager import EmailManager
        
        print("\n🔐 Iniciando autenticación IMAP...")
        authenticator = MicrosoftAuthenticator()
        
        if not authenticator.authenticate():
            print("\n❌ Autenticación IMAP falló.")
            print("💡 Prueba el Método 2 (Graph API) escribiendo: python main_alternative.py")
            return False
        
        # Crear gestor de correos
        email_manager = EmailManager(authenticator.get_imap_connection(), authenticator.get_email_address())
        
        # Continuar con el flujo normal
        return run_email_download(email_manager, authenticator)
        
    except ImportError as e:
        print(f"❌ Error importando módulos IMAP: {str(e)}")
        return False
    except Exception as e:
        print(f"❌ Error en método IMAP: {str(e)}")
        return False

def run_graph_method():
    """Ejecuta la aplicación usando Graph API con Device Code"""
    try:
        from device_auth import DeviceCodeAuthenticator
        from graph_email_manager import GraphEmailManager
        
        print("\n🌐 Iniciando autenticación Graph API (Device Code)...")
        print("📱 Este método es muy compatible - funciona con todas las cuentas Microsoft")
        authenticator = DeviceCodeAuthenticator()
        
        if not authenticator.authenticate():
            print("\n❌ Autenticación Graph API falló.")
            print("💡 Prueba el Método 1 (IMAP) si tienes IMAP habilitado.")
            return False
        
        # Crear gestor de correos
        email_manager = GraphEmailManager(authenticator.get_access_token())
        
        # Continuar con el flujo normal
        return run_email_download(email_manager, authenticator)
        
    except ImportError as e:
        print(f"❌ Error importando módulos Graph: {str(e)}")
        return False
    except Exception as e:
        print(f"❌ Error en método Graph: {str(e)}")
        return False

def get_date_range():
    """Solicita rango de fechas al usuario"""
    print("\n" + "="*50)
    print("📅 CONFIGURACIÓN DE FECHAS")
    print("="*50)
    
    today = datetime.now()
    default_end = today
    default_start = today - timedelta(days=7)
    
    print(f"Configurar el rango de fechas para buscar correos.")
    print(f"Formato: YYYY-MM-DD (ejemplo: {today.strftime('%Y-%m-%d')})")
    
    # Fecha de inicio
    while True:
        start_input = input(f"Fecha de inicio [por defecto: {default_start.strftime('%Y-%m-%d')}]: ").strip()
        
        if not start_input:
            start_date = default_start.replace(hour=0, minute=0, second=0, microsecond=0)
            break
        
        try:
            start_date = datetime.strptime(start_input, '%Y-%m-%d')
            start_date = start_date.replace(hour=0, minute=0, second=0, microsecond=0)
            break
        except ValueError:
            print("❌ Formato inválido. Use YYYY-MM-DD")
    
    # Fecha de fin
    while True:
        end_input = input(f"Fecha de fin [por defecto: {default_end.strftime('%Y-%m-%d')}]: ").strip()
        
        if not end_input:
            end_date = default_end.replace(hour=23, minute=59, second=59, microsecond=999999)
            break
        
        try:
            end_date = datetime.strptime(end_input, '%Y-%m-%d')
            end_date = end_date.replace(hour=23, minute=59, second=59, microsecond=999999)
            
            if end_date >= start_date:
                break
            else:
                print("❌ La fecha de fin debe ser igual o posterior a la de inicio.")
        except ValueError:
            print("❌ Formato inválido. Use YYYY-MM-DD")
    
    return start_date, end_date

def run_email_download(email_manager, authenticator):
    """Ejecuta la descarga de correos (común para ambos métodos)"""
    try:
        # Obtener info del usuario
        user_info = email_manager.get_user_info()
        if user_info:
            print(f"👤 Usuario: {user_info['nombre']} ({user_info['email']})")
        
        # Obtener fechas
        start_date, end_date = get_date_range()
        
        print(f"\n📥 Buscando correos desde {start_date.strftime('%Y-%m-%d')} hasta {end_date.strftime('%Y-%m-%d')}...")
        
        # Descargar correos (incluyendo carpeta JIRA)
        folders_to_search = ['inbox', '1 - JIRA']
        emails = email_manager.get_emails_in_date_range(start_date, end_date, folders_to_search)
        
        # Mostrar resumen
        show_results_summary(emails)
        
        # Exportar si hay correos
        if emails:
            export_choice = input("\n💾 ¿Exportar a Excel? (s/n): ").strip().lower()
            
            if export_choice in ['s', 'si', 'sí', 'yes', 'y']:
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                filename = f"correos_{timestamp}.xlsx"
                email_manager.export_to_excel(emails, filename)
        
        print(f"\n{'='*60}")
        print("✅ ¡Proceso completado exitosamente!")
        print("🙏 Gracias por usar el Descargador de Correos de Microsoft.")
        print(f"{'='*60}")
        
        return True
        
    except KeyboardInterrupt:
        print("\n🛑 Proceso cancelado por el usuario.")
        return False
    except Exception as e:
        print(f"❌ Error durante la descarga: {str(e)}")
        return False
    finally:
        # Limpiar conexiones
        if hasattr(authenticator, 'disconnect'):
            authenticator.disconnect()

def show_results_summary(emails):
    """Muestra resumen de resultados incluyendo análisis por carpeta"""
    if not emails:
        print("\n📭 No se encontraron correos en el rango especificado.")
        return
    
    print(f"\n{'='*50}")
    print(f"📊 RESUMEN DE RESULTADOS")
    print(f"{'='*50}")
    print(f"📧 Total de correos encontrados: {len(emails)}")
    
    # Análisis por carpeta
    carpetas = {}
    for email_data in emails:
        carpeta = email_data.get('carpeta', 'inbox')
        carpetas[carpeta] = carpetas.get(carpeta, 0) + 1
    
    if len(carpetas) > 1:
        print(f"\n📂 Distribución por carpetas:")
        for carpeta, count in carpetas.items():
            print(f"   {carpeta}: {count} correos")
    
    # Contar dominios
    domains = {}
    for email_data in emails:
        domain = email_data['dominio_remitente']
        domains[domain] = domains.get(domain, 0) + 1
    
    if domains:
        print(f"\n🏢 Dominios más frecuentes:")
        sorted_domains = sorted(domains.items(), key=lambda x: x[1], reverse=True)
        for domain, count in sorted_domains[:5]:
            print(f"   {domain}: {count} correos")

def main():
    """Función principal"""
    try:
        # Elegir método
        method = choose_authentication_method()
        
        if method == 'imap':
            success = run_imap_method()
        else:  # graph
            success = run_graph_method()
        
        if not success:
            print(f"\n🔄 ¿Quieres probar el otro método?")
            retry = input("Escribe 's' para intentar de nuevo o cualquier tecla para salir: ").strip().lower()
            
            if retry in ['s', 'si', 'sí', 'yes', 'y']:
                # Probar el otro método
                other_method = 'graph' if method == 'imap' else 'imap'
                
                if other_method == 'imap':
                    run_imap_method()
                else:
                    run_graph_method()
        
    except KeyboardInterrupt:
        print("\n\n🛑 Aplicación cancelada por el usuario.")
    except Exception as e:
        print(f"\n❌ Error inesperado: {str(e)}")
        print("💡 Intenta ejecutar de nuevo o reporta el problema.")

if __name__ == "__main__":
    main()