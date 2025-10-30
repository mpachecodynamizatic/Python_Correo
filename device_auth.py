# device_auth.py
"""
Autenticación usando Device Code Flow - método más compatible
"""

import requests
import time
import json
import webbrowser

class DeviceCodeAuthenticator:
    """Autenticador usando Device Code Flow (más compatible)"""
    
    def __init__(self):
        # Cliente público de Microsoft que funciona con Device Code Flow
        self.client_id = "14d82eec-204b-4c2f-b7e8-296a70dab67e"  # Microsoft Graph PowerShell
        self.tenant = "common"
        self.scopes = ["https://graph.microsoft.com/Mail.Read", "https://graph.microsoft.com/User.Read"]
        self.access_token = None
        self.user_info = None
    
    def authenticate(self):
        """Realiza autenticación usando Device Code Flow"""
        print("🔐 Iniciando autenticación con Microsoft Graph API...")
        print("📱 Este método usa 'Device Code' - muy compatible con todas las cuentas")
        print()
        
        try:
            return self._device_code_flow()
        except Exception as e:
            print(f"❌ Error durante la autenticación: {str(e)}")
            return False
    
    def _device_code_flow(self):
        """Implementa el flujo de Device Code"""
        try:
            # Paso 1: Solicitar device code
            print("🔄 Solicitando código de dispositivo...")
            
            device_code_url = f"https://login.microsoftonline.com/{self.tenant}/oauth2/v2.0/devicecode"
            
            device_data = {
                'client_id': self.client_id,
                'scope': ' '.join(self.scopes)
            }
            
            response = requests.post(device_code_url, data=device_data)
            
            if response.status_code != 200:
                print(f"❌ Error obteniendo device code: {response.status_code}")
                return False
            
            device_info = response.json()
            
            # Mostrar instrucciones al usuario
            print("="*60)
            print("📋 INSTRUCCIONES DE AUTENTICACIÓN")
            print("="*60)
            print(f"🌐 Ve a: {device_info['verification_uri']}")
            print(f"🔑 Ingresa este código: {device_info['user_code']}")
            print("👤 Inicia sesión con tu cuenta de Microsoft")
            print("✅ Autoriza la aplicación")
            print("="*60)
            
            # Abrir navegador automáticamente
            try:
                webbrowser.open(device_info['verification_uri'])
                print("🌐 Navegador abierto automáticamente")
            except:
                print("⚠️  No se pudo abrir el navegador automáticamente")
            
            print(f"\n⏳ Esperando autorización... (tienes {device_info['expires_in'] // 60} minutos)")
            print("💡 Presiona Ctrl+C para cancelar")
            
            # Paso 2: Esperar autorización
            return self._poll_for_token(device_info)
            
        except KeyboardInterrupt:
            print("\n🛑 Autenticación cancelada por el usuario")
            return False
        except Exception as e:
            print(f"❌ Error en device code flow: {str(e)}")
            return False
    
    def _poll_for_token(self, device_info):
        """Hace polling para obtener el token una vez autorizado"""
        token_url = f"https://login.microsoftonline.com/{self.tenant}/oauth2/v2.0/token"
        
        token_data = {
            'grant_type': 'urn:ietf:params:oauth:grant-type:device_code',
            'client_id': self.client_id,
            'device_code': device_info['device_code']
        }
        
        interval = device_info.get('interval', 5)
        expires_in = device_info['expires_in']
        start_time = time.time()
        
        while time.time() - start_time < expires_in:
            try:
                response = requests.post(token_url, data=token_data)
                result = response.json()
                
                if response.status_code == 200:
                    # ¡Éxito!
                    self.access_token = result['access_token']
                    print("\n🎉 ¡Autenticación exitosa!")
                    
                    # Obtener info del usuario
                    self._get_user_info()
                    return True
                
                elif result.get('error') == 'authorization_pending':
                    # Todavía esperando autorización
                    print(".", end="", flush=True)
                    time.sleep(interval)
                    continue
                
                elif result.get('error') == 'authorization_declined':
                    print("\n❌ Autorización denegada por el usuario")
                    return False
                
                elif result.get('error') == 'expired_token':
                    print("\n⏰ El código expiró. Inténtalo de nuevo.")
                    return False
                
                else:
                    print(f"\n❌ Error de autorización: {result.get('error_description', 'Error desconocido')}")
                    return False
                    
            except Exception as e:
                print(f"\n❌ Error en polling: {str(e)}")
                return False
        
        print("\n⏰ Tiempo agotado esperando autorización")
        return False
    
    def _get_user_info(self):
        """Obtiene información del usuario autenticado"""
        if not self.access_token:
            return
        
        try:
            headers = {'Authorization': f'Bearer {self.access_token}'}
            response = requests.get('https://graph.microsoft.com/v1.0/me', headers=headers)
            
            if response.status_code == 200:
                self.user_info = response.json()
                name = self.user_info.get('displayName', 'Usuario')
                email = self.user_info.get('mail') or self.user_info.get('userPrincipalName', 'No disponible')
                print(f"👤 Usuario autenticado: {name} ({email})")
            else:
                print("⚠️  No se pudo obtener información del usuario")
                
        except Exception as e:
            print(f"⚠️  Error obteniendo info del usuario: {str(e)}")
    
    def get_access_token(self):
        """Retorna el access token"""
        return self.access_token
    
    def get_user_info(self):
        """Retorna información del usuario"""
        if self.user_info:
            return {
                'nombre': self.user_info.get('displayName', 'Usuario'),
                'email': self.user_info.get('mail') or self.user_info.get('userPrincipalName', 'No disponible')
            }
        return None
    
    def is_authenticated(self):
        """Verifica si está autenticado"""
        return self.access_token is not None