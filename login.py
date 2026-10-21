# login.py
"""
Script para iniciar sesión en Outlook Web y generar o renovar el archivo sesion.json.
Úsalo cada vez que caduque la sesión o cuando configures el bot en una máquina nueva.
"""

from pathlib import Path
from playwright.sync_api import sync_playwright

ARCHIVO_SESION = Path(__file__).resolve().parent / "sesion.json"


def iniciar_sesion():
    print("\n" + "=" * 65)
    print("      INICIO DE SESIÓN EN OUTLOOK WEB (GENERADOR DE SESION.JSON)")
    print("=" * 65)
    print("1. Se abrirá una ventana de navegador.")
    print("2. Ingresa tu correo y contraseña corporativa.")
    print("3. Completa el código de verificación (MFA / Authenticator si aplica).")
    print("4. Marca 'Sí' en '¿Desea mantener la sesión iniciada?'.")
    print("5. Una vez que veas tu BANDEJA DE ENTRADA, regresa aquí y presiona ENTER.")
    print("=" * 65 + "\n")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context()
        page = context.new_page()

        page.goto("https://outlook.office.com/mail/")

        input("\n>>> Presiona ENTER aquí cuando ya estés en tu bandeja de entrada...\n")

        # Guarda cookies y estado de autenticación completo
        context.storage_state(path=str(ARCHIVO_SESION))
        print(f"\n[OK] ¡Sesión capturada y guardada exitosamente en:\n     {ARCHIVO_SESION}")
        browser.close()


if __name__ == "__main__":
    iniciar_sesion()

