# capturar_findconversation.py

import json
import os
import threading

from playwright.sync_api import sync_playwright


todas = {}
cerrando = False
enter_presionado = threading.Event()

# Cargar conversaciones existentes para no sobreescribir ni perder historial
if os.path.exists("conversaciones.json"):
    try:
        with open("conversaciones.json", "r", encoding="utf-8") as f:
            existentes = json.load(f)
            if isinstance(existentes, list):
                for c in existentes:
                    cid = c.get("ConversationId", {}).get("Id") or c.get("ConversationId")
                    if cid:
                        todas[str(cid)] = c
        print(f"[INFO] Se cargaron {len(todas)} conversaciones previas de conversaciones.json")
    except Exception as e:
        print(f"[WARN] Error al leer conversaciones previas: {e}")


def capturar(response):
    global todas, cerrando

    if cerrando or "FindConversation" not in response.url:
        return

    try:
        data = response.json()
    except Exception as e:
        print(f"[WARN] FindConversation no devolvió JSON válido: {response.status} {response.url} | {e}")
        return

    if not isinstance(data, dict):
        print(f"[WARN] FindConversation respondió con un JSON que no es un objeto: {type(data).__name__}")
        return

    body = data.get("Body")
    if not isinstance(body, dict):
        print(f"[WARN] FindConversation no tiene Body válido: {response.url}")
        return

    conversaciones = body.get("Conversations", [])
    if not isinstance(conversaciones, list):
        print(f"[WARN] FindConversation no tiene Conversations como lista: {response.url}")
        return

    nuevas = 0
    for c in conversaciones:
        try:
            conversation_id = c.get("ConversationId", {}).get("Id") or c.get("ConversationId")
            if not conversation_id:
                continue

            cid_str = str(conversation_id)
            if cid_str not in todas:
                nuevas += 1
            todas[cid_str] = c
        except Exception:
            pass

    print(f"FindConversation -> {len(conversaciones)} recibidas | +{nuevas} nuevas | Total acumulado: {len(todas)}")


def esperar_enter():
    try:
        input("\n>>> Presiona ENTER en esta consola cuando desees finalizar la captura...\n")
    finally:
        enter_presionado.set()


if not os.path.exists("sesion.json"):
    raise FileNotFoundError("No existe sesion.json. Guarda primero la sesión de Outlook en este directorio.")

with sync_playwright() as p:
    browser = None
    page = None
    try:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context(storage_state="sesion.json")
        page = context.new_page()

        page.on("response", capturar)

        print("\nAbriendo Outlook...")
        page.goto("https://outlook.office.com/mail/", wait_until="domcontentloaded")
        page.wait_for_timeout(5000)

        print("\n" + "=" * 60)
        print("CAPTURANDO CORREOS...")
        print("El bot hará scroll automático para cargar correos.")
        print("También puedes navegar o hacer scroll manualmente si lo deseas.")
        print("Presiona ENTER en esta consola para detener y guardar.")
        print("=" * 60 + "\n")

        threading.Thread(target=esperar_enter, daemon=True).start()

        # Ciclo con scroll automático suave mientras no se presione ENTER
        while not enter_presionado.is_set():
            # Scroll hacia abajo con la rueda del ratón en la zona central de la bandeja
            try:
                page.mouse.wheel(0, 1500)
            except Exception:
                pass
            page.wait_for_timeout(800)

        page.wait_for_timeout(1000)

        cerrando = True
        page.remove_listener("response", capturar)

        with open("conversaciones.json", "w", encoding="utf-8") as f:
            json.dump(list(todas.values()), f, indent=4, ensure_ascii=False)

        print(f"\n[OK] Guardadas {len(todas)} conversaciones en conversaciones.json")

        # Actualizar automáticamente el Excel
        try:
            print("\nActualizando Correos.xlsx...")
            from exportar_excel import main as exportar_a_excel
            exportar_a_excel()
        except Exception as e:
            print(f"[WARN] No se pudo exportar automáticamente a Excel: {e}")

    finally:
        if page is not None:
            try:
                page.close()
            except Exception:
                pass
        if browser is not None:
            try:
                browser.close()
            except Exception:
                pass