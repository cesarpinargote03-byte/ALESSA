# api_findconversation.py
"""
Módulo y CLI para sincronizar, filtrar y descargar correos de Outlook
directamente a través de la API interna de Exchange / OWA (sin scroll manual ni UI visible).

Permite filtrar por:
- Año (ej: 2026, 2025 o 'todos')
- Categoría (ej: SOLICITUD, PENDIENTE, CULMINADO o 'todas')
- Remitente (ej: nombre o dirección de correo)
"""

import json
import os
import sys
from datetime import datetime
from pathlib import Path
from playwright.sync_api import sync_playwright

from exportar_excel import guardar_excel

CARPETA = Path(__file__).resolve().parent
ARCHIVO_CONVERSACIONES = CARPETA / "conversaciones.json"
ARCHIVO_PAYLOAD = CARPETA / "payload.json"
ARCHIVO_SESION = CARPETA / "sesion.json"
CARPETA_REPORTES = CARPETA / "Reportes"

URL_OWA = (
    "https://outlook.office.com/"
    "owa/service.svc?action=FindConversation&app=Mail&n=1"
)


def cargar_conversaciones():
    """Carga el diccionario acumulado de conversaciones indexado por ID."""
    todas = {}
    if ARCHIVO_CONVERSACIONES.exists():
        try:
            with open(ARCHIVO_CONVERSACIONES, "r", encoding="utf-8") as f:
                datos = json.load(f)
                if isinstance(datos, list):
                    for c in datos:
                        cid = c.get("ConversationId", {}).get("Id") or c.get("ConversationId")
                        if cid:
                            todas[str(cid)] = c
        except Exception as e:
            print(f"[WARN] Error al leer conversaciones previas: {e}")
    return todas


def guardar_conversaciones(todas_dict):
    """Guarda todas las conversaciones ordenadas por fecha en conversaciones.json."""
    lista = list(todas_dict.values())
    # Ordenar por fecha descendente (más recientes primero)
    lista.sort(key=lambda x: str(x.get("LastDeliveryTime") or ""), reverse=True)
    with open(ARCHIVO_CONVERSACIONES, "w", encoding="utf-8") as f:
        json.dump(lista, f, indent=4, ensure_ascii=False)
    return len(lista)


def construir_payload(instance_key=None, max_entries=50):
    """Genera el payload para la llamada a FindConversation de Exchange."""
    if ARCHIVO_PAYLOAD.exists():
        with open(ARCHIVO_PAYLOAD, "r", encoding="utf-8") as f:
            payload = json.load(f)
    else:
        # Payload de reserva si no existe payload.json
        payload = {
            "__type": "FindConversationJsonRequest:#Exchange",
            "Header": {
                "__type": "JsonRequestHeaders:#Exchange",
                "RequestServerVersion": "V2018_01_08",
                "TimeZoneContext": {
                    "__type": "TimeZoneContext:#Exchange",
                    "TimeZoneDefinition": {
                        "__type": "TimeZoneDefinitionType:#Exchange",
                        "Id": "SA Pacific Standard Time"
                    }
                }
            },
            "Body": {
                "ConversationShape": {
                    "__type": "ConversationResponseShape:#Exchange",
                    "BaseShape": "IdOnly",
                    "AdditionalProperties": [
                        {"__type": "PropertyUri:#Exchange", "FieldURI": "ConversationFlagDueDate"}
                    ]
                },
                "ShapeName": "ReactConversationListView",
                "Paging": {
                    "__type": "SeekToConditionPageView:#Exchange",
                    "BasePoint": "Beginning",
                    "MaxEntriesReturned": max_entries
                },
                "ViewFilter": "All",
                "SortOrder": [
                    {
                        "__type": "SortResults:#Exchange",
                        "Order": "Descending",
                        "Path": {"__type": "PropertyUri:#Exchange", "FieldURI": "ConversationLastDeliveryOrRenewTime"}
                    },
                    {
                        "__type": "SortResults:#Exchange",
                        "Order": "Descending",
                        "Path": {"__type": "PropertyUri:#Exchange", "FieldURI": "ConversationLastDeliveryTime"}
                    }
                ],
                "FocusedViewFilter": -1
            }
        }

    body = payload.setdefault("Body", {})
    paging = body.setdefault("Paging", {})
    paging["MaxEntriesReturned"] = max_entries

    if instance_key:
        paging["Condition"] = {
            "__type": "RestrictionType:#Exchange",
            "Item": {
                "__type": "IsEqualTo:#Exchange",
                "FieldURIOrConstant": {
                    "__type": "FieldURIOrConstantType:#Exchange",
                    "Item": {"__type": "Constant:#Exchange", "Value": str(instance_key)}
                },
                "Item": {
                    "__type": "PropertyUri:#Exchange",
                    "FieldURI": "ConversationInstanceKey"
                }
            }
        }
    else:
        # Primera página: sin condición de salto
        paging.pop("Condition", None)

    return payload


def sincronizar_outlook_api(max_paginas=5, silencio=False):
    """
    Se conecta headless a Outlook y descarga páginas de correos directamente por API.
    Retorna la cantidad de correos nuevos agregados.
    """
    if not ARCHIVO_SESION.exists():
        raise FileNotFoundError(f"No se encontró sesión activa en: {ARCHIVO_SESION}")

    todas = cargar_conversaciones()
    inicial = len(todas)
    if not silencio:
        print(f"[INFO] Base local actual: {inicial} conversaciones.")
        print(f"[INFO] Consultando API de Outlook (máx. {max_paginas} páginas)...")

    nuevos_total = 0
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        try:
            context = browser.new_context(storage_state=str(ARCHIVO_SESION))
            page = context.new_page()
            solicitud_inicial = {
                "encabezados": {},
                "estado": None,
                "respuesta": None,
            }

            def capturar_encabezados(request):
                if (
                    "service.svc?action=FindConversation" in request.url
                    and not solicitud_inicial["encabezados"]
                ):
                    solicitud_inicial["encabezados"].update(
                        {
                            nombre: valor
                            for nombre, valor in request.headers.items()
                            if nombre not in {
                                "host",
                                "cookie",
                                "content-length",
                                "connection",
                                "accept-encoding",
                            }
                        }
                    )

            def capturar_respuesta(response):
                if (
                    "service.svc?action=FindConversation" in response.url
                    and solicitud_inicial["estado"] is None
                ):
                    solicitud_inicial["estado"] = response.status
                    solicitud_inicial["respuesta"] = response

            page.on("request", capturar_encabezados)
            page.on("response", capturar_respuesta)

            page.goto("https://outlook.office.com/mail/", wait_until="domcontentloaded")
            for _ in range(20):
                if solicitud_inicial["estado"] is not None:
                    break
                page.wait_for_timeout(1000)

            if solicitud_inicial["estado"] is None:
                raise RuntimeError(
                    "Outlook no generó la solicitud FindConversation de la bandeja. "
                    "Confirma que la sesión abre la bandeja y vuelve a intentarlo."
                )
            if solicitud_inicial["estado"] == 401:
                raise RuntimeError(
                    "Outlook rechazó la sesión (HTTP 401). "
                    "Renueva la sesión con 'python login.py'."
                )
            if solicitud_inicial["estado"] != 200:
                raise RuntimeError(
                    "La solicitud inicial de Outlook devolvió "
                    f"HTTP {solicitud_inicial['estado']}."
                )
            if "authorization" not in solicitud_inicial["encabezados"]:
                raise RuntimeError(
                    "Outlook no proporcionó el encabezado de autenticación requerido."
                )
            datos_iniciales = solicitud_inicial["respuesta"].json()

            instance_key_actual = None

            for num_pagina in range(1, max_paginas + 1):
                if num_pagina == 1:
                    data = datos_iniciales
                else:
                    payload = construir_payload(instance_key=instance_key_actual)
                    response = page.request.post(
                        URL_OWA,
                        data=payload,
                        headers=solicitud_inicial["encabezados"],
                    )
                    if response.status != 200:
                        if response.status == 401:
                            detalle = (
                                "Outlook rechazó la autenticación (HTTP 401). "
                                "Renueva la sesión con 'python login.py'."
                            )
                        else:
                            detalle = f"Outlook devolvió HTTP {response.status}."
                        raise RuntimeError(
                            f"{detalle} Falló la sincronización en la página {num_pagina}."
                        )
                    data = response.json()
                body = data.get("Body", {})
                conversaciones = body.get("Conversations", [])
                if not conversaciones:
                    if not silencio:
                        print(f"[INFO] No se devolvieron más conversaciones en página {num_pagina}.")
                    break

                nuevas_en_pagina = 0
                for c in conversaciones:
                    cid = c.get("ConversationId", {}).get("Id") or c.get("ConversationId")
                    if not cid:
                        continue
                    cid_str = str(cid)
                    if cid_str not in todas:
                        nuevas_en_pagina += 1
                    todas[cid_str] = c

                nuevos_total += nuevas_en_pagina
                if not silencio:
                    print(f"  > Página {num_pagina}: {len(conversaciones)} recibidas (+{nuevas_en_pagina} nuevas)")

                # Preparar el salto para la siguiente página
                ultimo = conversaciones[-1]
                instance_key_actual = ultimo.get("InstanceKey")
                if not instance_key_actual:
                    break
        finally:
            browser.close()

    guardar_conversaciones(todas)
    if not silencio:
        print(f"[OK] Sincronización completada: +{nuevos_total} nuevos. Total acumulado: {len(todas)}")
    return nuevos_total


def aplicar_filtros(conversaciones, anio=None, categoria=None, remitente=None):
    """
    Filtra una lista de conversaciones por Año, Categoría y Remitente.
    """
    resultados = []
    anio_str = str(anio).strip() if anio else ""
    cat_str = str(categoria).strip().upper() if categoria else ""
    rem_str = str(remitente).strip().lower() if remitente else ""

    for c in conversaciones:
        # 1. Filtro por Año
        if anio_str and anio_str.lower() != "todos":
            fecha = str(c.get("LastDeliveryTime") or "")
            if not fecha.startswith(anio_str):
                continue

        # 2. Filtro por Categoría
        if cat_str and cat_str != "TODAS":
            cats = [str(cat).upper() for cat in c.get("Categories", [])]
            if not any(cat_str in cat for cat in cats):
                continue

        # 3. Filtro por Remitente
        if rem_str and rem_str != "todos":
            last_sender = c.get("LastSender", {})
            mailbox = last_sender.get("Mailbox", {}) if isinstance(last_sender, dict) else {}
            nombre = str(mailbox.get("Name") or "").lower()
            email = str(mailbox.get("EmailAddress") or "").lower()
            participantes = " ".join(str(p) for p in c.get("UniqueSenders", [])).lower()
            texto_remitentes = f"{nombre} {email} {participantes}"
            if rem_str not in texto_remitentes:
                continue

        resultados.append(c)

    return resultados


def api_descargar(
    anio=None,
    categoria=None,
    remitente=None,
    sincronizar=False,
    max_paginas=5,
    archivo_salida=None,
):
    """
    Función de nivel de API:
    - Opcionalmente sincroniza los correos más recientes vía API.
    - Carga y filtra las conversaciones por los criterios solicitados.
    - Exporta un archivo Excel con formato y datos completos.
    Retorna un diccionario con el resultado y metadata.
    """
    if sincronizar:
        sincronizar_outlook_api(max_paginas=max_paginas, silencio=False)

    todas_dict = cargar_conversaciones()
    lista = list(todas_dict.values())
    lista.sort(key=lambda x: str(x.get("LastDeliveryTime") or ""), reverse=True)

    filtradas = aplicar_filtros(lista, anio=anio, categoria=categoria, remitente=remitente)

    # Determinar nombre del archivo de salida
    if not archivo_salida:
        CARPETA_REPORTES.mkdir(parents=True, exist_ok=True)
        partes_nombre = ["Correos"]
        if categoria and str(categoria).upper() != "TODAS":
            partes_nombre.append(str(categoria).strip().replace(" ", "_"))
        if anio and str(anio).lower() != "todos":
            partes_nombre.append(str(anio).strip())
        if remitente and str(remitente).lower() != "todos":
            rem_limpio = "".join(c for c in str(remitente) if c.isalnum() or c in ("_", "-"))[:15]
            if rem_limpio:
                partes_nombre.append(rem_limpio)
        nombre_archivo = "_".join(partes_nombre) + ".xlsx"
        ruta_salida = CARPETA_REPORTES / nombre_archivo
    else:
        ruta_salida = Path(archivo_salida)

    guardar_excel(filtradas, ruta_salida, nombre_hoja="Correos Filtrados")

    return {
        "status": "success",
        "total_disponibles": len(lista),
        "total_filtrados": len(filtradas),
        "criterios": {
            "anio": anio or "todos",
            "categoria": categoria or "todas",
            "remitente": remitente or "todos",
        },
        "archivo_salida": str(ruta_salida),
    }


def menu_interactivo():
    """Interfaz interactiva por consola para el usuario de Helpdesk."""
    print("\n" + "=" * 65)
    print("      OUTLOOK DOWNLOADER API - CONSULTA Y DESCARGA")
    print("=" * 65)

    todas = cargar_conversaciones()
    print(f"\n[BD LOCAL] Tienes actualmente {len(todas)} conversaciones guardadas.")

    # 1. Preguntar si sincronizar
    resp_sync = input("\n¿Deseas sincronizar los correos más recientes de Outlook por API? (S/N) [N]: ").strip().upper()
    if resp_sync == "S":
        try:
            paginas_input = input("¿Cuántas páginas consultar? (cada una trae ~50 correos) [5]: ").strip()
            max_pags = int(paginas_input) if paginas_input.isdigit() else 5
            sincronizar_outlook_api(max_paginas=max_pags, silencio=False)
        except Exception as e:
            print(f"[ERROR] Error durante la sincronización: {e}")
            return

    # 2. Criterios de filtrado
    print("\n--- CRITERIOS DE DESCARGA ---")
    print("(Presiona ENTER para omitir un criterio y traer todos)")

    anio = input("\n1. Año a filtrar (ej. 2026, 2025 o Enter=todos): ").strip()
    categoria = input("2. Categoría a filtrar (ej. SOLICITUD, PENDIENTE, CULMINADO o Enter=todas): ").strip()
    remitente = input("3. Remitente a filtrar (nombre o correo, ej. Jorge, @alessa / Enter=todos): ").strip()

    print("\nProcesando y generando reporte...")
    resultado = api_descargar(
        anio=anio or None,
        categoria=categoria or None,
        remitente=remitente or None,
        sincronizar=False,
    )

    total = resultado["total_filtrados"]
    archivo = resultado["archivo_salida"]

    print("\n" + "=" * 65)
    print(f"[RESULTADO] Se encontraron {total} correos que cumplen tus criterios.")
    print(f"[EXCEL GENERADO] {archivo}")
    print("=" * 65)

    if total > 0 and os.name == "nt":
        abrir = input("\n¿Deseas abrir el archivo Excel ahora? (S/N) [S]: ").strip().upper()
        if abrir != "N":
            try:
                os.startfile(archivo)
                print("[OK] Abriendo Excel...")
            except Exception as e:
                print(f"[WARN] No se pudo abrir automáticamente: {e}")


if __name__ == "__main__":
    menu_interactivo()
