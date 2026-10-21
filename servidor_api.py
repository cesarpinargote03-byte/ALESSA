# servidor_api.py
"""
Servidor FastAPI para consultar, filtrar y descargar correos de Outlook
mediante Swagger UI interactivo o endpoints REST.
"""

from pathlib import Path
from typing import Optional
from fastapi import FastAPI, Query, HTTPException
from fastapi.responses import RedirectResponse, FileResponse
from pydantic import BaseModel

from api_findconversation import (
    api_descargar,
    sincronizar_outlook_api,
    cargar_conversaciones,
    aplicar_filtros,
    ARCHIVO_SESION,
    ARCHIVO_CONVERSACIONES,
)

app = FastAPI(
    title="Outlook Downloader API",
    description=(
        "API para extraer, filtrar y descargar conversaciones de Outlook (Office 365) "
        "por Año, Categoría y Remitente con exportación directa a Excel."
    ),
    version="1.0.0",
)


@app.get("/", include_in_schema=False)
def inicio():
    """Redirige automáticamente a la documentación interactiva Swagger UI."""
    return RedirectResponse(url="/docs")


@app.get("/estado", tags=["Monitoreo"])
def obtener_estado():
    """Verifica el estado de la sesión, la base local y la cantidad de correos."""
    todas = cargar_conversaciones()
    lista = list(todas.values())
    fechas = [c.get("LastDeliveryTime") for c in lista if c.get("LastDeliveryTime")]

    return {
        "status": "online",
        "sesion_configurada": ARCHIVO_SESION.exists(),
        "total_conversaciones_guardadas": len(lista),
        "fecha_mas_reciente": max(fechas) if fechas else None,
        "fecha_mas_antigua": min(fechas) if fechas else None,
    }


@app.post("/sincronizar", tags=["Sincronización"])
def sincronizar_correos(
    max_paginas: int = Query(5, ge=1, le=50, description="Cantidad de páginas a consultar (~50 correos por página)")
):
    """
    Se conecta headless a Outlook para sincronizar los correos más recientes directamente vía API.
    Actualiza la base local y no sobreescribe correos previos.
    """
    if not ARCHIVO_SESION.exists():
        raise HTTPException(
            status_code=400,
            detail="No se encontró sesion.json. Debes montar o colocar tu sesión activa.",
        )

    try:
        nuevos = sincronizar_outlook_api(max_paginas=max_paginas, silencio=True)
        todas = cargar_conversaciones()
        return {
            "status": "success",
            "nuevos_descargados": nuevos,
            "total_acumulado": len(todas),
            "mensaje": f"Se sincronizaron {nuevos} correos nuevos correctamente.",
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error durante la sincronización: {str(e)}")


@app.get("/consultar", tags=["Consultas"])
def consultar_correos(
    anio: Optional[str] = Query(None, description="Año a filtrar (ej. '2026' o 'todos')"),
    categoria: Optional[str] = Query(None, description="Categoría a filtrar (ej. 'SOLICITUD', 'PENDIENTE', 'CULMINADO')"),
    remitente: Optional[str] = Query(None, description="Remitente a filtrar por nombre o correo (ej. 'Jorge', '@alessa')"),
    limite: int = Query(50, ge=1, le=500, description="Límite de resultados a retornar en JSON"),
):
    """
    Filtra los correos almacenados según tus criterios y los retorna en formato JSON.
    """
    todas = cargar_conversaciones()
    lista = list(todas.values())
    lista.sort(key=lambda x: str(x.get("LastDeliveryTime") or ""), reverse=True)

    filtrados = aplicar_filtros(lista, anio=anio, categoria=categoria, remitente=remitente)

    # Resumen limpio para el JSON
    items = []
    for c in filtrados[:limite]:
        rem = c.get("LastSender", {})
        mb = rem.get("Mailbox", {}) if isinstance(rem, dict) else {}
        cid = c.get("ConversationId", {}).get("Id") or c.get("ConversationId")
        items.append(
            {
                "id": str(cid),
                "asunto": c.get("ConversationTopic"),
                "fecha": c.get("LastDeliveryTime"),
                "categorias": c.get("Categories", []),
                "remitente": mb.get("Name"),
                "correo": mb.get("EmailAddress"),
                "resumen": c.get("Preview"),
                "tiene_adjuntos": bool(c.get("HasAttachments", False)),
            }
        )

    return {
        "total_coincidencias": len(filtrados),
        "mostrados": len(items),
        "criterios": {
            "anio": anio or "todos",
            "categoria": categoria or "todas",
            "remitente": remitente or "todos",
        },
        "resultados": items,
    }


@app.get("/descargar-excel", tags=["Descarga"])
def descargar_archivo_excel(
    anio: Optional[str] = Query(None, description="Año a filtrar (ej. '2026')"),
    categoria: Optional[str] = Query(None, description="Categoría a filtrar (ej. 'SOLICITUD')"),
    remitente: Optional[str] = Query(None, description="Remitente a filtrar (ej. 'Jorge')"),
    sincronizar_primero: bool = Query(False, description="Sincronizar con Outlook antes de exportar"),
):
    """
    Genera el archivo Excel filtrado con tablas, diseño y formatos profesionales,
    y lo envía directamente para descargarlo en tu navegador.
    """
    try:
        resultado = api_descargar(
            anio=anio,
            categoria=categoria,
            remitente=remitente,
            sincronizar=sincronizar_primero,
        )

        ruta_archivo = Path(resultado["archivo_salida"])
        if not ruta_archivo.exists():
            raise HTTPException(status_code=500, detail="No se pudo generar el archivo Excel.")

        return FileResponse(
            path=str(ruta_archivo),
            filename=ruta_archivo.name,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error generando Excel: {str(e)}")

