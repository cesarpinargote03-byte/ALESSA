from pathlib import Path

from outlook import conectar
from extractor import extraer_correos


def encabezado():

    print("\n" + "=" * 50)
    print("OUTLOOK DOWNLOADER v1.0")
    print("=" * 50)


def leer_texto(mensaje):

    return input(mensaje).strip()


def leer_entero(mensaje):

    while True:

        try:
            return int(input(mensaje).strip())
        except:
            print("Ingrese un número válido")


def seleccionar_carpeta():

    print("\nCARPETAS")

    print("1. Bandeja de entrada")
    print("2. Elementos enviados")
    print("3. Borradores")
    print("4. Archivo")
    print("5. Personalizada")

    opcion = leer_entero(
        "\nSeleccione una opción: "
    )

    if opcion == 1:
        return "Bandeja de entrada"

    elif opcion == 2:
        return "Elementos enviados"

    elif opcion == 3:
        return "Borradores"

    elif opcion == 4:
        return "Archivo"

    elif opcion == 5:
        return leer_texto(
            "Nombre carpeta: "
        )

    return "Bandeja de entrada"


def seleccionar_rango():

    print("\nRANGO DE FECHAS")

    print("1. Todo")
    print("2. Año específico")
    print("3. Desde/Hasta")
    print("4. Últimos N días")

    opcion = leer_entero(
        "\nSeleccione una opción: "
    )

    if opcion == 1:

        return {
            "tipo": "todo"
        }

    elif opcion == 2:

        anio = leer_texto(
            "Año: "
        )

        return {
            "tipo": "anio",
            "anio": anio
        }

    elif opcion == 3:

        desde = leer_texto(
            "Fecha desde (dd/mm/yyyy): "
        )

        hasta = leer_texto(
            "Fecha hasta (dd/mm/yyyy): "
        )

        return {
            "tipo": "rango",
            "desde": desde,
            "hasta": hasta
        }

    elif opcion == 4:

        dias = leer_entero(
            "Cantidad días: "
        )

        return {
            "tipo": "ultimos",
            "dias": dias
        }

    return {
        "tipo": "todo"
    }


def obtener_configuracion():

    carpeta = seleccionar_carpeta()

    categoria = leer_texto(
        "\nCategoría (Enter=todas): "
    )

    remitente = leer_texto(
        "Remitente (Enter=todos): "
    )

    asunto = leer_texto(
        "Asunto (Enter=todos): "
    )

    rango = seleccionar_rango()

    descargar_adjuntos = (
        leer_texto(
            "\nDescargar adjuntos (S/N): "
        ).upper() == "S"
    )

    exportar_excel = (
        leer_texto(
            "Exportar Excel (S/N): "
        ).upper() == "S"
    )

    exportar_json = (
        leer_texto(
            "Exportar JSON (S/N): "
        ).upper() == "S"
    )

    limite = leer_entero(
        "Cantidad máxima (0=sin límite): "
    )

    ruta_destino = leer_texto(
        "Ruta destino: "
    )

    return {

        "carpeta": carpeta,
        "categoria": categoria,
        "remitente": remitente,
        "asunto": asunto,
        "rango": rango,
        "descargar_adjuntos": descargar_adjuntos,
        "exportar_excel": exportar_excel,
        "exportar_json": exportar_json,
        "limite": limite,
        "ruta_destino": ruta_destino
    }


def crear_estructura(ruta):

    base = Path(ruta)

    (base / "Reportes").mkdir(
        parents=True,
        exist_ok=True
    )

    (base / "Adjuntos").mkdir(
        parents=True,
        exist_ok=True
    )

    (base / "Logs").mkdir(
        parents=True,
        exist_ok=True
    )

    return base

def aplicar_filtros(correos, config):

    resultado = []

    for correo in correos:

        if config["categoria"]:

            if correo["categoria"].upper() != config["categoria"].upper():
                continue

        if config["remitente"]:

            texto = (
                correo["remitente"]
                + " "
                + correo["email"]
            ).lower()

            if config["remitente"].lower() not in texto:
                continue

        if config["asunto"]:

            if config["asunto"].lower() not in correo["asunto"].lower():
                continue

        resultado.append(correo)

    return resultado

def main():

    encabezado()

    config = obtener_configuracion()

    crear_estructura(
        config["ruta_destino"]
    )

    print("\nCONFIGURACIÓN\n")

    for clave, valor in config.items():

        print(
            f"{clave}: {valor}"
        )

    print("\nConectando Outlook...")

    p, browser, context, page = conectar()

    correos = extraer_correos(page)
    correos = aplicar_filtros(correos, config)

    print(
        f"\nCorreos extraídos: {len(correos)}"
    )

    for correo in correos[:10]:

        print("\n-------------------")
        print(correo)

    browser.close()

    p.stop()

    print(
        "\nProceso finalizado."
    )


if __name__ == "__main__":
    main()