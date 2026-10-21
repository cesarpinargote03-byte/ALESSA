# capturar_todo.py

from playwright.sync_api import sync_playwright
import json

todas = []


def capturar(response):

    global todas

    if "FindConversation" in response.url:

        try:

            data = response.json()

            conversaciones = (
                data["Body"]["Conversations"]
            )

            todas.extend(
                conversaciones
            )

            print(
                f"Total acumulado: {len(todas)}"
            )

        except:
            pass


with sync_playwright() as p:

    browser = p.chromium.launch(
        headless=False
    )

    context = browser.new_context(
        storage_state="sesion.json"
    )

    page = context.new_page()

    page.on(
        "response",
        capturar
    )

    page.goto(
        "https://outlook.office.com/mail/"
    )

    print(
        "\nAhora navega por Outlook."
    )

    print(
        "Haz scroll y espera."
    )

    input(
        "\nENTER para guardar..."
    )

    with open(
        "todas_las_conversaciones.json",
        "w",
        encoding="utf-8"
    ) as f:

        json.dump(
            todas,
            f,
            indent=4,
            ensure_ascii=False
        )

    print(
        f"\nGuardadas {len(todas)} conversaciones"
    )

    browser.close()