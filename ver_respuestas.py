# ver_respuestas.py

from playwright.sync_api import sync_playwright

def capturar(response):

    if "service.svc" in response.url:

        print(
            response.status,
            response.url
        )

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

    input(
        "\nENTER para salir..."
    )

    browser.close()