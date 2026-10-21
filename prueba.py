from playwright.sync_api import sync_playwright

def capturar(response):

    url = response.url

    if "service.svc?action=" in url:

        try:

            accion = url.split("action=")[1].split("&")[0]
        except:
            accion = "DESCONOCIDA"

        print(
            f"{response.status} -> {accion}"
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

    page.wait_for_timeout(
        30000
    )

    input("ENTER")

    browser.close()