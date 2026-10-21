from playwright.sync_api import sync_playwright


def conectar():

    p = sync_playwright().start()

    browser = p.chromium.launch(
        headless=False
    )

    context = browser.new_context(
        storage_state="sesion.json"
    )

    page = context.new_page()

    page.goto(
        "https://outlook.office.com/mail/"
    )

    page.wait_for_timeout(5000)

    for _ in range(20):

        page.mouse.wheel(0, 5000)

        page.wait_for_timeout(1000)

    return p, browser, context, page