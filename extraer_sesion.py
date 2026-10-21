import json
from playwright.sync_api import sync_playwright

with sync_playwright() as p:

    browser = p.chromium.launch()

    context = browser.new_context(
        storage_state="sesion.json"
    )

    page = context.new_page()

    page.goto(
        "https://outlook.office.com/mail/"
    )

    page.wait_for_timeout(5000)

    print("URL:")
    print(page.url)

    print("\nCOOKIES:")
    cookies = context.cookies()

    for c in cookies:

        if c["name"] in [
            "ESTSAUTH",
            "ESTSAUTHPERSISTENT",
            "DefaultAnchorMailbox"
        ]:

            print(f"{c['name']} OK")

    browser.close()
    