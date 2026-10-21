from playwright.sync_api import sync_playwright

with sync_playwright() as p:

    browser = p.chromium.launch()

    context = browser.new_context(
        storage_state="sesion.json"
    )

    cookies = context.cookies()

    print("\nCOOKIES\n")

    for c in cookies:
        print(
            c["name"],
            "=",
            c["value"][:50]
        )

    browser.close()