import json
from playwright.sync_api import sync_playwright

with sync_playwright() as p:

    browser = p.chromium.launch()

    context = browser.new_context(
        storage_state="sesion.json"
    )

    cookies = context.cookies()

    with open(
        "cookies.json",
        "w",
        encoding="utf-8"
    ) as f:

        json.dump(
            cookies,
            f,
            indent=4
        )

    browser.close()

print(
    "cookies.json generado"
)