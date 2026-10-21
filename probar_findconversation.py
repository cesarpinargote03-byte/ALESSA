# probar_findconversation.py

import json
from playwright.sync_api import sync_playwright

URL = (
    "https://outlook.cloud.microsoft/"
    "owa/service.svc?action=FindConversation&app=Mail&n=1"
)

with open("payload.json", "r", encoding="utf-8") as f:
    payload = json.load(f)

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

    response = page.request.post(
        URL,
        data=payload
    )

    print(
        "STATUS:",
        response.status
    )

    texto = response.text()

    print("\nRESPUESTA:\n")
    print(texto[:3000])

    if response.status != 200:

        print("\nERROR DE AUTENTICACION")
        browser.close()
        quit()

    data = response.json()

    conversaciones = data["Body"]["Conversations"]

    print(
        f"\nCONVERSACIONES: {len(conversaciones)}"
    )

    print("\nPRIMEROS 5\n")

    for c in conversaciones[:5]:

        categorias = ", ".join(
            c.get("Categories", [])
        )

        remitente = (
            c.get("LastSender", {})
            .get("Mailbox", {})
            .get("Name", "")
        )

        correo = (
            c.get("LastSender", {})
            .get("Mailbox", {})
            .get("EmailAddress", "")
        )

        print("-" * 60)

        print(
            "Fecha:",
            c.get("LastDeliveryTime", "")
        )

        print(
            "Categoria:",
            categorias
        )

        print(
            "Remitente:",
            remitente
        )

        print(
            "Correo:",
            correo
        )

        print(
            "Asunto:",
            c.get("ConversationTopic", "")
        )

        print(
            "Adjuntos:",
            c.get("HasAttachments", False)
        )

    browser.close()