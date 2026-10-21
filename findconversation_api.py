import json

# Leer payload original capturado desde Outlook

with open(
    "payload.json",
    "r",
    encoding="utf-8"
) as f:

    payload = json.load(f)

print("\n=== BODY ===\n")

body = payload["Body"]

for k in body.keys():

    print(k)

print("\n=== PARENTFOLDERID ===\n")

print(
    json.dumps(
        body["ParentFolderId"],
        indent=4,
        ensure_ascii=False
    )
)

print("\n=== PAGING ===\n")

print(
    json.dumps(
        body["Paging"],
        indent=4,
        ensure_ascii=False
    )
)