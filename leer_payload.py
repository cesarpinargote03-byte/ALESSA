import json

with open(
    "payload.json",
    "r",
    encoding="utf-8"
) as f:

    payload = json.load(f)

print(
    payload.keys()
)