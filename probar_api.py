import json

with open(
    "cookies.json",
    "r",
    encoding="utf-8"
) as f:

    cookies = json.load(f)

print(
    f"Cookies encontradas: {len(cookies)}"
)

for c in cookies:

    print(
        c["name"]
    )