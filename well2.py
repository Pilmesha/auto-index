import requests
import time
from io import BytesIO
from openpyxl import load_workbook
import re
import os
# ================= CONFIG =================
TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")

DRIVE_ID = os.getenv("DRIVE_ID")
ITEM_ID = os.getenv("ITEM_ID")
POLL_INTERVAL = 60  # seconds
# ==========================================
def get_token():
    url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    data = {
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "grant_type": "client_credentials",
        "scope": "https://graph.microsoft.com/.default"
    }

    resp = requests.post(url, data=data).json()

    if "access_token" not in resp:
        raise Exception(f"Token error: {resp}")

    return resp["access_token"]


def download_excel(headers):
    url = f"https://graph.microsoft.com/v1.0/drives/{DRIVE_ID}/items/{ITEM_ID}/content"
    return requests.get(url, headers=headers).content


def upload_excel(headers, data):
    url = f"https://graph.microsoft.com/v1.0/drives/{DRIVE_ID}/items/{ITEM_ID}/content"
    requests.put(url, headers=headers, data=data)
ID_RE = re.compile(r"_(\d{4})$")

def assign_ids(file_bytes):
    wb = load_workbook(BytesIO(file_bytes))
    global_ws = wb["__GLOBAL__"]

    # 1️⃣ Read stored global counter
    stored_last = int(global_ws["B1"].value or 0)

    # 2️⃣ Scan workbook for existing suffixes
    max_found = 0

    for ws in wb.worksheets:
        if ws.title == "__GLOBAL__":
            continue

        for row in ws.iter_rows(min_row=2):
            name = row[1].value  # Column B
            if not isinstance(name, str):
                continue

            m = ID_RE.search(name)
            if m:
                max_found = max(max_found, int(m.group(1)))

    # 3️⃣ TRUE last_id = max of both
    last_id = max(stored_last, max_found)

    changed = False

    # 4️⃣ Assign IDs only where missing
    for ws in wb.worksheets:
        if ws.title == "__GLOBAL__":
            continue

        for row in ws.iter_rows(min_row=2):
            cell = row[1]  # Column B
            name = cell.value

            if not isinstance(name, str) or not name.strip():
                continue

            # Remove any trailing malformed suffix like _123 or _001_extra
            clean_name = re.sub(r"_(\d+).*?$", "", name)

            if ID_RE.search(name):
                continue

            last_id += 1
            cell.value = f"{clean_name}_{last_id:04d}"
            changed = True

    # 5️⃣ Persist correct global counter
    if changed:
        global_ws["B1"].value = last_id
        out = BytesIO()
        wb.save(out)
        out.seek(0)
        return out.read(), last_id

    return None, last_id

def get_last_modified(headers):
    url = f"https://graph.microsoft.com/v1.0/drives/{DRIVE_ID}/items/{ITEM_ID}"
    return requests.get(url, headers=headers).json()["lastModifiedDateTime"]


def main():
    token = get_token()
    headers = {"Authorization": f"Bearer {token}"}

    last_seen = None

    print("📡 Watching SharePoint file for changes...")

    while True:
        try:
            modified = get_last_modified(headers)

            if modified != last_seen:
                print("🔄 File changed, processing...")
                file_bytes = download_excel(headers)

                result, last_id = assign_ids(file_bytes)

                if result:
                    upload_excel(headers, result)
                    print(f"✅ IDs assigned. Last ID = {last_id}")
                else:
                    print("ℹ No new projects found")

                last_seen = modified

        except Exception as e:
            print("⚠ Error:", e)

        time.sleep(POLL_INTERVAL)


if __name__ == "__main__":
    main()