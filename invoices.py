import os
import sys
import time
import base64
import getpass
import importlib.util
import subprocess
import webbrowser

# Package name on PyPI -> module name used for import
REQUIRED_PACKAGES = {
    "pandas": "pandas",
    "openpyxl": "openpyxl",
    "requests": "requests",
}

SCRIPT_DIR = os.path.dirname(os.path.abspath(sys.argv[0]))
ENV_FILE = os.path.join(SCRIPT_DIR, ".env")

DEFAULT_CONFIG = {
    "FAKTUROID_SLUG": "benalix",
    "SUBJECT_ID": "25163046",
    "INVOICE_FOLDER": os.path.join(os.path.expanduser("~"), "Desktop", "Bednar_fakturoid"),
}
SECRET_KEYS = ["FAKTUROID_CLIENT_ID", "FAKTUROID_CLIENT_SECRET", "CONTACT_EMAIL"]

APP_URL = "https://app.fakturoid.cz"
API_URL = f"{APP_URL}/api/v3"


def ensure_dependencies():
    # A PyInstaller build already bundles its dependencies
    if getattr(sys, "frozen", False):
        return
    missing = [pkg for pkg, module in REQUIRED_PACKAGES.items() if importlib.util.find_spec(module) is None]
    if not missing:
        return
    print(f"Chybějící knihovny: {', '.join(missing)}. Instaluji...")
    command = [sys.executable, "-m", "pip", "install", *missing]
    # pip refuses --user inside a virtualenv
    if sys.prefix == sys.base_prefix:
        command.append("--user")
    try:
        subprocess.check_call(command)
    except Exception as e:
        raise Exception(
            f"Instalace knihoven selhala ({e}). "
            f"Nainstalujte je ručně: {sys.executable} -m pip install {' '.join(missing)}"
        )
    importlib.invalidate_caches()


def load_env_file():
    values = {}
    if os.path.exists(ENV_FILE):
        with open(ENV_FILE, encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line or line.startswith("#") or "=" not in line:
                    continue
                key, value = line.split("=", 1)
                values[key.strip()] = value.strip().strip('"').strip("'")
    return values


def load_config():
    file_values = load_env_file()
    config = {}
    for key, default in DEFAULT_CONFIG.items():
        config[key] = os.getenv(key) or file_values.get(key) or default
    for key in SECRET_KEYS:
        config[key] = os.getenv(key) or file_values.get(key)

    # First run: ask for the credentials and remember them in .env next to the script
    missing = [key for key in SECRET_KEYS if not config[key]]
    if missing:
        print("Chybí přihlašovací údaje k Fakturoidu (Nastavení > Uživatelský účet > API v3).")
        prompts = {
            "FAKTUROID_CLIENT_ID": ("Client ID: ", input),
            "FAKTUROID_CLIENT_SECRET": ("Client Secret: ", getpass.getpass),
            "CONTACT_EMAIL": ("Kontaktní e-mail pro API: ", input),
        }
        for key in missing:
            label, ask = prompts[key]
            value = ask(label).strip()
            while not value:
                value = ask(label).strip()
            config[key] = value
            file_values[key] = value
        with open(ENV_FILE, "w", encoding="utf-8") as f:
            for key, value in file_values.items():
                f.write(f"{key}={value}\n")
        print(f"Údaje uloženy do {ENV_FILE}\n")

    config["USER_AGENT"] = f"Bednar_import_faktur ({config['CONTACT_EMAIL']})"
    return config


def main():
    try:
        ensure_dependencies()
        config = load_config()
        df = read_file(config["INVOICE_FOLDER"])
        items = extract_data(df)
        invoice_data = create_invoice_data(items, config)
        invoice_id = send_invoice_to_fakturoid(invoice_data, config)
        webbrowser.open_new_tab(f"{APP_URL}/{config['FAKTUROID_SLUG']}/invoices/{invoice_id}")
    except Exception as e:
        print(f"CHYBA: {e}")
        time.sleep(10)
        sys.exit(1)


def get_token(config):
    import requests

    auth_string = f"{config['FAKTUROID_CLIENT_ID']}:{config['FAKTUROID_CLIENT_SECRET']}"
    base64_encoded = base64.b64encode(auth_string.encode()).decode()

    headers = {
        "User-Agent": config["USER_AGENT"],
        "Content-Type": "application/x-www-form-urlencoded",
        "Accept": "application/json",
        "Authorization": f"Basic {base64_encoded}"
    }
    body = {"grant_type": "client_credentials"}

    response = requests.post(f"{API_URL}/oauth/token", data=body, headers=headers)
    if response.status_code == 200:
        return response.json().get("access_token")
    else:
        raise Exception(f"Nastala chyba při získávání tokenu: {response.text}")


def send_invoice_to_fakturoid(invoice_data, config):
    import requests

    try:
        token = get_token(config)
        headers = {
            "User-Agent": config["USER_AGENT"],
            "Content-Type": "application/json",
            "Authorization": f"Bearer {token}"
        }

        url = f"{API_URL}/accounts/{config['FAKTUROID_SLUG']}/invoices.json"
        response = requests.post(url, json=invoice_data, headers=headers)
        if response.status_code == 201:
            return response.json().get("id")
        else:
            raise Exception(response.text)
    except Exception as e:
        raise Exception(f"Nastala chyba při odesílání faktury: {e}")


def group_items_by_order_number(items):
    items_group_by_order_number = {}
    for item in items:
        items_group_by_order_number.setdefault(item["order_number"], []).append(item)
    return items_group_by_order_number


def create_invoice_data(items, config):
    try:
        invoice_number = input("Zadejte číslo faktury: ")
        while not invoice_number.isdigit():
            invoice_number = input("Není číslo, zadejte znovu: ")

        invoice_note = input("Zadejte číslo dodacího listu: ")
        if invoice_note:
            invoice_note = f"Dodací list č. {invoice_note} \nFakturujeme vám zboží dle objednávky:"
        else:
            invoice_note = "Fakturujeme vám zboží dle objednávky:"

        items_group_by_order_number = group_items_by_order_number(items)

        invoice_data = {
            "number": invoice_number,
            "note": invoice_note,
            "subject_id": int(config["SUBJECT_ID"]),
            "issued_on": items[0]["delivery_date"],
            "lines": []
        }

        for order_number, grouped_items in items_group_by_order_number.items():
            invoice_data["lines"].append({
                "name": f"\nObjednávka č. {order_number}",
                "quantity": 1,
                "unit_price": 0,
                "vat_rate": "0"
            })
            for item in grouped_items:
                invoice_data["lines"].append({
                    "name": item["invoice_item_name"],
                    "quantity": item["invoice_item_quantity"],
                    "unit_price": item["invoice_item_price"],
                    "vat_rate": "21"
                })
        return invoice_data
    except Exception as e:
        raise Exception(f"Nastala chyba při vytváření dat faktury: {e}")


def extract_data(df):
    items = []
    try:
        if df.empty:
            raise Exception("Soubor je prázdný")

        for _, row in df.iterrows():
            item = {
                "order_number": str(row["Potvrzená objednávka"]),
                "delivery_date": row["Datum návozu"].strftime('%Y-%m-%d'),
                "invoice_item_name": f"{row['Číslo produktu']} - {row['Název produktu']}",
                "invoice_item_quantity": str(row["Přijaté množství"]),
                "invoice_item_price": str(round(float(row["Cena"]), 2))
            }
            if int(item["invoice_item_quantity"]) > 0:
                items.append(item)

        if not items:
            raise Exception("V souboru nejsou žádné položky s kladným množstvím")

        if not check_delivery_date(items):
            raise Exception("Datumy závozu nejsou stejné")
    except Exception as e:
        raise Exception(f"Nastala chyba při extrakci dat: {e}")
    return items


def check_delivery_date(items):
    dates = sorted({item["delivery_date"] for item in items})
    if len(dates) > 1:
        print("\nV souboru jsou položky z různých dnů závozu. Prosím opravte soubor.")
        print(f"Nalezená data: {', '.join(dates)}\n")
        return False
    return True


def find_file(dir_path):
    if not os.path.exists(dir_path):
        print(f"Cesta {dir_path} neexistuje")
        os.makedirs(dir_path)
        print(f"Vytvořena cesta {dir_path}. Uložte sem soubor a spusťte program znovu.")
        raise Exception("Cesta neexistuje")
    for file in os.listdir(dir_path):
        if file.endswith('.xlsx'):
            return os.path.join(dir_path, file)
    raise Exception("Soubor nenalezen nebo je neplatný nebo prázdný")


def read_file(dir_path):
    import pandas as pd

    try:
        file_path = find_file(dir_path)
        df = pd.read_excel(file_path)
        df["Datum návozu"] = pd.to_datetime(df["Datum návozu"])
        return df
    except Exception as e:
        raise Exception(f"Nastala chyba při čtení souboru: {e}")


if __name__ == "__main__":
    main()
