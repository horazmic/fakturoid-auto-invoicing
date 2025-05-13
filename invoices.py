import os
import pandas as pd
import requests
import base64
import time
import webbrowser
import sys
from dotenv import load_dotenv

load_dotenv() 

required_vars = ["CLIENT_ID", "CLIENT_SECRET", "SLUG", "USER_AGENT", "SUBJECT_ID", "INVOICE_FOLDER"]
for var in required_vars:
    if not os.getenv(var):
        raise EnvironmentError(f"Environment variable '{var}' is not set. Check your .env file.")

INVOICE_FOLDER = os.getenv("INVOICE_FOLDER")
CLIENT_ID = os.getenv("FAKTUROID_CLIENT_ID")
CLIENT_SECRET = os.getenv("FAKTUROID_CLIENT_SECRET")
SLUG = os.getenv("FAKTUROID_SLUG")
USER_AGENT = os.getenv("USER_AGENT")
API_URL = os.getenv("FAKTUROID_API_URL", "https://app.fakturoid.cz/api/v3")
TOKEN_URL = f"{API_URL}/oauth/token"
INVOICE_URL = f"{API_URL}/accounts/{SLUG}/invoices.json"
SUBJECT_ID = os.getenv("SUBJECT_ID")

def main():
    try:
        df = read_file()
        items = extract_data(df)
        invoice_data = create_invoice_data(items)
        invoice_id = send_invoice_to_fakturoid(invoice_data)
        webbrowser.open_new_tab(f"{API_URL}/{SLUG}/invoices/{invoice_id}")
    except Exception as e:
        print(f"ERROR: {e}")
        time.sleep(10)
        sys.exit(1)

def get_token():
    auth_string = f'{CLIENT_ID}:{CLIENT_SECRET}'
    base64_encoded = base64.b64encode(auth_string.encode()).decode()

    headers = {
        "User-Agent": USER_AGENT,
        "Content-Type": "application/x-www-form-urlencoded",
        "Accept": "application/json",
        "Authorization": f"Basic {base64_encoded}"
    }
    body = {"grant_type": "client_credentials"}

    response = requests.post(TOKEN_URL, data=body, headers=headers)
    if response.status_code == 200:
        return response.json().get("access_token")
    else:
        raise Exception(f"Error fetching token: {response.text}")

def send_invoice_to_fakturoid(invoice_data):
    try:
        token = get_token()
        headers = {
            "User-Agent": USER_AGENT,
            "Content-Type": "application/json",
            "Authorization": f"Bearer {token}"
        }

        response = requests.post(INVOICE_URL, json=invoice_data, headers=headers)
        if response.status_code == 201:
            return response.json().get("id")
        else:
            raise Exception(f"Error sending invoice: {response.text}")
    except Exception as e:
        raise Exception(f"Failed to send invoice: {e}")

def group_items_by_order_number(items):
    items_group_by_order_number = {}
    for item in items:
        items_group_by_order_number.setdefault(item["order_number"], []).append(item)
    return items_group_by_order_number

def create_invoice_data(items):
    try:
        invoice_number = input("Enter Invoice Number: ")
        while not invoice_number.isdigit():
            invoice_number = input("Invalid number, try again: ")

        invoice_note = input("Enter Delivery Note Number: ")
        invoice_note = f"Delivery Note No. {invoice_note} \nInvoicing based on order:" if invoice_note else "Invoicing based on order:"

        items_group_by_order_number = group_items_by_order_number(items)

        invoice_data = {
            "number": invoice_number,
            "note": invoice_note,
            "subject_id": SUBJECT_ID,
            "issued_on": items[0]["delivery_date"],
            "lines": []
        }

        for order_number, grouped_items in items_group_by_order_number.items():
            invoice_data["lines"].append({
                "name": f"\nOrder No. {order_number}",
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
        raise Exception(f"Failed to create invoice data: {e}")

def extract_data(df):
    items = []
    try:
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

        if df.empty:
            raise Exception("The Excel file is empty.")
            
        if not check_delivery_date(items):
            raise Exception("Delivery dates do not match")
    except Exception as e:
        raise Exception(f"Data extraction error: {e}")
    return items

def check_delivery_date(items):
    dates = {item["delivery_date"] for item in items}
    if len(dates) > 1:
        print("The file contains items from different delivery dates. Please fix the file.")
        return False
    return True

def find_file(dir_path):
    if not os.path.exists(dir_path):
        os.makedirs(dir_path)
        print(f"Path created: {dir_path}. Please place the file and rerun the program.")
        raise Exception("Directory not found")
    for file in os.listdir(dir_path):
        if file.endswith('.xlsx'):
            return os.path.join(dir_path, file)
    raise Exception("Valid file not found")

def read_file():
    try:
        dir_path = INVOICE_FOLDER
        file_path = find_file(dir_path)
        df = pd.read_excel(file_path)
        df["Datum návozu"] = pd.to_datetime(df["Datum návozu"])
        return df
    except Exception as e:
        raise Exception(f"File read error: {e}")

if __name__ == "__main__":
    main()
