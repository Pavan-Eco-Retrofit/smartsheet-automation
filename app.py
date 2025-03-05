import os
import shutil
import openpyxl
import smartsheet
import pandas as pd
import requests
from flask import Flask, request, jsonify

app = Flask(__name__)

# === Configuration ===
API_KEY = os.getenv("SMARTSHEET_API_KEY")  # Use environment variable
SHEET_ID = int(os.getenv("SMARTSHEET_SHEET_ID"))  # Store as env variable
WEBHOOK_URL = "https://web-production-f336.up.railway.app/webhook"
TEMPLATE_PATH = r"Updated Schedule.xlsx"
OUTPUT_DIRECTORY = r"property_folders"

# Initialize Smartsheet client
client = smartsheet.Smartsheet(API_KEY, api_base="https://api.smartsheet.eu/2.0")  # Use EU API base if required
client.errors_as_exceptions(True)  # Raise exceptions for better error handling


def register_webhook():
    """Automatically register or update the webhook."""
    headers = {
        "Authorization": f"Bearer {API_KEY}",
        "Content-Type": "application/json"
    }
    payload = {
        "name": "Auto Property Webhook",
        "callbackUrl": WEBHOOK_URL,
        "scope": "sheet",
        "scopeObjectId": SHEET_ID,
        "events": ["UPDATE", "ADD_ROW", "DELETE_ROW"],
        "version": 1,
        "enabled": True
    }

    # Check if webhook exists
    response = requests.get("https://api.smartsheet.com/2.0/webhooks", headers=headers)
    if response.status_code == 200:
        webhooks = response.json().get("data", [])
        for webhook in webhooks:
            if webhook["callbackUrl"] == WEBHOOK_URL:
                print("✅ Webhook already exists. Updating...")
                update_payload = {"enabled": True}
                requests.put(f"https://api.smartsheet.com/2.0/webhooks/{webhook['id']}", json=update_payload, headers=headers)
                return

    # Register a new webhook if not found
    print("🚀 Registering new webhook...")
    response = requests.post("https://api.smartsheet.com/2.0/webhooks", json=payload, headers=headers)
    if response.status_code == 200:
        print("✅ Webhook registered successfully!")
    else:
        print("❌ Webhook registration failed:", response.text)


@app.route("/webhook", methods=["POST", "GET"])
def webhook_listener():
    """Handles Smartsheet webhook requests."""
    
    if request.method == "GET":
        challenge = request.args.get("smartsheetHookChallenge")
        if challenge:
            return challenge, 200  # ✅ Respond with the challenge string for verification!
        return "✅ Webhook is set up correctly!", 200

    elif request.method == "POST":
        data = request.get_json()
        print("📥 Webhook received!", data)

        # Proceed with processing webhook events...
        df, row_id_map = fetch_smartsheet_data()
        if df is not None and not df.empty:
            create_property_files(df)
            attach_excel_files_to_smartsheet(row_id_map)
            return jsonify({"message": "Files updated & attached!"}), 200
        else:
            return jsonify({"message": "No checked rows found!"}), 400


def fetch_smartsheet_data():
    """Fetch data from Smartsheet where 'Check Box' is checked."""
    try:
        sheet = client.Sheets.get_sheet(SHEET_ID)
        column_map = {col.id: col.title for col in sheet.columns}
        sheet_data, row_id_map = [], {}

        for row in sheet.rows:
            row_data = {column_map[cell.column_id]: cell.value for cell in row.cells if cell.value}
            if row_data.get("Check Box") is True:
                sheet_data.append(row_data)
                if "Property Address" in row_data:
                    row_id_map[row_data["Property Address"]] = row.id

        return pd.DataFrame(sheet_data), row_id_map

    except smartsheet.exceptions.ApiError as e:
        print(f"❌ Smartsheet API Error: {e}")
        return None, {}


def create_property_files(df):
    """Generate Excel files for each checked property row."""
    if not os.path.exists(OUTPUT_DIRECTORY):
        os.makedirs(OUTPUT_DIRECTORY)

    mapping_positions = {
        "Property Address": "B3",
        "Local authority": "B5",
        "EPC Score ( Rd SAP)": "B6",
        "Tenure": "B7",
    }

    for _, row in df.iterrows():
        property_address = row.get("Property Address")
        if not property_address:
            continue

        property_folder = os.path.join(OUTPUT_DIRECTORY, property_address)
        os.makedirs(property_folder, exist_ok=True)
        property_file_path = os.path.join(property_folder, f"{property_address}.xlsx")
        shutil.copy(TEMPLATE_PATH, property_file_path)

        wb = openpyxl.load_workbook(property_file_path)
        ws = wb.active

        for key, cell_ref in mapping_positions.items():
            if key in row and row[key]:
                ws[cell_ref] = row[key]

        wb.save(property_file_path)

    print("✅ Excel files generated successfully.")


def attach_excel_files_to_smartsheet(row_id_map):
    """Attach generated Excel files to corresponding Smartsheet rows."""
    for property_folder in os.listdir(OUTPUT_DIRECTORY):
        folder_path = os.path.join(OUTPUT_DIRECTORY, property_folder)

        if not os.path.isdir(folder_path):
            continue

        excel_files = [f for f in os.listdir(folder_path) if f.endswith(".xlsx") and not f.startswith("~$")]
        if not excel_files:
            continue

        excel_file_path = os.path.join(folder_path, excel_files[0])
        row_id = row_id_map.get(property_folder, None)

        if not row_id:
            continue

        print(f"📤 Attaching {excel_file_path} to Smartsheet row {row_id}")

        try:
            with open(excel_file_path, 'rb') as file:
                client.Attachments.attach_file_to_row(
                    SHEET_ID, row_id,
                    (os.path.basename(excel_file_path), file, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                )
            print(f"✅ Successfully attached: {excel_file_path}")
        except smartsheet.exceptions.ApiError as e:
            print(f"❌ Smartsheet API Error: {e}")

    print("🎉 All files attached successfully!")


@app.route("/", methods=["GET"])
def home():
    return "✅ Smartsheet Automation is Running!", 200


if __name__ == "__main__":
    register_webhook()  # Auto-register webhook when the app starts
    app.run(debug=True)
