import os
import shutil
import openpyxl
import smartsheet
from flask import Flask, request, jsonify

app = Flask(__name__)

# === Configuration ===
API_KEY = os.getenv("SMARTSHEET_API_KEY")  # Use environment variable
SHEET_ID = int(os.getenv("SMARTSHEET_SHEET_ID"))  # Store as env variable

TEMPLATE_PATH = r"Updated Schedule.xlsx"  # Keep this file in your project folder
OUTPUT_DIRECTORY = r"property_folders"  # Directory to store generated files

# Initialize Smartsheet client
client = smartsheet.Smartsheet(API_KEY, api_base="https://api.smartsheet.eu/2.0")
client.errors_as_exceptions(True)  # Raise exceptions for better error handling


def create_property_file(row_data):
    """Generate Excel file for the specific row."""
    property_address = row_data.get("Property Address")
    if not property_address:
        return

    property_folder = os.path.join(OUTPUT_DIRECTORY, property_address)
    os.makedirs(property_folder, exist_ok=True)
    property_file_path = os.path.join(property_folder, f"{property_address}.xlsx")
    shutil.copy(TEMPLATE_PATH, property_file_path)

    wb = openpyxl.load_workbook(property_file_path)
    ws = wb.active

    # Mapping the data to the respective cell positions in the template
    mapping_positions = {
        "Property Address": "B3",
        "Local authority": "B5",
        "EPC Score ( Rd SAP)": "B6",
        "Tenure": "B7",
    }

    for key, cell_ref in mapping_positions.items():
        if key in row_data and row_data[key]:
            ws[cell_ref] = row_data[key]

    wb.save(property_file_path)

    return property_file_path


def attach_excel_file_to_smartsheet(property_address, row_id):
    """Attach generated Excel file to corresponding Smartsheet row."""
    property_folder = os.path.join(OUTPUT_DIRECTORY, property_address)
    excel_file_path = os.path.join(property_folder, f"{property_address}.xlsx")

    if os.path.exists(excel_file_path):
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
    else:
        print(f"❌ Excel file not found: {excel_file_path}")


@app.route("/webhook", methods=["POST", "GET"])
def webhook_listener():
    """Handles Smartsheet webhook verification and events."""

    if request.method == "GET":
        challenge = request.args.get("smartsheetHookChallenge")
        if challenge:
            print(f"✅ Responding to Smartsheet verification with challenge: {challenge}")
            return challenge, 200  # ✅ Send back challenge string as plain text

        return "Webhook is running!", 200  # Fallback response

    elif request.method == "POST":
        data = request.get_json()
        print("📥 Webhook received!", data)

        # Extract the event details from the webhook payload
        for event in data.get('events', []):
            row_id = event.get('rowId')
            changed_columns = event.get('changedColumns', [])

            # Check if the "Check Box" column was changed
            for column in changed_columns:
                if column.get('columnTitle') == 'Check Box' and column.get('newValue') == 'TRUE':
                    # Only process this row if the checkbox was checked
                    print(f"✅ 'Check Box' checked for row ID: {row_id}")

                    # Fetch the specific row data from Smartsheet
                    try:
                        row = client.Sheets.get_row(SHEET_ID, row_id)
                        row_data = {cell.column_title: cell.value for cell in row.cells}

                        # If "Check Box" is checked, create the Excel file and attach it
                        if row_data.get("Check Box") is True:
                            # Create the file for this row
                            file_path = create_property_file(row_data)
                            # Attach the file to Smartsheet
                            attach_excel_file_to_smartsheet(row_data["Property Address"], row_id)
                            
                            return jsonify({"message": "File updated & attached!"}), 200
                        else:
                            return jsonify({"message": "Checkbox unchecked, no action taken."}), 200
                    except smartsheet.exceptions.ApiError as e:
                        print(f"❌ Error fetching row data: {e}")
                        return jsonify({"message": "Error processing row."}), 500

        return jsonify({"message": "No relevant events found."}), 400


@app.route("/", methods=["GET"])
def home():
    return "✅ Smartsheet Automation is Running!", 200


if __name__ == "__main__":
    app.run(debug=True)
