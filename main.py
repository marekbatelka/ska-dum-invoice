import os
import datetime
import webbrowser  # Add this import at the top of the file
import config
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from ares_util.ares import call_ares
from qrplatba import QRPlatbaGenerator
from datetime import datetime, timedelta

# If modifying these scopes, delete the file token.json.
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets"]
INPUT_SPREADSHEET_ID = config.INPUT_SPREADSHEET_ID
INPUT_RANGE_NAME = config.INPUT_RANGE_NAME
INVOICE_SPREADSHEET_ID = config.INVOICE_SPREADSHEET_ID


def get_credentials():
    """Handles the Google API credentials."""
    creds = None
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)

    # If there are no valid credentials, request new ones
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
            except Exception as e:  # Catching all exceptions to reinitialize the flow
                print(
                    f"Could not refresh credentials: {e}. Requesting new credentials."
                )
                flow = InstalledAppFlow.from_client_secrets_file(
                    "credentials.json", SCOPES
                )
                creds = flow.run_local_server(port=0)
        else:
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
            creds = flow.run_local_server(port=0)

    # Save the credentials for the next run
    with open("token.json", "w") as token:
        token.write(creds.to_json())
    return creds


def read_form(sheet_service, creds):
    result = (
        sheet_service.values()
        .get(spreadsheetId=INPUT_SPREADSHEET_ID, range=INPUT_RANGE_NAME)
        .execute()
    )
    values = result.get("values", [])

    if not values:
        print("No data found.")
        return

    start_date = input("Datum startu: ")
    for row in values:
        google_form = {}
        if start_date in row[13]:
            google_form["timestamp"] = row[0]
            google_form["name"] = row[1]
            google_form["site"] = row[2]
            google_form["ico"] = row[3]
            google_form["street"] = row[4]
            google_form["city"] = row[5]
            google_form["psc"] = row[6]
            google_form["stat"] = row[7]
            google_form["firstnamesurname"] = row[8]
            google_form["nickname"] = row[9]
            google_form["email"] = row[10]
            google_form["phone"] = row[12]
            google_form["datecheckin"] = row[13]
            google_form["timecheckin"] = row[14]
            google_form["datecheckout"] = row[15]
            google_form["timecheckout"] = row[16]
            google_form["pax"] = row[17]
            print(
                f'{google_form["name"]} {google_form["firstnamesurname"]} - {google_form["nickname"]}'
            )

            confirm = input(
                f"jsou to oni? Pokud ano zmackni klavesu 'a' a potvrd Enter. "
            )
            if confirm == "a":
                mannights = input("kolik osobo-noci? ")
                pax = input("kolik osob? ")
                ico = google_form["ico"]
                data = call_ares(ico)
                if data == False:
                    raise ValueError(f"Ičo není vyplněno.")

                # Prepare the form_data for the invoice
                form_data = {
                    "issue_date": datetime.now().strftime("%Y-%m-%d"),
                    "due_date": (datetime.now() + timedelta(days=14)).strftime(
                        "%Y-%m-%d"
                    ),
                    "recipient_email":google_form["email"],
                    "recipient_name": google_form["name"],
                    "recipient_street": google_form["street"],
                    "recipient_city": google_form["city"],
                    "recipient_zip": google_form["psc"],
                    "ico": google_form["ico"],
                    "price": config.cena * int(mannights),
                    "invoice_number": "20258001",  # Example invoice number/
                    "datecheckin": google_form["datecheckin"],
                    "datecheckout": google_form["datecheckout"],
                    "mannights": mannights,
                    "pax_form": google_form["pax"],
                    "pax_input": pax,
                }
                return form_data
            else:
                print("ok, preskakuju")
    print("Nenasel jsem")
    return None


def generate_invoice(creds, form_data):
    """
    Generates an invoice by cloning the last sheet in the spreadsheet,
    renaming it, and using the name as the variable_symbol.
    Updates specific cells in the new sheet with provided data.
    Generates a QR code for payment and embeds it in the sheet.
    """
    try:
        # Build the Sheets API service
        sheet_service = build("sheets", "v4", credentials=creds).spreadsheets()

        # Get metadata of the spreadsheet to find the last sheet
        spreadsheet_metadata = sheet_service.get(spreadsheetId=INVOICE_SPREADSHEET_ID).execute()
        sheets = spreadsheet_metadata.get("sheets", [])
        last_sheet = sheets[-1]  # Get the last sheet
        last_sheet_id = last_sheet["properties"]["sheetId"]
        last_sheet_name = last_sheet["properties"]["title"]

        # Increment the name of the new sheet
        new_sheet_name = str(int(last_sheet_name) + 1)

        # Clone the last sheet and place it at the far right
        copy_request = {
            "requests": [
                {
                    "duplicateSheet": {
                        "sourceSheetId": last_sheet_id,
                        "newSheetName": new_sheet_name,
                        "insertSheetIndex": len(sheets),  # Place at the far right
                    }
                }
            ]
        }
        sheet_service.batchUpdate(
            spreadsheetId=INVOICE_SPREADSHEET_ID, body=copy_request
        ).execute()

        # Use the new sheet name as the variable_symbol
        form_data["variable_symbol"] = new_sheet_name

        # Prepare the data to update the new sheet
        updates = [
            {"range": f"'{new_sheet_name}'!F1", "values": [[new_sheet_name]]},
            {
                "range": f"'{new_sheet_name}'!F4",
                "values": [[datetime.now().strftime("%d.%m.%Y")]],
            },
            {
                "range": f"'{new_sheet_name}'!A21",
                "values": [[form_data["recipient_name"]]],
            },
            {
                "range": f"'{new_sheet_name}'!A22",
                "values": [[form_data["recipient_street"]]],
            },
            {
                "range": f"'{new_sheet_name}'!A23",
                "values": [[form_data["recipient_city"]]],
            },
            {
                "range": f"'{new_sheet_name}'!A24",
                "values": [[form_data["recipient_zip"]]],
            },
            {
                "range": f"'{new_sheet_name}'!A25",
                "values": [[f'IČ:{form_data["ico"]}']],
            },
            {"range": f"'{new_sheet_name}'!F30", "values": [[f'{form_data["price"]}']]},
            {
                "range": f"'{new_sheet_name}'!A30",
                "values": [
                    [
                        f"Fakturujeme vám pronájem skautské základny v termínu \nod {form_data['datecheckin']} do {form_data['datecheckout']} pro {form_data['pax_input']} osob"
                    ]
                ],
            },
        ]

        # Update the new sheet with the prepared data
        body = {"valueInputOption": "RAW", "data": updates}
        sheet_service.values().batchUpdate(
            spreadsheetId=INVOICE_SPREADSHEET_ID, body=body
        ).execute()

        # Generate QR code for payment
        due = datetime.now() + timedelta(days=7)
        price = form_data["price"]
        variable_symbol = new_sheet_name
        bank_account_number = config.bank_account_number
        generator = QRPlatbaGenerator(
            bank_account_number,
            int(price),
            x_vs=int(variable_symbol),
            message=config.qr_message,
            due_date=due,
        )
        img = generator.make_image()
        qr_file_svg = f"qr/{variable_symbol}.svg"
        os.makedirs("qr", exist_ok=True)  # Ensure the 'qr' directory exists
        img.save(qr_file_svg)
        img.save("qr_latest.svg")

        # Open the QR code in the default web browser
        webbrowser.open("qr_latest.svg")
        print(f"QR code opened in browser: qr_latest.svg")

        return new_sheet_name

    except HttpError as err:
        print(f"An error occurred: {err}")
        return None

def main():
    # TODO > handle exceptions
    # TODO > use config file
    creds = get_credentials()

    # Get input data from Google Sheet

    sheet_service = build("sheets", "v4", credentials=creds).spreadsheets()
    form_data = read_form(sheet_service, creds)
    if form_data is None:
        print("No valid form data found.")
        return

    # Generate invoice
    new_sheet_name = generate_invoice(creds, form_data)
    if new_sheet_name:
        print(f"Invoice generated with variable symbol: {new_sheet_name}")
        # Open the spreadsheet in the default web browser
        spreadsheet_url = (
            f"https://docs.google.com/spreadsheets/d/{INVOICE_SPREADSHEET_ID}"
        )
        webbrowser.open(spreadsheet_url)
        print(f"Spreadsheet opened in browser: {spreadsheet_url}")
    else:
        print("Failed to generate invoice.")

    # TODO > copy qr_latest.svg to the invoice sheet

    print("Replace the QR code in the invoice sheet with qr_latest.svg")

    # TODO > send email with invoice link


if __name__ == "__main__":
    main()
