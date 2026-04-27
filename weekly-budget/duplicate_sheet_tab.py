"""Duplicates the 'Weekly Budget Template' tab in a Google Sheet using gcloud ADC."""

from typing import Any
import argparse

import google.auth
from googleapiclient.discovery import build


TEMPLATE_TAB_NAME = "Weekly Budget Template"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]


def get_service():
    creds, _ = google.auth.default(scopes=SCOPES)
    return build("sheets", "v4", credentials=creds)


def get_sheets(service, spreadsheet_id: str) -> list[dict[str, Any]]:
    result = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    return result["sheets"]


def find_sheet(sheets: list[dict[str, Any]], name: str) -> dict[str, Any]:
    for sheet in sheets:
        if sheet["properties"]["title"] == name:
            return sheet
    raise ValueError(f"Tab '{name}' not found in spreadsheet.")


def duplicate_sheet(service, spreadsheet_id: str, source_sheet_id: int, new_title: str) -> dict[str, Any]:
    body = {
        "requests": [
            {
                "duplicateSheet": {
                    "sourceSheetId": source_sheet_id,
                    "newSheetName": new_title,
                }
            }
        ]
    }
    return service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()


def main():
    parser = argparse.ArgumentParser(description="Duplicate the 'Weekly Budget Template' tab in a Google Sheet.")
    parser.add_argument("spreadsheet_id", help="The Google Sheets spreadsheet ID")
    parser.add_argument("--new-name", default=None, help="Name for the duplicated tab (default: auto-generated copy name)")
    args = parser.parse_args()

    service = get_service()
    sheets = get_sheets(service, args.spreadsheet_id)
    template = find_sheet(sheets, TEMPLATE_TAB_NAME)
    source_sheet_id = template["properties"]["sheetId"]

    if args.new_name:
        new_title = args.new_name
    else:
        existing_titles = {s["properties"]["title"] for s in sheets}
        new_title = f"Copy of {TEMPLATE_TAB_NAME}"
        counter = 2
        while new_title in existing_titles:
            new_title = f"Copy of {TEMPLATE_TAB_NAME} ({counter})"
            counter += 1

    result = duplicate_sheet(service, args.spreadsheet_id, source_sheet_id, new_title)
    duplicated = result["replies"][0]["duplicateSheet"]["properties"]
    print(f"Created tab '{duplicated['title']}' (sheetId={duplicated['sheetId']})")


if __name__ == "__main__":
    main()
