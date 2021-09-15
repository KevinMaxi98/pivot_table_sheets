from __future__ import print_function
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

SAMPLE_SPREADSHEET_ID = '19MlHV_OPcxEr1Vh8a53-jI46P0YsaNkUMpINDoEn8Tc'
SAMPLE_RANGE_NAME = 'Hoja 1'


def pivot_tables(service, spreadsheet_id, end_row_index):
    body = {
        'requests': [{
            'addSheet': {}
        }]
    }
    batch_update_response = service.spreadsheets() \
        .batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()
    target_sheet_id = batch_update_response.get('replies')[0] \
        .get('addSheet').get('properties').get('sheetId')
    target_sheet_title = batch_update_response.get('replies')[0] \
        .get('addSheet').get('properties').get('title')
    requests = [{
        'updateCells': {
            'rows': {
                'values': [
                    {
                        'pivotTable': {
                            'source': {
                                'sheetId': 0,
                                'startRowIndex': 0,
                                'startColumnIndex': 0,
                                'endRowIndex': end_row_index,
                                'endColumnIndex': 4
                            },
                            'rows': [
                                {
                                    'sourceColumnOffset': 0,
                                    'sortOrder': 'ASCENDING',
                                    'repeatHeadings': False,

                                },
                                {
                                    'sourceColumnOffset': 1,
                                    'sortOrder': 'ASCENDING',
                                    'repeatHeadings': False,

                                },

                            ],
                            'columns': [
                                {
                                    'sourceColumnOffset': 2,
                                    'sortOrder': 'ASCENDING',
                                    'repeatHeadings': False,

                                },
                            ],
                            'values': [
                                {
                                    'summarizeFunction': 'COUNTUNIQUE',
                                    'sourceColumnOffset': 1,
                                    'name': ' '
                                },
                            ],
                            'valueLayout': 'HORIZONTAL'
                        },
                    }
                ]
            },
            'start': {
                'sheetId': target_sheet_id,
                'rowIndex': 0,
                'columnIndex': 0
            },
            'fields': 'pivotTable'
        }
    }]
    body = {
        'requests': requests
    }
    response = service.spreadsheets() \
        .batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()
    result = service.spreadsheets().values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                                 range=target_sheet_title).execute()
    values = result.get('values', [])
    end_column_index = len(values[1])
    end_row_index1 = len(values)
    requests = [{
        'updateCells': {
            'rows': {
                'values': [
                    {
                        'pivotTable': {
                            'source': {
                                'sheetId': 0,
                                'startRowIndex': 0,
                                'startColumnIndex': 0,
                                'endRowIndex': end_row_index,
                                'endColumnIndex': 4
                            },
                            'rows': [
                                {
                                    'sourceColumnOffset': 0,
                                    'sortOrder': 'ASCENDING',
                                    'repeatHeadings': False,

                                },
                                {
                                    'sourceColumnOffset': 1,
                                    'sortOrder': 'ASCENDING',
                                    'repeatHeadings': False,

                                },

                            ],
                            'columns': [
                                {
                                    'sourceColumnOffset': 3,
                                    'sortOrder': 'ASCENDING',
                                    'repeatHeadings': False,
                                },
                            ],
                            'values': [
                                {
                                    'summarizeFunction': 'COUNTUNIQUE',
                                    'sourceColumnOffset': 1,
                                    'name': ' ',
                                },
                            ],
                            'valueLayout': 'HORIZONTAL'
                        },
                    }
                ]
            },
            'start': {
                'sheetId': target_sheet_id,
                'rowIndex': 0,
                'columnIndex': end_column_index,
            },
            'fields': 'pivotTable'
        }
    }]
    body = {
        'requests': requests
    }
    response = service.spreadsheets() \
        .batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()
    result = service.spreadsheets().values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                                 range=target_sheet_title).execute()
    values = result.get('values', [])
    end_column_index1 = len(values[1])
    requests = [{
        'cutPaste': {
            'source': {
                'sheetId': target_sheet_id,
                'startRowIndex': 0,
                'endRowIndex': end_row_index1,
                'startColumnIndex': 0,
                'endColumnIndex': end_column_index1
            },
            'destination': {
                'sheetId': target_sheet_id,
                'rowIndex': 0,
                'columnIndex': 0
            },
            'pasteType': 'PASTE_VALUES'
        }
    }, {
        "deleteDimension": {
            "range": {
                "sheetId": target_sheet_id,
                "dimension": "COLUMNS",
                "startIndex": end_column_index,
                "endIndex": end_column_index + 2
            }
        }
    }, {
        "repeatCell": {
            "range": {
                "sheetId": target_sheet_id,
                "startRowIndex": 1,
                "endRowIndex": end_row_index1,
                "startColumnIndex": 0,
                "endColumnIndex": end_column_index1-2
            },
            "cell": {
                "userEnteredFormat": {
                    "horizontalAlignment": "CENTER",
                }
            },
            "fields": "userEnteredFormat(horizontalAlignment)"
        }
    }]

    body = {
        'requests': requests
    }
    response = service.spreadsheets() \
        .batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()
    batch_update(service, SAMPLE_SPREADSHEET_ID, target_sheet_id, '^[0-9]', 'TRUE')
    batch_update(service, SAMPLE_SPREADSHEET_ID, target_sheet_id, '^(?![\s\S])', 'FALSE')
    requests = [{
        "mergeCells": {
            "range": {
                "sheetId": target_sheet_id,
                "startRowIndex": 0,
                "endRowIndex": 1,
                "startColumnIndex": 2,
                "endColumnIndex": end_column_index
            },
            "mergeType": "MERGE_ALL"
        }
    }, {
        "mergeCells": {
            "range": {
                "sheetId": target_sheet_id,
                "startRowIndex": 0,
                "endRowIndex": 1,
                "startColumnIndex": end_column_index,
                "endColumnIndex": end_column_index1 - 2
            },
            "mergeType": "MERGE_ALL"
        }
    }, {
        "repeatCell": {
            "range": {
                "sheetId": target_sheet_id,
                "startRowIndex": 0,
                "endRowIndex": 1
            },
            "cell": {
                "userEnteredFormat": {
                    "horizontalAlignment": "CENTER",
                }
            },
            "fields": "userEnteredFormat(horizontalAlignment)"
        }
    }]
    body = {
        'requests': requests
    }
    response = service.spreadsheets() \
        .batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()

    return target_sheet_id


def batch_update(service, spreadsheet_id, sheet_id, find, replacement):
    # [START sheets_batch_update]
    requests = []
    # Change the spreadsheet's title.
    # Find and replace text
    requests.append({
        'findReplace': {
            'find': find,
            'sheetId': sheet_id,
            'searchByRegex': True,
            'replacement': replacement,
        }
    })
    # Add additional requests (operations) ...

    body = {
        'requests': requests
    }
    response = service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body=body).execute()
    # find_replace_response = response.get('replies')[1].get('findReplace')
    # print('{0} replacements made.'.format(
    #     find_replace_response.get('occurrencesChanged')))
    # [END sheets_batch_update]
    return response


def main():
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    service = build('sheets', 'v4', credentials=creds)

    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                range=SAMPLE_RANGE_NAME).execute()
    values = result.get('values', [])
    end_row_index = len(values)
    sheet_id = pivot_tables(service, SAMPLE_SPREADSHEET_ID, end_row_index)



if __name__ == '__main__':
    main()
