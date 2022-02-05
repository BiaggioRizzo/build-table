import os.path

from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow

# If modifying these SCOPES_SHEETS, delete the file token.json.
# More information in https://developers.google.com/sheets/api/guides/authorizing
SCOPES_SHEETS = ['https://www.googleapis.com/auth/spreadsheets']


def initialize_sheets():
    """
    ->Function to get/create token.

    return: creds: (String) Is a token.
    """

    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES_SHEETS)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES_SHEETS)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    return creds


def auto_resize(sheetId=0, dimension="COLUMNS", start=0, end=1):
    """
    -> Function to auto resize COLUMNS or ROWS
    :param sheetId: Is an identifier of a Sheet in a Google Spreadsheet.
    :param dimension: Use to adjust "COLUMNS" or "ROWS".
    :param start: The start of columns or row is inclusive.
    :param end: The end of columns or row is exclusive.
    """

    resize = [
        {
            "autoResizeDimensions": {
                "dimensions": {
                    "sheetId": sheetId,
                    "dimension": dimension,
                    "startIndex": start,
                    "endIndex": end
                }
            }
        }
    ]

    return resize


def color_cell(sheetId=0, red=0, green=0, blue=0, alpha=1, startRow=0, endRow=1,
               startColumn=0, endColumn=1, type="repeatCell"):
    """
    -> Function to color cell.
    ATTENTION: The function used color system RGBA, use the website #https://rgbacolorpicker.com/ for get your color.

    :param sheetId: (Number) Is an identifier of a Sheet in a Google Spreadsheet. 
    :param red: (Number) Channel red is value from 0 to 255.
    :param green: (Number) Channel green is value from 0 to 255.
    :param blue: (Number) Channel blue is value from 0 to 255.
    :param alpha: (Number) Channel alpha is value from 0 to 1. Represent level of transparencey/opacity.
    :param startRow: (Number) Start Row.
    :param endRow: (Number) End Row. Need a number greater than startRow.
    :param startColumn: (Number) Start Column
    :param endColumn: (Number) End Column. Need a number greater than startColumn.
    :param type:(String) You can use 'updateCells' for one cell or 'repeatCell' for multiple cells.
    """

    color = [
        {
            type:  # "updateCells": -> utilizado para atualizar única célula, repeatCell para altera várias
            {
                "rows":
                [
                    {
                        "values":
                        [
                            {
                                "userEnteredFormat":
                                {
                                    "backgroundColor": {
                                        "red": red / 255,
                                        "green": green / 255,
                                        "blue": blue / 255,
                                        "alpha": alpha
                                    },
                                },
                            }
                        ]
                    }
                ],
                "range":
                {
                    "sheetId": sheetId,
                    "startRowIndex": startRow,
                    "endRowIndex": endRow,
                    "startColumnIndex": startColumn,
                    "endColumnIndex": endColumn
                },
                "fields": "userEnteredFormat"
            }
        }
    ]

    return color


def alignment(sheetId=0, startRow=0, endRow=1, startColumn=0, endColumn=1, property="CENTER", position="horizontalAlignment",
              type="repeatCell"):
    """
    -> Function to alignment cells.

    :param sheetId: (Number) Is an identifier of a Sheet in a Google Spreadsheet. 
    :param startRow: (Number) Start Row.
    :param endRow: (Number) End Row. Need a number greater than startRow.
    :param startColumn: (Number) Start Column
    :param endColumn: (Number) End Column. Need a number greater than startColumn.
    :param property:(String) You can use property 'LEFT', 'CENTER', 'RIGHT' or 'JUSTIFY' alignment option if use position 'horizontalAlignment.
    If use position 'vertialAlignment' you can use 'BOTTOM', 'CENTER' or 'TOP' alignment option.
    :param position:(String) You can use the 'horizontalAlignment' to horizontal alignment of table cells  or 'vertialAlignment' to vertical alignment.
    :param type:(String) You can use 'updateCells' for one cell or 'repeatCell' for multiple cells.
    
    More information in https://developers.google.com/apps-script/reference/document/horizontal-alignment
    """

    result_alignment = [
        {
            type:
            {
                "cell":
                {
                    "userEnteredFormat":
                    {
                        position: property
                    }
                },
                "range":
                {
                    "sheetId": sheetId,
                    "startRowIndex": startRow,
                    "endRowIndex": endRow,
                    "startColumnIndex": startColumn,
                    "endColumnIndex": endColumn
                },
                "fields": "userEnteredFormat"
            }
        }
    ]

    return result_alignment


def merge_cell(sheetId=0, startRow=0, endRow=1, startColumn=0, endColumn=1, mergetype="MERGE_ROWS"):
    """
    -> Function to merge cells.

    :param sheetId: (Number) Is an identifier of a Sheet in a Google Spreadsheet. 
    :param startRow: (Number) Start Row.
    :param endRow: (Number) End Row. Need a number greater than startRow.
    :param startColumn: (Number) Start Column
    :param endColumn: (Number) End Column. Need a number greater than startColumn.
    :param mergetype:(String) You can use property 'LEFT', 'CENTER', 'RIGHT' or 'JUSTIFY' alignment option if use position 'horizontalAlignment.
    """

    merge = {
        "mergeCells":
        {
            'mergeType': mergetype,
            'range':
            {
                'endColumnIndex': endColumn,
                'endRowIndex': endRow,
                'sheetId': sheetId,
                'startColumnIndex': startColumn,
                'startRowIndex': startRow
            }
        }
    }

    return merge


def create_page(sheetId, title):
    """
    -> Function to create page.

    :param sheetId: (Number) Is an identifier of a Sheet in a Google Spreadsheet. 
    :param title: (String) Is name of the page. 
    """

    page = [
        {
            "addSheet": {
                "properties": {
                    "title": title,
                    "sheetId": sheetId
                }
            }
        }
    ]

    return page


def borders(sheetId, red=0, green=0, blue=0, alpha=1, startRow=0, endRow=1, startColumn=0, endColumn=1, style="SOLID", width=1):
    """
    -> Function to color cell.
    ATTENTION: The function used color system RGBA, use the website #https://rgbacolorpicker.com/ for get your color.

    :param sheetId: (Number) Is an identifier of a Sheet in a Google Spreadsheet. 
    :param red: (Number) Channel red is value from 0 to 255.
    :param green: (Number) Channel green is value from 0 to 255.
    :param blue: (Number) Channel blue is value from 0 to 255.
    :param alpha: (Number) Channel alpha is value from 0 to 1. Represent level of transparencey/opacity.
    :param startRow: (Number) Start Row.
    :param endRow: (Number) End Row. Need a number greater than startRow.
    :param startColumn: (Number) Start Column
    :param endColumn: (Number) End Column. Need a number greater than startColumn.
    :param style:(String) You can use 'DOTTED', 'DASHED', 'SOLID', SOLID_MEDIUM', SOLID_THICK' or 'DOUBLE' line borders.
    :param width:(Number) Width of the borders.
    """

    cell_borders = [
        {
            "updateBorders": {
                "range": {
                    "sheetId": sheetId,
                    "startRowIndex": startRow,
                    "endRowIndex": endRow,
                    "startColumnIndex": startColumn,
                    "endColumnIndex": endColumn
                },
                "top": {
                    "style": style,
                    "width": width,
                    "color": {
                        "red": red / 255,
                        "green": green / 255,
                        "blue": blue / 255,
                        "alpha": alpha
                    },
                },
                "bottom": {
                    "style": style,
                    "width": width,
                    "color": {
                        "red": red / 255,
                        "green": green / 255,
                        "blue": blue / 255,
                        "alpha": alpha
                    },
                },
                "right": {
                    "style": style,
                    "width": width,
                    "color": {
                        "red": red / 255,
                        "green": green / 255,
                        "blue": blue / 255,
                        "alpha": alpha
                    },
                },
                "left": {
                    "style": style,
                    "width": width,
                    "color": {
                        "red": red / 255,
                        "green": green / 255,
                        "blue": blue / 255,
                        "alpha": alpha
                    },
                },
                "innerHorizontal": {
                    "style": style,
                    "width": width,
                    "color": {
                        "red": red / 255,
                        "green": green / 255,
                        "blue": blue / 255,
                        "alpha": alpha
                    },
                },
                "innerVertical": {
                    "style": style,
                    "width": width,
                    "color": {
                        "red": red / 255,
                        "green": green / 255,
                        "blue": blue / 255,
                        "alpha": alpha
                    },
                }
            }
        }
    ]

    return cell_borders
