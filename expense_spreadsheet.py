from tracemalloc import start
import useful_sheets
import useful_drive

from googleapiclient.discovery import build
from googleapiclient.errors import HttpError


# The ID and range of a sample spreadsheet.
SAMPLE_SPREADSHEET_ID = '1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms'
#SAMPLE_RANGE_NAME = 'Página1!A1:E'  If you want static range and page


def main():
    #Get token
    creds = useful_sheets.initialize_sheets()
    creds_drive = useful_drive.initialize_drive()

    try:

        service_drive = build('drive', 'v3', credentials=creds_drive)
        service = build('sheets', 'v4', credentials=creds)
        sheet = service.spreadsheets()

        month = ['January', 'February', 'March', 'April', 'May', 'June',
                 'July', 'August', 'September', 'October', 'November', 'December']
        cred_card_table_title = [
            ['Bill', 'Current Installment', 'End Installment', 'Cost']
        ]
        cred_card_table = [
            ["Gifts", 1, 3, '$ 50,00'],
            ["Clothing", 1, 5, '$ 23,99'],
        ]

        #Create pages for each month
        for index, month in enumerate(month):
            title = f'{month}'
            value_index = index + 1
            new_page = useful_sheets.create_page(value_index, month)
            sheet.batchUpdate(spreadsheetId=SAMPLE_SPREADSHEET_ID, body={
                              "requests": new_page}).execute()

            update_cell = [
                [f'{month} budget - 2022'],
                ['Bill', 'Cost', 'Date', 'Pay'],
                ['Water', '', '1','Ok'],
                ['Eletricity', '$ 55,00', '5'],
                ['Phone', '$ 68,00', '15'],
                ['Supplies', '$ 65,00', '15'],
                ['Credit Card', '$ 50,00', '15'],
                ['Gas', '$ 40,00', '5'],
                ['Gym', '', '5'],
                ['Cable TV', '$ 50,00', '15'],
                [],
                ['Total', '=CONCAT(CONCAT("$"; SOMA(B3:B11));",00")']
            ]
            
            range_table = f'{title}!A1:E'   #Name of page + range
            result_table = sheet.values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=range_table,
                                                 valueInputOption='USER_ENTERED', body={"values": update_cell}).execute()

            merge = useful_sheets.merge_cell(value_index, endRow=1, endColumn=4)
            resize = useful_sheets.auto_resize(value_index)
            color = useful_sheets.color_cell(
                value_index, 129, 111, 253, 1, 11, 12, 0, 1, type="updateCells")
            horizontal_alignment = useful_sheets.alignment(
                value_index, endRow = 2, endColumn = 4)
            horizontal_alignment_cred_table = useful_sheets.alignment(
                value_index, startRow = 2, endRow=3, startColumn = 6,  endColumn = 10)
            cell_borders = useful_sheets.borders(
                value_index, endRow = 12, endColumn = 4)
            cell_borders_cred_table = useful_sheets.borders(
                value_index, startRow = 2, endRow=7, startColumn = 6,  endColumn = 10)


            oneRequest = [resize, horizontal_alignment,
                          cell_borders, merge, color,horizontal_alignment_cred_table, cell_borders_cred_table]

            resultRequest = sheet.batchUpdate(spreadsheetId=SAMPLE_SPREADSHEET_ID, body={
                                              "requests": oneRequest}).execute()

            range_name_title = f'{title}!G3:J'
            result_table = sheet.values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=range_name_title,
                                           valueInputOption='USER_ENTERED', body={"values": cred_card_table_title}).execute()

            range_name = f'{title}!G4:J'
            result = sheet.values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=range_name,
                                           valueInputOption='USER_ENTERED', body={"values": cred_card_table}).execute()
                                           
            #For each cred_card_table verify if end installment was finish.
            for index, installment in enumerate(cred_card_table):
                installment[1] += 1 
                if installment[1] > installment[2]:     
                    cred_card_table.pop(index)       
                    break

        # Create permision and send file to email. 
        useful_drive.create_permission(service_drive, SAMPLE_SPREADSHEET_ID, 'abcdefcg@gmail.com')

    except HttpError as err:
        print(err)


if __name__ == '__main__':
    main()
    # help(useful_sheets.merge_cell)
    # help(useful_sheets.resize)
    # help(useful_sheets.color_cell)
    # help(useful_sheets.alignment)
    # help(useful_sheets.borders)
    # help(useful_sheets.create_page)
    # help(useful_drive.create_permission)