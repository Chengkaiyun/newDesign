from __future__ import print_function
import newDesign as index
import getData
import globals as Var

def mainUF():
    # 設計師分頁
    result = getData.sheet.values().get(spreadsheetId=getData.SPREADSHEET_ID,
                                range=getData.designer_RANGE_NAME,
                                majorDimension='COLUMNS').execute()
    Var.designer = result.get('values', [])

    result = getData.sheet.values().get(spreadsheetId=getData.SPREADSHEET_ID,
                                range=getData.designer_RANGE_NAME).execute()
    Var.header = result.get('values', [])

    # device分頁
    result = getData.sheet.values().get(spreadsheetId=getData.SPREADSHEET_ID,
                                range='device b2c!A1:Z',
                                majorDimension='COLUMNS').execute()
    Var.deviceB2C = result.get('values', [])

    result = getData.sheet.values().get(spreadsheetId=getData.SPREADSHEET_ID,
                                range='device b2b!A1:Z',
                                majorDimension='COLUMNS').execute()
    Var.deviceB2B = result.get('values', [])

    result = getData.sheet.values().get(spreadsheetId=getData.SPREADSHEET_ID,
                                range='device kroma!A1:Z',
                                majorDimension='COLUMNS').execute()
    Var.deviceKroma = result.get('values', [])

    result = getData.sheet.values().get(spreadsheetId=getData.SPREADSHEET_ID,
                                range='device b2c!A1:Z').execute()
    Var.allBrandB2C = result.get('values', [])

    result = getData.sheet.values().get(spreadsheetId=getData.SPREADSHEET_ID,
                                range='device b2b!A1:Z').execute()
    Var.allBrandB2B = result.get('values', [])

    result = getData.sheet.values().get(spreadsheetId=getData.SPREADSHEET_ID,
                                range='device kroma!A1:Z').execute()
    Var.allBrandKroma = result.get('values', [])

    # 更新時間
    index.updateTime("time")

    # 檢查設計師名稱有沒有填到
    if not Var.designer:
        print('No design found.')
    else:
        print('抓到 google sheet 資料')

    if len(Var.designer[1]) == 1:
        print('未填設計師名稱')
        index.updateTime("Error Design")