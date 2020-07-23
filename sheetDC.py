from __future__ import print_function
import globals as Var
import getData


def mainDC():
    result = getData.sheet.values().get(spreadsheetId=getData.SPREADSHEET_ID,
                                range=getData.color_RANGE_NAME).execute()
    values = result.get('values', [])

    result = getData.sheet.values().get(spreadsheetId=getData.SPREADSHEET_ID,
                                range=getData.preorder_RANGE_NAME,
                                majorDimension='COLUMNS').execute()
    Var.preOrder = result.get('values', [])

    result = getData.sheet.values().get(spreadsheetId=getData.SPREADSHEET_ID,
                                range=getData.priceFR_RANGE_NAME,
                                majorDimension='COLUMNS').execute()
    Var.priceFR = result.get('values', [])

    Var.copy_priceFR = Var.priceFR[:]

    if not values:
        print('No data found.')
    else:
        for i, row in enumerate(values):
            # 先記顏色代碼
            if i == 0:
                # 從第三欄開始計顏色碼
                Var.allColor = row
                del(Var.allColor[0:2]) # 只保留['52-', '53-', '57-', '54-', '26L-', '77-', 'E6-', 'E7-']

            # 開始判斷要什麼顏色
            else:
                for c in range(2, len(row)):
                    # 找出有打勾勾的顏色
                    if row[c] != "":
                        Var.needColor.append(row[1] + Var.allColor[c-2])

                    # 找出有要預購的人SSA
                    if '月' in row[c]:
                        Var.preOrderDevice.append(row[1] + Var.allColor[c-2])
                        Var.preOrderDate.append(row[c])

        Var.preOrderMsg = {}
        Var.preOrderMsg = dict(zip(Var.preOrderDevice, Var.preOrderDate))
        print(Var.preOrderMsg)