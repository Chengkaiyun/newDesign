from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

import tkinter as tk
from tkinter import filedialog
import tkinter.font as tkFont
from tkinter import ttk
import csv
from pandas.core.frame import DataFrame
import pandas as pd
import time
import re

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

# The ID and range of a sample spreadsheet.
SPREADSHEET_ID = '1BD1A01VL2Bx_bMBgoc3eCwb9cT4BWI2fzuQdQop8yUo'
deviceWhich_RANGE_NAME = 'device b2c!A1:Z'
designer_RANGE_NAME = 'designer!A1:Z'
color_RANGE_NAME = 'SSA!A2:Z'

needColor = []
allColor = []
allBrand = []
ChooseSheet = ""
deviceB2C = []
deviceB2B = []
deviceKroma = []
deviceWhich = []
allBrandB2C =[]
allBrandB2B = []
allBrandKroma =[]

def main():
    """Shows basic usage of the Sheets API.
    Prints values from a sample spreadsheet.
    """
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'client_secret.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('sheets', 'v4', credentials=creds)

    # Call the Sheets API
    sheet = service.spreadsheets()

    # =============================== upload file用 =============================== #
    global designer, deviceWhich, header, allBrand
    global deviceWhich_RANGE_NAME
    global deviceB2C, deviceB2B, deviceKroma, allBrandB2C, allBrandB2B, allBrandKroma
    updatelabel['text'] = ""

    # 看傳進來的下拉選單是什麼決定抓哪分頁

    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID,
                                range=designer_RANGE_NAME,
                                majorDimension='COLUMNS').execute()
    designer = result.get('values', [])

    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID,
                                range=designer_RANGE_NAME).execute()
    header = result.get('values', [])

    # 看device分頁
    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID,
                                range='device b2c!A1:Z',
                                majorDimension='COLUMNS').execute()
    deviceB2C = result.get('values', [])

    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID,
                                range='device b2b!A1:Z',
                                majorDimension='COLUMNS').execute()
    deviceB2B = result.get('values', [])

    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID,
                                range='device kroma!A1:Z',
                                majorDimension='COLUMNS').execute()
    deviceKroma = result.get('values', [])

    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID,
                                range='device b2c!A1:Z').execute()
    allBrandB2C = result.get('values', [])

    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID,
                                range='device b2b!A1:Z').execute()
    allBrandB2B = result.get('values', [])

    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID,
                                range='device kroma!A1:Z').execute()
    allBrandKroma = result.get('values', [])

    timeHM = time.strftime("%H:%M", time.localtime())
    updatelabel['text'] = timeHM + "更新"
    oklabel['text'] = ""

    if not designer:
        print('No design found.')
    else:
        print(timeHM+'抓到 google sheet 資料')

    if len(designer[1]) == 1:
        print('沒有填到設計師名稱')
        uploadrootlabel['text'] = "沒有填到設計師名稱"


    # =============================== delete color 用 =============================== #
    global needColor
    global allColor
    global preOrder
    global preOrderDevice, preOrderDate
    global preOrderMsg
    needColor = []
    allColor = []
    preOrder = []
    preOrderDevice = []
    preOrderDate = []


    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID,
                                range=color_RANGE_NAME).execute()
    values = result.get('values', [])

    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID,
                                range='preorder!A1:Z',
                                majorDimension='COLUMNS').execute()
    preOrder = result.get('values', [])

    if not values:
        print('No data found.')
    else:
        for i, row in enumerate(values):
            # 先記顏色代碼
            if i == 0:
                # 從第三欄開始計顏色碼
                allColor = row
                del(allColor[0:2]) # 只保留['52-', '53-', '57-', '54-', '26L-', '77-', 'E6-', 'E7-']

            # 開始判斷要什麼顏色
            else:
                for c in range(2, len(row)):
                    # 找出有打勾勾的顏色
                    if row[c] != "":
                        needColor.append(row[1] + allColor[c-2])

                    # 找出有要預購的人SSA
                    if '月' in row[c]:
                        preOrderDevice.append(row[1] + allColor[c-2])
                        preOrderDate.append(row[c])
        preOrderMsg = {}
        preOrderMsg = dict(zip(preOrderDevice, preOrderDate))
        print(preOrderMsg)

def makeFile():
    oklabel['text'] = ""
    global fileName
    getCountry(country)

    brandNum = len(allBrand[0])
    for i in range(brandNum):
        fileName = allBrand[0][i] + r"_" + designer[1][1] + r".xlsx"
        brand = checkBrand[i].get()
        print(brand)
        if brand != "":  # 如果是有打勾的手機品牌
            getdevice(brand)
            getTogether(brand)
            writeFile()


def getdevice(brand):
    global thisBrand
    thisBrand = []  # 存放那brand裡的所有型號
    brandNum = len(deviceWhich)
    # print(brandNum)

    for i in range(brandNum):
        if (deviceWhich[i][0] == brand):
            thisBrand = deviceWhich[i]
            #thisBrand.pop(0)
            break
    #print("getdevice:" + " ".join(thisBrand))


def getTogether(brand):
    global allData, ChooseSheet
    allData = []
    row = 0

    for image in designer[0]:
        for i, phone in enumerate(thisBrand):
            # print(phone)
            if image == "sku":
                allData.append(header[0])
                break
            else:
                # FOR 去除title(mod/iphone/samsung...)
                if i == 0:
                    continue

                # FOR B2C MOD
                elif ChooseSheet == "B2C" and brand == "mod" and ("iPhone 6" in phone or "iPhone 5" in phone):
                    list = [image, designer[1][1], phone, designer[3][row], "12345", "", "bundle", "mod"]

                # FOR B2B MOD
                elif ChooseSheet == "B2B" and brand == "mod (bundle)" and ("iPhone 6" in phone or "iPhone 5" in phone):
                    list = [image, designer[1][1], phone, designer[3][row], "12345", "", "bundle", "mod"]
                elif ChooseSheet == "B2B" and brand == "mod (single)" and ("iPhone 6" in phone or "iPhone 5" in phone):
                    list = [image, designer[1][1], phone, designer[3][row], "12345", "", "single", "mod"]
                elif ChooseSheet == "B2B" and brand == "mod (single)":
                    list = [image, designer[1][1], phone, designer[3][row], "12345", "", "single", "nx"]

                # FOR Kroma MOD
                elif ChooseSheet == "Kroma" and brand == "mod" and \
                        ("iPhone 6" in phone or "iPhone 7" in phone or "iPhone 8" in phone or \
                         "iPhone SE" in phone or phone == "iPhone X"):
                    list = [image, designer[1][1], phone, designer[3][row], "12345", "", "bundle", "mod"]

                # FOR all not MOD
                else:
                    list = [image, designer[1][1], phone, designer[3][row], "12345", "", "bundle", "nx"]
                allData.append(list)
        row = row + 1


def getCountry(country):
    if country == "tw":
        if (r"iPhone SE (第2世代)" in deviceWhich[0]):
            deviceWhich[0].remove(r"iPhone SE (第2世代)")
            deviceWhich[0].remove(r"iPhone SE (2nd generation)")
            deviceWhich[1].remove(r"iPhone SE (第2世代)")
            deviceWhich[1].remove(r"iPhone SE (2nd generation)")

    elif country == "jp":
        if (r"iPhone SE (第2代)" in deviceWhich[0]):
            deviceWhich[0].remove(r"iPhone SE (第2代)")
            deviceWhich[0].remove(r"iPhone SE (2nd generation)")
            deviceWhich[1].remove(r"iPhone SE (第2代)")
            deviceWhich[1].remove(r"iPhone SE (2nd generation)")

    else:
        if (r"iPhone SE (第2世代)" in deviceWhich[0]):
            deviceWhich[0].remove(r"iPhone SE (第2世代)")
            deviceWhich[0].remove(r"iPhone SE (第2代)")
            deviceWhich[1].remove(r"iPhone SE (第2世代)")
            deviceWhich[1].remove(r"iPhone SE (第2代)")


def writeFile():
    data = DataFrame(allData)
    # print(data)
    data.to_excel(uploadRoot + r"/" + fileName, encoding='utf8', index=False, header=False)
    oklabel['text'] = "OK"
    print("ok")


def selectDirUpload():
    oklabel['text'] = ""
    global uploadRoot
    uploadRoot = tk.filedialog.askdirectory()  # 資料夾本人路徑 ex.C:/Users/Evolutivelabs/Roaming/Downloads/2000
    uploadrootlabel['text'] = uploadRoot
    if uploadRoot != "":
        okBtn['state'] = tk.NORMAL


def radioColorUpload():
    oklabel['text'] = ""
    global country
    if countryRadio.get() == 1:
        r1['bg'] = '#fcea95'
        r2['bg'] = '#F0F0F0'
        r3['bg'] = '#F0F0F0'
        country = "tw"
    if countryRadio.get() == 2:
        r2['bg'] = '#fcea95'
        r1['bg'] = '#F0F0F0'
        r3['bg'] = '#F0F0F0'
        country = "jp"
    if countryRadio.get() == 3:
        r3['bg'] = '#fcea95'
        r1['bg'] = '#F0F0F0'
        r2['bg'] = '#F0F0F0'
        country = "other"


# ===============================  這裡是刪除顏色的部分  =============================== #
def DeletingColor(colSKU):
    #  先讀csv檔進來
    try:
        # 單一檔案的
        if fileORdir == 1:
            print(inputfile)
            input = open(inputfile, 'r', encoding="utf-8", newline='')
        # 資料夾的
        if fileORdir == 2:
            print(dirRoot + r"/withoutTrim.csv")
            input = open(dirRoot + r"/withoutTrim.csv", 'r', encoding="utf-8", newline='')

        output = open(dirRoot + "/" + outputfile + r".csv", 'w', encoding="utf-8", newline='')
        writer = csv.writer(output)


        for i, row in enumerate(csv.reader(input)):
            if row != ['\x1a']:
                # 為了把第一行header寫進去
                if i == 0:
                    writer.writerow(row)
                    # 判斷是不是給shopify/stock
                    if "sku" not in row[colSKU].lower():
                        print("你是不是選錯了阿(shopify or stock)")
                        coloroklabel['text'] = "選錯shopify or stock了吧"

                        if os.path.exists(dirRoot + "/" + outputfile + r".csv"):
                            output.close()
                            os.remove(dirRoot + "/" + outputfile + r".csv")
                        return

                # 開始把要的顏色一行行寫進去: row[colSKU] = 'SSA03173E7-OA06'
                else:
                    pattern = '[0-9]{5}[A-Z0-9]{2,3}-'
                    PreDevice = re.search(pattern, row[colSKU]).group()

                    # 如果是需要的顏色
                    if PreDevice in needColor:
                        # 如果有預購狀態 row[8] and 是選擇上傳shopify的
                        if PreDevice in preOrderDevice and typeRadio.get() == 1:
                            temp = preOrderMsg[PreDevice].split('月')
                            month = int(temp[0])     # 幾月
                            around = temp[1]    # 上中下旬

                            if (around == '上旬'):
                                preOrderMsgGlobal = preOrder[countryDelRadio.get()][1]
                            elif (around == '中旬'):
                                preOrderMsgGlobal = preOrder[countryDelRadio.get()][2]
                            elif (around == '下旬'):
                                preOrderMsgGlobal = preOrder[countryDelRadio.get()][3]

                            # 把@換成各國的月份，並清除空白(怕月份裡結尾有空白)
                            preOrderMsgGlobal = " (" +preOrderMsgGlobal.replace('@',preOrder[countryDelRadio.get()][month+3].strip()) + ")"
                            row[8] = row[8] + preOrderMsgGlobal
                        writer.writerow(row)

        coloroklabel['text'] = "OK"
        input.close()

    except Exception as e:
        print(e)
        coloroklabel['text'] = "選錯shopify or stock了吧"
        if os.path.exists(dirRoot + "/" + outputfile + r".csv"):
            output.close()
            os.remove(dirRoot + "/" + outputfile + r".csv")



def selectPath():
    global fileORdir
    fileORdir = 1
    global inputfile, splitRoot, dirRoot, fileRoot
    inputfile = tk.filedialog.askopenfilename()

    splitRoot = os.path.split(inputfile)
    dirRoot = splitRoot[0]
    fileRoot = splitRoot[1]
    rootlabel['text'] = splitRoot[1]
    coloroklabel['text'] = ""
    colorokBtn['state'] = tk.NORMAL


def selectDirColor():
    global fileORdir
    fileORdir = 2
    global dirRoot, fileRoot
    global files
    files = []
    dirRoot = tk.filedialog.askdirectory()  # 資料夾本人路徑 ex.C:/Users/Evolutivelabs/Roaming/Downloads/2000
    allfiles = os.listdir(dirRoot)  # 資料夾內所有檔案
    for file in allfiles:
        if file.endswith(".csv"):
            files.append(file)
    print(files)

    rootlabel['text'] = dirRoot
    coloroklabel['text'] = ""
    colorokBtn['state'] = tk.NORMAL

    allInOne()


# 開始把所有檔案合併
def allInOne():
    for i in range(len(files)):
        if files[i] == r"withoutTrim.csv":
            continue
        elif i == 0:
            all = pd.read_csv(dirRoot + r'/' + files[i], encoding="utf-8")
        else:
            temp = pd.read_csv(dirRoot + r'/' + files[i], encoding="utf-8")
            all = pd.concat([all, temp])

    all.to_csv(dirRoot + r"/withoutTrim.csv", encoding='utf8', index=False)  # 存檔至New_Data.csv中


def Makedelete():
    global outputfile
    outputfile = outputName.get()
    style = typeRadio.get()
    #  判斷是傳shopify 1的檔案還是stock 2
    if style == 1:
        colSKU = 13
    elif style == 2:
        colSKU = 3
    else:
        print("輸入錯囉!!")
        quit()

    DeletingColor(colSKU)

# 判斷shopify/stock : typeRadio
def radioColor():
    coloroklabel['text'] = ""
    if typeRadio.get() == 1:
        rShopify['bg'] = '#ffc9c2'
        rStock['bg'] = '#F0F0F0'
    if typeRadio.get() == 2:
        rStock['bg'] = '#ffc9c2'
        rShopify['bg'] = '#F0F0F0'

# 判斷預購要掛個國家 countryDelRadio
def radioColorCountry():
    coloroklabel['text'] = ""
    for i, Dr in enumerate(DrArr):
        Dr['bg'] = '#F0F0F0'
        if countryDelRadio.get() == i:
            Dr['bg'] = '#fcea95'


def callbackFunc(event):
    global ChooseSheet
    global brandNum, deviceWhich, allBrand, oneBrand, checkBrand
    global cBox
    ChooseSheet = combo.get()
    print(ChooseSheet)

    # 紀錄上個狀態的數量，刪除原本的 Checkbox
    copyBrandNum = len(allBrand[0])
    for i in range(copyBrandNum):
        cBox[i].destroy()

    # b2b,kroma不能選國家
    if ChooseSheet == "B2C":
        countryRadio.set(1)
        radioColorUpload()
        r1['state'] = tk.NORMAL
        r2['state'] = tk.NORMAL
        r3['state'] = tk.NORMAL
        allBrand = allBrandB2C
        deviceWhich = deviceB2C
        #print(allBrand[0])

    elif ChooseSheet == "B2B":
        countryRadio.set(1)
        radioColorUpload()
        r1['state'] = tk.NORMAL
        r2['state'] = tk.DISABLED
        r3['state'] = tk.DISABLED
        allBrand = allBrandB2B
        deviceWhich = deviceB2B
        #print(allBrand[0])

    elif ChooseSheet == "Kroma":
        countryRadio.set(3)
        radioColorUpload()
        r1['state'] = tk.DISABLED
        r2['state'] = tk.DISABLED
        r3['state'] = tk.NORMAL
        allBrand = allBrandKroma
        deviceWhich = deviceKroma
        #print(allBrand[0])

    else:
        print("Error in callbackFunc")

    # 重新放一次選擇brand的checkbox
    brandNum = len(allBrand[0])
    cBox = [""] * brandNum
    checkBrand = [""] * brandNum
    for i in range(brandNum):
        checkBrand[i] = tk.StringVar()

    i = 0
    for oneBrand in allBrand[0]:
        cBox[i] = tk.Checkbutton(my_window, text=oneBrand, variable=checkBrand[i], onvalue=oneBrand, offvalue="",
                                 font=myfont)
        cBox[i].grid(padx=5, pady=5, row=2, column=(i + 1))
        cBox[i].toggle()
        i = i + 1

# ===============================   這裡GOGOGO  =============================== #
if __name__ == '__main__':

    # 使用者介面 https://itw01.com/CQU7EVM.html
    my_window = tk.Tk()
    path = tk.StringVar()
    my_window.geometry('1100x500')
    myfont = tkFont.Font(family='Lucida Console', size=9)

    #################---------- 製作 upload file ----------#################

    # 下拉選單
    combo = ttk.Combobox(my_window, values=["B2C", "B2B", "Kroma"], state="readonly")
    combo.current(0)
    combo.bind("<<ComboboxSelected>>", callbackFunc)
    combo.grid(padx=5, pady=10, row=1, column=5)

    tk.Label(my_window, text="製作 Upload File", bg="#C4E1E1", width="150", height="2", font=myfont).grid(padx=5, pady=10, row=0,
                                                                                            column=0, columnspan=10)
    # 更新(重抓google sheet)
    updateBtn = tk.Button(my_window, text="重抓google sheet", command=main, font=myfont)
    updateBtn.grid(padx=5, pady=10, row=5, column=3)
    updatelabel = tk.Label(my_window, width=10, text="", bg='#C4E1FF')
    updatelabel.grid(padx=5, pady=10, row=5, column=4)
    oklabel = tk.Label(my_window, width=10, text="", bg='#C4E1FF', font=myfont)
    oklabel.grid(padx=5, pady=10, row=4, column=4)

    # 選擇要存放的資料夾
    tk.Button(my_window, text="選擇資料夾", command=selectDirUpload, font=myfont).grid(padx=5, pady=10, row=1, column=1)
    uploadrootlabel = tk.Label(my_window, width=40, bg="white", font=myfont)
    uploadrootlabel.grid(padx=5, pady=10, row=1, column=2, columnspan=5, sticky=tk.W)

    # 抓google sheet
    main()
    allBrand = allBrandB2C
    deviceWhich = deviceB2C
    ChooseSheet = "B2C"

    # checkbox for brand
    brandNum = len(allBrand[0])
    cBox = [""] * brandNum
    checkBrand = [""] * brandNum
    for i in range(brandNum):
        checkBrand[i] = tk.StringVar()

    for i, oneBrand in enumerate(allBrand[0]):
        cBox[i] = tk.Checkbutton(my_window, text=oneBrand, variable=checkBrand[i], onvalue=oneBrand, offvalue="",
                                 font=myfont)
        cBox[i].grid(padx=5, pady=5, row=2, column=(i + 1))
        cBox[i].toggle()


    # 選擇國家(為了SE名字)
    countryRadio = tk.IntVar()
    countryRadio.set(1)
    r1 = tk.Radiobutton(my_window, text='tw', variable=countryRadio, value=1, command=radioColorUpload, font=myfont)
    r1.grid(padx=5, pady=5, row=3, column=1)
    r2 = tk.Radiobutton(my_window, text='jp', variable=countryRadio, value=2, command=radioColorUpload, font=myfont)
    r2.grid(padx=5, pady=5, row=3, column=2)
    r3 = tk.Radiobutton(my_window, text='other', variable=countryRadio, value=3, command=radioColorUpload, font=myfont)
    r3.grid(padx=5, pady=5, row=3, column=3)
    radioColorUpload()

    # 確定
    okBtn = tk.Button(my_window, text="輸出", command=makeFile, state=tk.DISABLED, font=myfont)
    okBtn.grid(padx=5, pady=10, row=4, column=3)

    tk.Label(my_window, text="刪除顏色", bg="#C4E1E1", width="150", height="2", font=myfont).grid(padx=5, pady=10, row=8, column=0,
                                                                                  columnspan=10)
    #################---------- 刪除顏色 ----------#################

    coloroklabel = tk.Label(my_window, width=25, text="", bg='#C4E1FF', font=myfont)
    coloroklabel.grid(padx=5, pady=5, row=14, column=5)

    # 路徑以及選擇路徑的按鈕
    tk.Label(my_window, text="路徑:", font=myfont).grid(padx=5, pady=5, row=10, column=1)
    tk.Button(my_window, text="選擇檔案", command=selectPath, font=myfont).grid(padx=5, pady=5, row=10, column=4)
    tk.Button(my_window, text="選擇資料夾", command=selectDirColor, font=myfont).grid(padx=5, pady=5, row=10, column=5)
    rootlabel = tk.Label(my_window, width=40, bg="white", font=myfont)
    rootlabel.grid(padx=5, pady=5, row=10, column=2, columnspan=8, sticky=tk.W)

    # 選擇國家(為了預購)
    countryDelRadio = tk.IntVar()
    countryDelRadio.set(0)
    DrArrCountry = ['tw/b2b','io/eu','fr','de','jp','es']
    DrArr = [""] * len(DrArrCountry)
    for i in range(len(DrArr)):
        DrArr[i] = tk.Radiobutton(my_window, text=DrArrCountry[i], variable=countryDelRadio, value=i,
                                  command=radioColorCountry, font=myfont)
        DrArr[i].grid(padx=5, pady=10, row=11, column=i+1)
    radioColorCountry()

    # 選擇上傳stock or shopify
    typeRadio = tk.IntVar()
    typeRadio.set(1)
    outputName = tk.StringVar()
    outputName.set("all_ok")
    rShopify = tk.Radiobutton(my_window, text='上傳Shopify', variable=typeRadio, value=1, command=radioColor, font=myfont)
    rShopify.grid(padx=5, pady=10, row=12, column=2)
    rStock = tk.Radiobutton(my_window, text='上傳stock', variable=typeRadio, value=2, command=radioColor, font=myfont)
    rStock.grid(padx=5, pady=10, row=12, column=3)
    radioColor()

    # 新文件檔名以及確定執行按鈕
    tk.Label(my_window, text="新文件檔名:", font=myfont).grid(padx=5, pady=14, row=14, column=1)
    tk.Entry(my_window, textvariable=outputName).grid(padx=5, pady=5, row=14, column=2)
    colorokBtn = tk.Button(my_window, text="確定", command=Makedelete, state=tk.DISABLED, font=myfont)
    colorokBtn.grid(padx=5, pady=5, row=14, column=4)

    my_window.mainloop()
