from __future__ import print_function
import os.path
import tkinter as tk
from tkinter import filedialog
import csv
import pandas as pd
import re

import globals as Var
import newDesign as index

def Makedelete():
    global outputfile
    outputfile = index.getOutputName()
    style = Var.ShopifyStock
    #  判斷是傳shopify 1的檔案還是stock 2
    if style == 1:
        colSKU = 13
    elif style == 2:
        colSKU = 3
    else:
        print("輸入錯囉!!")
        #quit()

    DeletingColor(colSKU)

def DeletingColor(colSKU):
    #  先讀csv檔進來
    try:
        # 單一檔案的
        if Var.fileORdir == 1:
            print(inputfile)
            input = open(inputfile, 'r', encoding="utf-8", newline='')
        # 資料夾的
        if Var.fileORdir == 2:
            print(dirRoot + r"/withoutTrim.csv")
            input = open(dirRoot + r"/withoutTrim.csv", 'r', encoding="utf-8", newline='')

        output = open(dirRoot + "/" + outputfile + r".csv", 'w', encoding="utf-8", newline='')
        writer = csv.writer(output)

        for i, row in enumerate(csv.reader(input)):
            if row != ['\x1a'] and i == 0:
                # 為了把第一行header寫進去
                writer.writerow(row)
                # 判斷是不是給shopify/stock
                if "sku" not in row[colSKU].lower():
                    print("你是不是選錯了阿(shopify or stock)")
                    index.okLabel("Error")

                    if os.path.exists(dirRoot + "/" + outputfile + r".csv"):
                        output.close()
                        os.remove(dirRoot + "/" + outputfile + r".csv")
                    return

            # 開始把要的顏色一行行寫進去: row[colSKU] = 'SSA03173E7-OA06'
            elif row != ['\x1a']:
                pattern = '[0-9]{5}[A-Z0-9]{2,3}-'
                PreDevice = re.search(pattern, row[colSKU]).group()

                # 如果是需要的顏色
                if PreDevice in Var.needColor:
                    # 如果有預購狀態 row[8] and 是選擇上傳shopify的
                    if PreDevice in Var.preOrderDevice and Var.ShopifyStock == 1:
                        PutPreorder(PreDevice, row)
                    writer.writerow(row)

        index.okLabel("DCok")
        input.close()

    except Exception as e:
        print(e)
        index.okLabel("Error")
        if os.path.exists(dirRoot + "/" + outputfile + r".csv"):
            output.close()
            os.remove(dirRoot + "/" + outputfile + r".csv")


def PutPreorder(PreDevice, row):
    temp = Var.preOrderMsg[PreDevice].split('月')
    month = int(temp[0])  # 幾月
    around = temp[1]  # 上中下旬

    if (around == '上旬'):
        Var.preOrderGlobal = Var.preOrder[Var.DCgetCountry][1]
    elif (around == '中旬'):
        Var.preOrderGlobal = Var.preOrder[Var.DCgetCountry][2]
    elif (around == '下旬'):
        Var.preOrderGlobal = Var.preOrder[Var.DCgetCountry][3]

    # 把@換成各國的月份，並清除空白(怕月份裡結尾有空白)
    Var.preOrderGlobal = " (" + Var.preOrderGlobal.replace('@', Var.preOrder[Var.DCgetCountry][month + 3].strip()) + ")"
    row[8] = row[8] + Var.preOrderGlobal

def selectFile():
    Var.fileORdir = 1
    global inputfile, dirRoot
    inputfile = tk.filedialog.askopenfilename()

    splitRoot = os.path.split(inputfile)
    dirRoot = splitRoot[0]
    fileRoot = splitRoot[1]

    index.root("DC", fileRoot)
    index.okLabel("DCno")
    index.okBtn("DC",fileRoot)


def selectDir():
    Var.fileORdir = 2
    global dirRoot, files
    files = []
    dirRoot = tk.filedialog.askdirectory()  # 資料夾本人路徑 ex.C:/Users/Evolutivelabs/Roaming/Downloads/2000
    allfiles = os.listdir(dirRoot)  # 資料夾內所有檔案
    for file in allfiles:
        if file.endswith(".csv"):
            files.append(file)
    print(files)

    index.root("DC", dirRoot)
    index.okLabel("DCno")
    index.okBtn("DC",dirRoot)
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

    all.to_csv(dirRoot + r"/withoutTrim.csv", encoding='utf_8_sig', index=False)  # 存檔至New_Data.csv中









