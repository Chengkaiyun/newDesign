from __future__ import print_function
import tkinter as tk
from tkinter import filedialog
from pandas.core.frame import DataFrame
import re

import globals as Var
import newDesign as index

def init():
    Var.allBrand = Var.allBrandB2C
    Var.deviceWhich = Var.deviceB2C
    Var.ChooseSheet = "B2C"

# 選擇要存放的路徑 uploadRoot
def selectDir():
    global uploadRoot
    index.okLabel("UFno")
    uploadRoot = tk.filedialog.askdirectory()  # 資料夾本人路徑 ex.C:/Users/Evolutivelabs/Roaming/Downloads/2000
    index.root("UF",uploadRoot)
    index.okBtn("UF",uploadRoot)

def makeFile():
    global fileName
    getCountry(Var.UFcountry) # 判斷是要哪個SE

    Var.brandNum = len(Var.allBrand[0])
    for i in range(Var.brandNum):
        fileName = Var.allBrand[0][i] + r"_" + Var.designer[1][1] + r".xlsx"

        brand = index.checkedBrand(i)
        print(brand)

        if brand != "":  # 如果是有打勾的手機品牌
            getdevice(brand)
            getTogether(brand)
            writeFile()

def getdevice(brand):
    global thisBrand
    thisBrand = []  # 存放那brand裡的所有型號
    Var.brandNum = len(Var.deviceWhich)

    for i in range(Var.brandNum):
        if (Var.deviceWhich[i][0] == brand):
            thisBrand = Var.deviceWhich[i]
            break


def getTogether(brand):
    global allData
    allData = []
    B2CPattern = "iPhone [56]"
    KrPattern = "iPhone [678S]|X$"
    row = 0

    for image in Var.designer[0]:
        for i, phone in enumerate(thisBrand):
            if image == "sku":
                allData.append(Var.header[0])
                break
            else:
                # FOR 去除title(mod/iphone/samsung...)
                if i == 0:
                    continue

                # FOR B2C MOD
                elif Var.ChooseSheet == "B2C" and brand == "mod" and re.search(B2CPattern, phone) != None:
                    list = [image, Var.designer[1][1], phone, Var.designer[3][row], "12345", "", "bundle", "mod"]

                # FOR B2B MOD
                elif Var.ChooseSheet == "B2B" and brand == "mod (bundle)" and re.search(B2CPattern, phone) != None:
                    list = [image, Var.designer[1][1], phone, Var.designer[3][row], "12345", "", "bundle", "mod"]
                elif Var.ChooseSheet == "B2B" and brand == "mod (single)" and re.search(B2CPattern, phone) != None:
                    list = [image, Var.designer[1][1], phone, Var.designer[3][row], "12345", "", "single", "mod"]
                elif Var.ChooseSheet == "B2B" and brand == "mod (single)":
                    list = [image, Var.designer[1][1], phone, Var.designer[3][row], "12345", "", "single", "nx"]

                # FOR Kroma MOD
                elif Var.ChooseSheet == "Kroma" and brand == "mod" and (re.search(KrPattern, phone) != None):
                    list = [image, Var.designer[1][1], phone, Var.designer[3][row], "12345", "", "bundle", "mod"]

                # FOR all not MOD
                else:
                    list = [image, Var.designer[1][1], phone, Var.designer[3][row], "12345", "", "bundle", "nx"]
                allData.append(list)
        row = row + 1

def writeFile():
    data = DataFrame(allData)
    data.to_excel(uploadRoot + r"/" + fileName, encoding='utf8', index=False, header=False)
    index.okLabel("UFok")
    print("ok")

def getCountry(country):
    if country == "tw":
        if (r"iPhone SE (第2世代)" in Var.deviceWhich[0]):
            Var.deviceWhich[0].remove(r"iPhone SE (第2世代)")
            Var.deviceWhich[0].remove(r"iPhone SE (2nd generation)")
            Var.deviceWhich[1].remove(r"iPhone SE (第2世代)")
            Var.deviceWhich[1].remove(r"iPhone SE (2nd generation)")

    elif country == "jp":
        if (r"iPhone SE (第2代)" in Var.deviceWhich[0]):
            Var.deviceWhich[0].remove(r"iPhone SE (第2代)")
            Var.deviceWhich[0].remove(r"iPhone SE (2nd generation)")
            Var.deviceWhich[1].remove(r"iPhone SE (第2代)")
            Var.deviceWhich[1].remove(r"iPhone SE (2nd generation)")

    else:
        if (r"iPhone SE (第2世代)" in Var.deviceWhich[0]):
            Var.deviceWhich[0].remove(r"iPhone SE (第2世代)")
            Var.deviceWhich[0].remove(r"iPhone SE (第2代)")
            Var.deviceWhich[1].remove(r"iPhone SE (第2世代)")
            Var.deviceWhich[1].remove(r"iPhone SE (第2代)")










