from __future__ import print_function
import tkinter as tk
import tkinter.font as tkFont
from tkinter import ttk
import tkinter.messagebox
import time

import globals as Var
import getData
import uploadFile as UF
import deleteColor as DC


#----------------     upload file     ----------------#

def callbackFunc(event):
    global checkBrand

    Var.ChooseSheet = combo.get()
    print(Var.ChooseSheet)

    # 紀錄上個狀態的數量，刪除原本的 Checkbox
    for i in range(len(Var.allBrand[0])):
        Var.cBox[i].destroy()

    # b2b,kroma不能選國家
    if Var.ChooseSheet == "B2C":
        SEcountryText.set(0)
        Var.allBrand = Var.allBrandB2C
        Var.deviceWhich = Var.deviceB2C

    elif Var.ChooseSheet == "B2B":
        SEcountryText.set(0)
        Var.allBrand = Var.allBrandB2B
        Var.deviceWhich = Var.deviceB2B

    elif Var.ChooseSheet == "Kroma":
        SEcountryText.set(2)
        Var.allBrand = Var.allBrandKroma
        Var.deviceWhich = Var.deviceKroma

    else:
        print("Error in callbackFunc")

    radioColorUF()
    makeCheckBox() # 重新放一次選擇brand的checkbox

def makeCheckBox():
    Var.brandNum = len(Var.allBrand[0])
    Var.cBox = [""] * Var.brandNum
    Var.checkBrand = [""] * Var.brandNum

    for i, oneBrand in enumerate(Var.allBrand[0]):
        Var.checkBrand[i] = tk.StringVar()
        Var.cBox[i] = tk.Checkbutton(my_window, text=oneBrand, variable=Var.checkBrand[i], onvalue=oneBrand, offvalue="",
                                 font=myfont)
        Var.cBox[i].grid(padx=5, pady=5, row=2, column=(i + 1))
        Var.cBox[i].toggle()

def okLabel(ok):
    if ok == "UFok":
        Var.UFokLabel['text'] = "OK"
    elif ok == "UFno":
        Var.UFokLabel['text'] = ""
    elif ok == "DCok":
        Var.DCokLabel['text'] = "OK"
    elif ok == "DCno":
        Var.DCokLabel['text'] = ""
    elif ok == "Error":
        Var.DCokLabel['text'] = "選錯shopify or stock了吧"

def root(which,root):
    if which == "UF":
        Var.UFrootLabel['text'] = root
    if which == "DC":
        Var.DCrootLabel['text'] = root

def okBtn(which,root):
    if root != "":
        if which == "UF":
            Var.UFokBtn['state'] = tk.NORMAL
        if which == "DC":
            Var.DCokBtn['state'] = tk.NORMAL

def checkedBrand(i):
    return Var.checkBrand[i].get()

def updateTime(Msg):
    if Msg == "time":
        timeHM = time.strftime("%H:%M", time.localtime())
        Var.updateLabel['text'] = timeHM + "更新"
    elif Msg == "Error Design":
        Var.updateLabel['text'] = "未填設計師名稱"

def getOutputName():
    return Var.outputName.get()

def radioColorUF():
    Var.UFcountry = Var.SEcountry[SEcountryText.get()]
    okLabel("UFno")
    for i, Dr in enumerate(SEcountryRadio):
        Dr['bg'] = '#F0F0F0'
        if SEcountryText.get() == i:
            Dr['bg'] = '#fcea95'


#----------------     delete color     ----------------#
# 判斷預購要掛個國家
def radioColorCountry():
    Var.DCgetCountry = AllcountryText.get()
    okLabel("DCno")
    for i, Dr in enumerate(AllcountryRadio):
        Dr['bg'] = '#F0F0F0'
        if AllcountryText.get() == i:
            Dr['bg'] = '#fcea95'


# 判斷shopify/stock : typeRadio
def radioColorDC():
    Var.ShopifyStock = typeRadio.get()
    okLabel("DCno")
    if typeRadio.get() == 1:
        rShopify['bg'] = '#ffc9c2'
        rStock['bg'] = '#F0F0F0'
    if typeRadio.get() == 2:
        rStock['bg'] = '#ffc9c2'
        rShopify['bg'] = '#F0F0F0'

def infoMsg(changePriceMsg):
    tk.messagebox.showinfo('注意', changePriceMsg)

def checkMsg(countryName):
    Var.checked = tk.messagebox.askyesnocancel('注意', '確定國家是' + countryName +  '嗎')

#-----------------------------------     GOGOGO     ----------------------------------#
if __name__ == '__main__':

    # 使用者介面 https://itw01.com/CQU7EVM.html
    my_window = tk.Tk()
    path = tk.StringVar()
    my_window.geometry('1100x500')
    myfont = tkFont.Font(family='Lucida Console', size=9)

    #################---------- 製作 upload file ----------#################

    tk.Label(my_window, text="製作 Upload File", bg="#C4E1E1", width="150", height="2", font=myfont).\
            grid(padx=5, pady=10, row=0,column=0, columnspan=10)

    # 下拉選單
    combo = ttk.Combobox(my_window, values=["B2C", "B2B", "Kroma"], state="readonly")
    combo.current(0)
    combo.bind("<<ComboboxSelected>>", callbackFunc)
    combo.grid(padx=5, pady=10, row=1, column=5)

    # 更新(重抓google sheet)
    updateBtn = tk.Button(my_window, text="重抓google sheet", command=getData.main, font=myfont)
    updateBtn.grid(padx=5, pady=10, row=5, column=3)
    Var.updateLabel = tk.Label(my_window, width=15, text="", bg='#C4E1FF')
    Var.updateLabel.grid(padx=5, pady=10, row=5, column=4)
    Var.UFokLabel = tk.Label(my_window, width=10, text="", bg='#C4E1FF', font=myfont)
    Var.UFokLabel.grid(padx=5, pady=10, row=4, column=4)

    # 選擇要存放的資料夾
    tk.Button(my_window, text="選擇資料夾", command=UF.selectDir, font=myfont).grid(padx=5, pady=10, row=1, column=1)
    Var.UFrootLabel = tk.Label(my_window, width=40, bg="white", font=myfont)
    Var.UFrootLabel.grid(padx=5, pady=10, row=1, column=2, columnspan=5, sticky=tk.W)

    # 抓google sheet
    getData.main()
    UF.init()

    # checkbox for brand
    makeCheckBox()

    # 選擇國家(為了SE名字)
    SEcountryText = tk.IntVar()
    SEcountryText.set(0)
    SEcountryRadio = [""] * len(Var.SEcountry)
    for i in range(len(SEcountryRadio)):
        SEcountryRadio[i] = tk.Radiobutton(my_window, text=Var.SEcountry[i], variable=SEcountryText, value=i,
                                  command=radioColorUF, font=myfont)
        SEcountryRadio[i].grid(padx=5, pady=10, row=3, column=i + 1)
    radioColorUF()

    # 確定
    Var.UFokBtn = tk.Button(my_window, text="輸出", command=UF.makeFile, state=tk.DISABLED, font=myfont)
    Var.UFokBtn.grid(padx=5, pady=10, row=4, column=3)

    tk.Label(my_window, text="刪除顏色", bg="#C4E1E1", width="150", height="2", font=myfont).\
            grid(padx=5, pady=10, row=8, column=0, columnspan=10)

    #################---------- 刪除顏色 ----------#################

    Var.DCokLabel = tk.Label(my_window, width=25, text="", bg='#C4E1FF', font=myfont)
    Var.DCokLabel.grid(padx=5, pady=5, row=14, column=5)

    # 路徑以及選擇路徑的按鈕
    tk.Label(my_window, text="路徑:", font=myfont).grid(padx=5, pady=5, row=10, column=1)
    tk.Button(my_window, text="選擇檔案", command=DC.selectFile, font=myfont).grid(padx=5, pady=5, row=10, column=4)
    tk.Button(my_window, text="選擇資料夾", command=DC.selectDir, font=myfont).grid(padx=5, pady=5, row=10, column=5)
    Var.DCrootLabel = tk.Label(my_window, width=40, bg="white", font=myfont)
    Var.DCrootLabel.grid(padx=5, pady=5, row=10, column=2, columnspan=8, sticky=tk.W)

    # 選擇國家(為了預購)
    AllcountryText = tk.IntVar()
    AllcountryText.set(0)
    AllcountryRadio = [""] * len(Var.Allcountry)
    for i in range(len(AllcountryRadio)):
        AllcountryRadio[i] = tk.Radiobutton(my_window, text=Var.Allcountry[i], variable=AllcountryText, value=i,
                                  command=radioColorCountry, font=myfont)
        AllcountryRadio[i].grid(padx=5, pady=10, row=11, column=i + 1)
    radioColorCountry()

    # 選擇上傳stock or shopify
    typeRadio = tk.IntVar()
    typeRadio.set(1)
    rShopify = tk.Radiobutton(my_window, text='上傳Shopify', variable=typeRadio, value=1, command=radioColorDC, font=myfont)
    rShopify.grid(padx=5, pady=10, row=12, column=2)
    rStock = tk.Radiobutton(my_window, text='上傳Stock', variable=typeRadio, value=2, command=radioColorDC, font=myfont)
    rStock.grid(padx=5, pady=10, row=12, column=3)
    radioColorDC()


    # 新文件檔名以及確定執行按鈕
    tk.Label(my_window, text="新文件檔名:", font=myfont).grid(padx=5, pady=14, row=14, column=1)
    Var.outputName = tk.StringVar()
    Var.outputName.set("all_ok")
    tk.Entry(my_window, textvariable=Var.outputName).grid(padx=5, pady=5, row=14, column=2)
    Var.DCokBtn = tk.Button(my_window, text="確定", command=DC.Makedelete, state=tk.DISABLED, font=myfont)
    Var.DCokBtn.grid(padx=5, pady=5, row=14, column=4)

    my_window.mainloop()

