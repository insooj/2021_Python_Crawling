import tkinter as tk
from tkinter import *
import tkinter.font
from openpyxl.workbook import Workbook
from openpyxl import load_workbook
import requests
import pandas as pd
import numpy as np
import folium
from folium.plugins import MiniMap
import webview

class Main(tk.Tk):
    def __init__(self, *args, **kwargs):
        tk.Tk.__init__(self, *args, **kwargs)

        self.title("바른치과 길잡이")
        self.geometry("1375x850+100+100")
        self.resizable(0,0)
        
        container = tk.Frame(self)
        container.pack(side="top", fill="both", expand=True)
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)

        def close():
            self.quit()
            self.destroy()

        menubar = tk.Menu(self)

        menu_1=tk.Menu(menubar, tearoff=0)
        menu_1.add_command(label="제작자")
        menu_1.add_command(label="버전")
        menu_1.add_separator()
        menu_1.add_command(label="종료하기", command=close)
        menubar.add_cascade(label="메뉴", menu=menu_1)

        self.config(menu=menubar)
        
        self.frames = {}
        for F in (StartPage, PageOne, PageTwo):
            page_name = F.__name__
            frame = F(parent=container, controller=self)
            self.frames[page_name] = frame
            frame.grid(row=0, column=0, sticky="nsew")

        self.show_frame("StartPage")

    def show_frame(self, page_name):
        frame = self.frames[page_name]
        frame.tkraise()

class StartPage(tk.Frame):
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        self.controller = controller

        buttonStart = tk.Button(self, overrelief="solid", text="시작하기", command=lambda: controller.show_frame("PageOne"))

        font = tk.font.Font(family="맑은고딕", size=50)

        labelTitle = tk.Label(self,text="바른치과 길잡이", font=font)
        labelMake = tk.Label(self, text="Version_0.0\nmade by 서준일, 신준호, 장인수", justify = "right")

        labelTitle.place(relx = 0.45, rely = 0.3)
        labelMake.pack(side="right", anchor="s")
        buttonStart.place(relx = 0.45, rely = 0.7, width = 100, height = 50)


class PageOne(tk.Frame):
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        self.controller = controller

        SelcetDentCount = 0

        def getAdress():
            global entrytext
            entrytext = entry.get()
            self.address = entry.get()
            showListbox()
            

        def showListbox():        
            SelcetDentCount = 0
            cnt = 0
            listbox.delete(0, END)
            for item in col_address:
                if(entrytext in item.value):
                    SelectDentArray[0][SelcetDentCount] = DentArray[0][cnt]
                    SelectDentArray[1][SelcetDentCount] = DentArray[3][cnt]

                    strSelectDent = '치과이름: ' + SelectDentArray[0][SelcetDentCount] +'  주소: ' + SelectDentArray[1][SelcetDentCount]

                    SelcetDentCount = SelcetDentCount + 1

                    listbox.insert(END, strSelectDent)
                cnt = cnt+1
            

        def showMap():
            tk.PanedWindow = webview.create_window('카카오맵','http://127.0.0.1:1030')
            webview.start()


        savingExcel = load_workbook('C:/Users/insu4/Desktop/crawling/selenium/치과 데이터.xlsx')
        SheetExcel = savingExcel.active

        col_name = SheetExcel["B"]      #치과명
        col_name2 = SheetExcel["C"]     #원장 이름
        col_number = SheetExcel["D"]    #전화번호
        col_address = SheetExcel["E"]   #주소

        DentArray=[['a' for _ in range(1666)]for _ in range(int(4))] #발표 전에 확인하고 수정
        i = 0
        for item in col_name:
            DentArray[0][i] =item.value
            i = i+1
        i=0

        for item in col_name2:
            DentArray[1][i] = item.value
            i = i+1
        i=0

        for item in col_number:
            DentArray[2][i] = item.value
            i = i+1
        i=0

        for item in col_address:
            DentArray[3][i] = item.value
            i = i+1
        i=0

        SelectDentArray=[['a' for _ in range(500)]for _ in range(int(2))] #발표 전에 확인하고 수정
        
        
        entry = tk.Entry(self)
        entry.insert(0, "주소를 입력하세요")
        buttonEnterAdress = tk.Button(self, overrelief="solid", text="입력", command=getAdress)
        buttonlistbox = tk.Button(self, overrelief="solid", text="선택", command=showMap)
        label = tk.Label(self, text = "주소")
        
        

        listbox = tk.Listbox(self, selectmode="extended")


        #def getSelect(self):
        #    selection = listbox.curselection()
        #    if(len(selection) == 0):
        #        return

        #    Selectadress = listbox.get(selection[0])
        #    return Selectadress

        #getSelect(self)

        listbox.yview()
        entry.place(relx=0.05, rely=0.1, width = 200)
        buttonEnterAdress.place(relx = 0.25, rely = 0.15)
        listbox.place(relx = 0.35, rely = 0.1, width = 500, height = 500)
        buttonlistbox.place(relx = 0.2, rely = 0.6)
        label.place(relx=0.05, rely=0.3)

class PageTwo(PageOne):
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        self.controller = controller  

        font = tk.font.Font(family="맑은고딕", size=10)
        labelTitle = tk.Label(self, font=font, text=" ")

        labelTitle.place(relx=0.05, rely=0.1, width = 200)
        label = tk.Label(self)
        label.place(relx=0.05, rely=0.3)
        


if __name__ == "__main__":
    app = Main()
    app.mainloop()