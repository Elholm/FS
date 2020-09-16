#imports
import openpyxl
import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import filedialog as fd
import os
import time

class fs_ejd:
    '''Pre-requisites'''
    blå = ["Blå medarbejder", "Servicemedarbejder"]
    cts = ["1-4 billeder", "5-10 billeder", "11-20 billeder", "21+ billeder"]
    grå = ["Grå medarbejder"]
    grøn = ["Gartner","Gartner m. lille rider"]
    sc = ["DEAS Service Center"]
    uger_pr_år = 50
    # accepted_xls = [".xls", ".xlsm", ".xlsx", ".csv"]
    pr_uge = lambda a : a / 50
    #Functions
    def get_wb(self,wb_path):
        self.wb = openpyxl.load_workbook(wb_path,data_only = True)
        return
    def __init__(self,wb_path):
        try:
            self.wb = openpyxl.load_workbook(wb_path,data_only = True)
            self.aftale = self.wb["FS Aftale"]
            self.støtteark = self.wb["Støtteark"]
        except:
            print("Loading failed. Try another SLA with .get_wb()")
            return
        # self.nr = f"{self.aftale['F1'].value}-{self.aftale['G1'].value}"
        # print(wb_path)
        # self.nr = str(self.wb.path.split("/")[-1].split(" ")[0])
        # filename, file_ext = os.path.splitext(self.wb_path)
        self.nr = os.path.basename(wb_path).split(" ")[0]
        if self.aftale["D3"].value == "Niveau":
            self.aftale.delete_cols(4)
        else:
            pass
        return
    def get_feje_timer(self):
        if self.aftale["D52"].value == False:
            return -1
        return self.aftale["F52"].value*self.aftale["G52"].value / self.uger_pr_år
    def get_SC_money(self):
        if self.aftale["J20"].value == None:
            return -1
        return self.aftale["J20"].value
    def get_CTS_money(self):
        cts_money = 0
        for x in self.aftale["D"]:
            if x.value == True:
                if str(self.aftale[f"H{x.row}"].value) in self.cts:
                    cts_money += self.aftale[f"I{x.row}"].value
        return cts_money
    def get_grøn_timer(self):
        vv_timer = 0
        aftale = self.aftale
        for x in aftale["D"]:
            if x.value == True:
                if aftale[f"H{x.row}"].value in self.grøn:
                    vv_timer += aftale[f"F{x.row}"].value*aftale[f"G{x.row}"].value
        return vv_timer / self.uger_pr_år
    def get_blå_timer(self):
        #aftale must be a SLA excel workbook
        vv_timer = 0
        aftale = self.aftale
        for x in aftale["D"]:
          if x.value == True:
            if aftale[f"H{x.row}"].value in self.blå:
              if aftale[f"F{x.row}"].value is not None and aftale[f"G{x.row}"].value is not None:
                vv_timer += aftale[f"F{x.row}"].value*aftale[f"G{x.row}"].value
        return vv_timer / self.uger_pr_år
    def get_grå_timer(self):
        vv_timer = 0
        aftale = self.aftale
        for x in aftale["D"]:
            if x.value == True and x.row != 52:
                if aftale[f"H{x.row}"].value in self.grå:
                    vv_timer += aftale[f"F{x.row}"].value*aftale[f"G{x.row}"].value
        return vv_timer / self.uger_pr_år
    def get_worksheets(self):
        return
    def convert_ejd(self):
        old_nr = self.nr
        left, right = old_nr.split("-")
        while len(left) < 3:
            left_list = list(left)
            left_list = ["0"] + left_list
            left = ''.join(left_list)
        while len(right) < 3:
            right_list = list(right)
            right_list = ["0"] + right_list
            right = ''.join(right_list)
        self.nr = f"{left}-{right}"
        print(f"Conversion succesful: {self.nr}")
        return    
    def get_version(self):
        aftale = self.aftale
        B1 = aftale["B1"].value
        version = B1.split(" ")[1]
        return version        
    def to_sql(self): #not done
        print("Function is not yet finished.")
        return
    def to_pdf(self): #not done
        print("Function is not yet finished.")
        return
    def create_csv(self): #not done
        print("Function is not yet finished.")
        return

class App:
    def __init__(self):
    #Variables
        self.root = tk.Tk()
        self.canvas1 = tk.Canvas(self.root,width = 350,height = 400)
        self.canvas1.pack()
        self.label = tk.StringVar()
        self.label.set("No file is selected")
        tk.Label(self.root,textvariable = self.label).pack()
        self.root.title("FS SLA Værktøj")
        self.accepted_xls = [".xls", ".xlsm", ".xlsx", ".csv"]
        self.texttype = "helvetica"
        self.textsize = 12
    #Windows with text
        #Blue window
        self.text_blue = tk.StringVar()
        self.text_blue.set("Not calculated")
        self.label_blue = tk.Label(self.root,textvariable=self.text_blue, fg="blue", font=(self.texttype,12,"bold"))
        self.canvas1.create_window(200,70,window=self.label_blue)
        #Grey window
        self.text_grey = tk.StringVar()
        self.text_grey.set("Not calculated")
        self.label_grey = tk.Label(self.root,textvariable=self.text_grey, fg="grey", font=(self.texttype,12,"bold"))
        self.canvas1.create_window(200,120,window=self.label_grey)
        #Green window
        self.text_green = tk.StringVar()
        self.text_green.set("Not calculated")
        self.label_green = tk.Label(self.root,textvariable=self.text_green, fg="green", font=(self.texttype,12,"bold"))
        self.canvas1.create_window(200,170,window=self.label_green)
        #CTS window
        self.text_cts = tk.StringVar()
        self.text_cts.set("Not calculated")
        self.label_cts = tk.Label(self.root,textvariable=self.text_cts, fg="black", font=(self.texttype,12,"bold"))
        self.canvas1.create_window(200,220,window=self.label_cts)
        #SC window
        self.text_sc = tk.StringVar()
        self.text_sc.set("Not calculated")
        self.label_sc = tk.Label(self.root,textvariable=self.text_sc, fg="black", font=(self.texttype,12,"bold"))
        self.canvas1.create_window(200,270,window=self.label_sc)
        #Feje window
        self.text_feje = tk.StringVar()
        self.text_feje.set("Not calculated")
        self.label_feje = tk.Label(self.root,textvariable=self.text_feje, fg="black", font=(self.texttype,12,"bold"))
        self.canvas1.create_window(200,320,window=self.label_feje)
    #buttons init
        #Load file button
        self.button_file = tk.Button(self.root, text="Load file",command = self.callback)
        self.canvas1.create_window(50,20,window=self.button_file)
        #Get blue button
        self.button_blue = tk.Button(text="Vis Blå timer/uge",command=self.get_blue,bg="brown",fg="white")
        self.canvas1.create_window(50,70,window = self.button_blue)
        #Get grey button
        self.button_grey = tk.Button(text="Vis Grå timer/uge",command=self.get_grey,bg="brown",fg="white")
        self.canvas1.create_window(50,120,window=self.button_grey)
        #get green button
        self.button_green = tk.Button(text="Vis Grøn timer/uge",command=self.get_green,bg="brown",fg="white")
        self.canvas1.create_window(50,170,window=self.button_green)
        #get cts button
        self.button_cts = tk.Button(text="Vis CTS penge",command=self.get_cts,bg="brown",fg="white")
        self.canvas1.create_window(50,220,window=self.button_cts)
        #get sc button
        self.button_sc = tk.Button(text="Vis SC penge",command=self.get_sc,bg="brown",fg="white")
        self.canvas1.create_window(50,270,window=self.button_sc)
        #get feje button
        self.button_feje = tk.Button(text="Vis Feje timer/uge",command=self.get_feje,bg="brown",fg="white")
        self.canvas1.create_window(50,320,window=self.button_feje)
        #get_all button
        self.button_all = tk.Button(text="Vis alle nøgletal", command=self.get_all,bg="brown",fg="white")
        self.canvas1.create_window(150,20,window=self.button_all)
        #Show kunde spec (sat i standby)
        # self.button_spec = tk.Button(text="Vis kundespecifikke\nydelser", command=self.show_kunde_spec)
        # self.canvas1.create_window(50, 370,window = self.button_spec)
    #mainloop
        self.root.mainloop()
        pass
   
    #Functions
    def callback(self):
        self.wb_path = fd.askopenfilename()
        filename, file_ext = os.path.splitext(self.wb_path)
        if file_ext.lower() in self.accepted_xls:
          self.wb = fs_ejd(self.wb_path)
        else:
          self.label.set("File chosen is not Excel file")
          return
        # self.wb = self.wb["FS Aftale"]
        self.label.set(f"File chosen is:\n{filename.split('/')[-1]+file_ext} \nSLA Version: {self.wb.get_version()} \nAntal Kundespecifikke ydelser: {self.find_kunde_spec()}")
        # INSERT CLEAR ALL FUNC
        self.clear_all()

        return
    def get_blue(self):
        # label1 = tk.Label(self.root,text=f"Timer/uge: {self.wb.get_blå_timer():.2f}", fg="blue", font=(self.texttype,12,"bold"))
        # self.canvas1.create_window(200,70,window=label1)
        try:
            self.text_blue.set(f"Timer/uge: {self.wb.get_blå_timer():.2f}")
        except:
            self.text_blue.set("FAILED")
    def get_grey(self):
        # label1 = tk.Label(self.root,text=f"Timer/uge: {self.wb.get_grå_timer():.2f}", fg="grey", font=(self.texttype,12,"bold"))
        # self.canvas1.create_window(200,120,window=label1)
        try:
            self.text_grey.set(f"Timer/uge: {self.wb.get_grå_timer():.2f}")
        except:
            self.text_grey.set("FAILED")
    def get_green(self):
        # label1 = tk.Label(self.root,text=f"Timer/uge: {self.wb.get_grøn_timer():.2f}", fg="green", font=(self.texttype,12,"bold"))
        # self.canvas1.create_window(200,170,window=label1)
        try:
            self.text_green.set(f"Timer/uge: {self.wb.get_grøn_timer():.2f}")
        except:
            self.text_green.set("FAILED")
    def get_cts(self):
        try:
            self.text_cts.set(f"Pr. år: {self.wb.get_CTS_money():.2f}")
        except:
            self.text_cts.set("FAILED")
    def get_sc(self):
        try:
            self.text_sc.set(f"Pr. år: {self.wb.get_SC_money():.2f}")
        except:
            self.text_sc.set("FAILED")
    def get_feje(self):
        try:
            self.text_feje.set(f"Pr. år: {self.wb.get_feje_timer():.2f}")
        except:
            self.text_feje.set("FAILED")
    def get_all(self):
        self.get_blue()
        self.get_grey()
        self.get_green()
        self.get_cts()
        self.get_sc()
        self.get_feje()
        pass
    def clear_all(self):
        self.text_blue.set("Not calculated")
        self.text_grey.set("Not calculated")
        self.text_green.set("Not calculated")
        self.text_sc.set("Not calculated")
        self.text_cts.set("Not calculated")
        self.text_feje.set("Not calculated")
    def find_kunde_spec(self):
        self.kunde_specs = []
        spec_nr = 0
        for x in self.wb.aftale["D130":"D149"]:
            # print(x)
            if x[0].value == True:
                spec_nr += 1
                # self.kunde_specs.append(self.wb.aftale[f"C{x[0].row}"].value)
        return spec_nr
    # def show_kunde_spec(self):
    #     self.top = tk.Toplevel(width=700)
    #     self.top.title("Kundespecifikke ydelser")
    #     text = []
    #     for x in self.kunde_specs:
    #         text.append(f"{str(self.kunde_specs.index(x)+1)}: {str(x)}")
    #     tk.Message(self.top, text = '\n'.join(text)).pack()

if __name__ == "__main__":
    main = App()


