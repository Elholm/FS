#test_layout

from test_dict import App,fs_ejd

import pandas as pd

import tkinter as tk
from tkinter import filedialog as fd

import docx

import pylatex as pl

from test_dict import App, fs_ejd



## Functions

def create_pdf(self):
    #Read the template for PDF
    wb.to_pdf()
    return

def load_file(self):
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


def get_nøgletal(self):
    self.get_blue()
    self.get_grey()
    self.get_green()
    self.get_cts()
    self.get_sc()
    self.get_feje()
    return

def send_to_sql(self):
    wb.to_sql()
    return

def create_csv(self):
    wb.create_csv("path")
    return

def send_to_esdh(self):
    wb.to_sql()
    #is this even possible?
    return


