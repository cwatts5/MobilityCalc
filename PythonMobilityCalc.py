# -*- coding: utf-8 -*-
"""
Created on Thu Feb 21 09:43:14 2019

@author: wattsc2
"""

# imports
import sys
from PyQt5.QtCore import pyqtSlot
from PyQt5.QtWidgets import QApplication, QDialog, QProgressBar
from PyQt5.uic import loadUi
import pandas as pd
import tkinter as tk
import numpy as np
import re
from tkinter import filedialog


class MobilityCalc(QDialog):
    def __init__(self):
        super(MobilityCalc, self).__init__()
        loadUi('MobilityCalcMul.ui', self)  # loads UI
        self.setWindowTitle('Transistor Mobilty Analyzer')  # sets window title
        self.ChooseFile.clicked.connect(self.openfiledialog)
        self.Calculate.clicked.connect(self.calcfunction)
        self.ExportFile.clicked.connect(self.savefunction)

    @pyqtSlot()
    def openfiledialog(self):
        root = tk.Tk()
        root.withdraw()
        global names
        FN = filedialog.askopenfilenames(parent=root, title='Choose a file')  # opens file dialog
        skip = "\n"  # supposed to be skipline. Probably fix later.
        names = [fn.split(r"/")[-1] for fn in FN]  # Used to name the excel sheets in Savefunction
        Cfbxtxt = skip.join(names)  # makes tuple (FN) into usable string
        self.ChosenFilesBox.setText(Cfbxtxt)  # sets text box to be names of files.
        global dfs  # declares dataframe list as global variable
        dfs = []  # creates empty list
        for file_ in FN:
            df = pd.read_csv(file_, skiprows=5)  # reads in csv data skips top 5 rows
            dfs.append(df)  # creates a list of the dataframes

    def calcfunction(self):
        global results
        results = []
        for i in range(0, len(dfs)):
            Vg = dfs[i]['Vg']  # finds Vg values
            Vgd = Vg.diff(periods=2)  # differentiates Vg
            Idvalues = dfs[i].filter(regex='Id')  # finds all Id values
            Iddiff = Idvalues.diff(periods=2)  # differentiates Id values.
            Idvaluessqrt = np.sqrt(Idvalues)  # Finds sqrt of Id values
            Idsqrtdiff = Idvaluessqrt.diff(periods=2)  # finds 2 point difference in sqrt(Id) values
            X = Idvalues.columns  # logs names of Id columns
            X = ''.join(str(e) for e in X)  # creates string from X for finding Vd
            X = [float(s) for s in re.findall(r'-?\d+\.?\d*', X)]  # extracts Vd values from X
            X = np.asarray(X)  # creates an array from Vd values
            LinMobNC = Iddiff.div(Vgd, axis=0)  # the derivative of any I drain values and gate Voltage.
            SatMobNC = Idsqrtdiff.div(Vgd, axis=0)  # the derivative of any sqrt(I) drain values and gate Voltage.
            SatMobNC = SatMobNC ** 2  # squares the derivative of saturation mobility (no constants)
            l = float(self.LBox.text())  # reads in data from Length box as float
            w = float(self.WBox.text())  # reads in data from width box as float
            c = float(self.Cbox.text()) * 1E-9  # reads in data from Capacitance box as float
            CLM = l / (w * c)  # calculates constants for linear mobility based on l,w,c values
            CSM = (2 * l) / (w * c)  # calculates constants for saturation mobility based on l,w,c values
            CLMX = np.true_divide(CLM, X)  # divides by obtained Vd values
            LinMob = LinMobNC.multiply(CLMX)  # combines constants and LinmobNC
            LinMob = LinMob.rename(
                columns=lambda x: x.replace('Id', 'Linear Mobility cm^2/V-sec'))  # Renames any values
            LinMob = LinMob.shift(-1)  # shifts values up one row
            SatMob = SatMobNC.multiply(CSM)  # combines constants and SatmobNC
            SatMob = SatMob.rename(
                columns=lambda x: x.replace('Id', 'Saturation Mobility cm^2/V-sec'))  # Renames any values
            SatMob = SatMob.shift(-1)  # shifts values up one row
            A = pd.concat([dfs[i], LinMob, SatMob], axis=1)  # creates a new dataframe for each dfs iteration
            results.append(A)  # adds the results of the for loop to the empty list A

    def savefunction(self):  # function that occurs when export file is pressed.
        export_file_path = filedialog.asksaveasfilename(defaultextension='.xlsx')  # names export path and adds .xlsx
        writer = pd.ExcelWriter(export_file_path)  # tells panda it is writing to excel
        for i, df in enumerate(results):  # loop for adding multiple dataframes to output
            df.to_excel(writer, index=False, sheet_name=names[i])  # create a new sheet for each output
        writer.save()  # saves values


app = QApplication(sys.argv)
widget = MobilityCalc()
widget.show()
sys.exit(app.exec_())
