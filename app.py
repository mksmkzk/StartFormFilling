# Dependencies
import this
import tkinter as tk
import tkinter.ttk as ttk
from tkinter import filedialog as fd


import pandas as pd
import xlwings as xw

# Global Variables
opt_count = 0

# Class to do all excel processing
class ExcelProcessor:
    # Variables
    START = None
    SLABS = None
    FW = None

    slabPlans = []
    fwPlans = []

    # constructor
    def __init__(self, file_path):
        self.START = xw.Book(file_path)
        self.SLABS = self.START.sheets['START-SLABS']
        self.FW = self.START.sheets['START-FW']
        self.SetPlans()

    # Function to get the SLAB and FW plan from the START excel sheet
    def SetPlans(self):
        for i in range(77, 200):
            if self.SLABS.range('F' + str(i)).value is not None:
                self.slabPlans.append(self.SLABS.range('F' + str(i) + ":I" + str(i)).value)

        for i in range(76, 200):
            if self.FW.range('F' + str(i)).value is not None:
                self.fwPlans.append(self.FW.range('F' + str(i) + ":I" + str(i)).value)

# Main window which will have you select the start file to load.
class App(tk.Tk):

    document = None

    def __init__(self):
        super().__init__()

        self.title("Load Excel File")
        self.geometry("200x50")

        load_file_label = tk.Label(self, text="Select START to Load")
        load_file_label.pack()
        load_file_btn=tk.Button(self,text = 'Load Excel File', command = self.OpenFile)
        load_file_btn.pack()
    
    # Function to load the excel file
    def OpenFile(self):
        # opening the file
        filename = fd.askopenfilename(filetypes=[('Excel', '*.xlsx')])
        self.document = ExcelProcessor(filename)

        # creating the main window
        window = ExcelWindow(self)
        window.grab_set()
        self.withdraw()

# This is the window which will have the entry fields needed for the START file.
class ExcelWindow(tk.Toplevel):
    opt_entries = []

    def __init__(self, parent):
        super().__init__(parent)

        self.title("STARTS Excel Helper")
        self.geometry("1000x500")

        lot_var= tk.StringVar()
        address_var= tk.StringVar()
        
        gar_orr_options = ['LEFT', 'CENTER', 'RIGHT']

        lot_label = tk.Label(self, text = 'Lot #', font=('calibre',10, 'bold')) 
        lot_entry = tk.Entry(self,textvariable = lot_var, font=('calibre',10,'normal'))
        
        address_label = tk.Label(self, text = 'Address', font = ('calibre',10,'bold'))
        address_entry=tk.Entry(self, textvariable = address_var, font = ('calibre',10,'normal'))

        gar_orr_label = tk.Label(self, text = 'Garage Orientation', font = ('calibre',10,'bold'))
        gar_orr_combo=ttk.Combobox(self, values= gar_orr_options)

        slab_plan_label = tk.Label(self, text = 'SLAB Plan', font = ('calibre',10,'bold'))
        slab_plan_combo=ttk.Combobox(self, values= app.document.slabPlans, width= 40)

        fw_plan_label = tk.Label(self, text = 'FW Plan', font = ('calibre',10,'bold'))
        fw_plan_combo=ttk.Combobox(self, values= app.document.fwPlans, width= 40)
        
        # creating a button using the widget
        done_btn=tk.Button(self,text = 'Done', command = self.Close)
        add_opt_btn=tk.Button(self,text = 'Add Options', command = self.AddOptions)

        # Placing the entries and button in a grid
        lot_label.grid(row=0,column=0)
        lot_entry.grid(row=1,column=0)
        address_label.grid(row=0,column=1)
        address_entry.grid(row=1,column=1)
        gar_orr_label.grid(row=0,column=2)
        gar_orr_combo.grid(row=1,column=2)
        slab_plan_label.grid(row=0,column=3)
        slab_plan_combo.grid(row=1,column=3)
        fw_plan_label.grid(row=0,column=4)
        fw_plan_combo.grid(row=1,column=4)
        add_opt_btn.grid(row=2,column=1)
        done_btn.grid(row=20,column=5, pady = 200)

    # Function to add new option to the entry
    def AddOptions(self):
        global opt_count
        MAX_OPTIONS = 6
        if opt_count < MAX_OPTIONS:
            opt_label = tk.Label(self, text = 'Option ' + str(opt_count + 1), font = ('calibre',10,'bold'))
            opt_label.grid(row=2 + opt_count,column=2)
    
            self.opt_entries.append([ttk.Combobox(self, values= app.document.slabPlans, width= 40), ttk.Combobox(self, values= app.document.fwPlans, width= 40)])
            self.opt_entries[-1][0].grid(row=2+ opt_count,column=3)
            self.opt_entries[-1][1].grid(row=2+ opt_count,column=4)
            opt_count += 1

    # Function to close the window
    def Close(self):
        # print(list(map(lambda x: x.get(), self.opt_entries)))
        self.destroy()

app = App()
app.mainloop()