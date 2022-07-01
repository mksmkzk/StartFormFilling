# Dependencies
import tkinter as tk
import tkinter.ttk as ttk
from tkinter import filedialog as fd


import pandas as pd
import xlwings as xw

# Global Variables


# Class to do all excel processing
class ExcelProcessor:
    # Variables
    START = None
    SLABS = None
    FW = None

    # List of plans that are available for the SLAB and FW
    slabPlans = []
    fwPlans = []

    # Keep track of the current row in each sheet
    slab_row = 0
    fw_row = 0

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

    # Function to input the data into the excel sheet
    def InputData(self, data):

        # Set the lot number
        self.SLABS.range('C' + str(9 + self.slab_row)).value = data[0]
        self.FW.range('C' + str(9 + self.fw_row)).value = data[0]

        # Set the address
        self.SLABS.range('D' + str(9 + self.slab_row)).value = data[1]
        self.FW.range('D' + str(9 + self.fw_row)).value = data[1]

        # Set the garage orrientation
        self.SLABS.range('E' + str(9 + self.slab_row)).value = data[2]
        self.FW.range('E' + str(9 + self.fw_row)).value = data[2]

        # Set the slab plan
        if data[3] != -1:
            self.SLABS.range('F' + str(9 + self.slab_row) + ":I" + str(9 + self.slab_row)).value = self.slabPlans[data[3]]

        # Set the fw plan
        if data[4] != -1:
            self.FW.range('F' + str(9 + self.fw_row) + ":I" + str(9 + self.fw_row)).value = self.fwPlans[data[4]]


        # Set the elevation
        self.SLABS.range('H' + str(9 + self.slab_row)).value = data[5]
        self.FW.range('H' + str(9 + self.fw_row)).value = data[5]

        # Set the options
        print(data[6])
        if data[6]:
            for slab_option in data[6][0]:
                if slab_option.current() != -1:
                    self.slab_row += 1
                    self.SLABS.range('F' + str(9 + self.slab_row) + ":I" + str(9 + self.slab_row)).value = self.slabPlans[slab_option.current()]
            for fw_option in data[6][1]:
                if fw_option.current() != -1:
                    self.fw_row += 1
                    self.FW.range('F' + str(9 + self.fw_row) + ":I" + str(9 + self.fw_row)).value = self.fwPlans[fw_option.current()]    


        # Update the current row
        self.slab_row += 2
        self.fw_row += 2
    

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
    # Variables
    opt_entries = []

    opt_count = 0

    def __init__(self, parent):
        super().__init__(parent)

        self.title("STARTS Excel Helper")
        self.geometry("1000x500")

        lot_var= tk.StringVar()
        address_var= tk.StringVar()
        elv_var= tk.StringVar()
        
        gar_orr_options = ['LEFT', 'RIGHT']

        lot_label = tk.Label(self, text = 'Lot #', font=('calibre',10, 'bold')) 
        lot_entry = tk.Entry(self,textvariable = lot_var, font=('calibre',10,'normal'))
        
        address_label = tk.Label(self, text = 'Address', font = ('calibre',10,'bold'))
        address_entry=tk.Entry(self, textvariable = address_var, font = ('calibre',10,'normal'))

        elv_label = tk.Label(self, text = 'ELV', font = ('calibre',10,'bold'))
        elv_entry=tk.Entry(self, textvariable = elv_var, font = ('calibre',10,'normal'))

        gar_orr_label = tk.Label(self, text = 'Garage Orientation', font = ('calibre',10,'bold'))
        gar_orr_combo=ttk.Combobox(self, values = gar_orr_options)

        slab_plan_label = tk.Label(self, text = 'SLAB Plan', font = ('calibre',10,'bold'))
        slab_plan_combo=ttk.Combobox(self, values = app.document.slabPlans, width= 40)

        fw_plan_label = tk.Label(self, text = 'FW Plan', font = ('calibre',10,'bold'))
        fw_plan_combo=ttk.Combobox(self, values = app.document.fwPlans, width= 40)
        
        # creating a button using the widget
        done_btn=tk.Button(self,text = 'Done', command = self.Close)
        add_opt_btn=tk.Button(self,text = 'Add Options', command = self.AddOptions)
        add_lot_btn=tk.Button(self,text = 'Add Lot', command = lambda: self.AddLot(lot_var, address_var, elv_var, gar_orr_combo.get(), slab_plan_combo.current(), fw_plan_combo.current()))

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
        add_lot_btn.grid(row=20,column=4, pady = 200)
        done_btn.grid(row=20,column=5, pady = 200)

    # Function to add the entry fields into excel
    def AddLot(self, lot, address, elv, gar_orr, slab_plan, fw_plan):
        # creating a list of the values to be added to the excel sheet
        values = [lot.get(), address.get().upper(), gar_orr, 
                  slab_plan, fw_plan, elv.get().upper(), self.opt_entries]

        # adding the values to the excel sheet
        app.document.InputData(values)


    # Function to add new option to the entry
    def AddOptions(self):
        MAX_OPTIONS = 7
        if self.opt_count < MAX_OPTIONS:
            opt_label = tk.Label(self, text = 'Option ' + str(self.opt_count + 1), font = ('calibre',10,'bold'))
            opt_label.grid(row=2 + self.opt_count,column=2)
    
            self.opt_entries.append([ttk.Combobox(self, values= app.document.slabPlans, width= 40), ttk.Combobox(self, values= app.document.fwPlans, width= 40)])
            self.opt_entries[-1][0].grid(row=2+ self.opt_count,column=3)
            self.opt_entries[-1][1].grid(row=2+ self.opt_count,column=4)
            self.opt_count += 1

    # Function to close the window
    def Close(self):
        self.destroy()

app = App()
app.mainloop()