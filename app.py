# Dependencies
import tkinter as tk
import tkinter.ttk as ttk
from tkinter import filedialog as fd


import pandas as pd
import xlwings as xw

# Global Variables


# Class to do all excel processing
class ExcelProcessor:
    # constructor
    def __init__(self, file_path):
        self.START = xw.Book(file_path)
        self.SLABS = self.START.sheets['START-SLABS']
        self.FW = self.START.sheets['START-FW']

        # List of plans that are available for the SLAB and FW
        self.slabPlans = [[]]
        self.fwPlans = [[]]

        # Keep track of the current row in each sheet
        self.slab_row = 0
        self.fw_row = 0

        self.SetPlans()


    # Function to get the SLAB and FW plan from the START excel sheet
    def SetPlans(self):
        for i in range(77, 200):
            if self.SLABS.range('F' + str(i)).value is not None:
                self.slabPlans.append(self.SLABS.range('F' + str(i) + ":I" + str(i)).value)

        for i in range(76, 200):
            if self.FW.range('F' + str(i)).value is not None:
                self.fwPlans.append(self.FW.range('F' + str(i) + ":I" + str(i)).value)

    # Function to add custom plans to the correct plans list
    def AddCustomOption(self, data, type):
        if type == 'slab':
            self.slabPlans.append(data)
        elif type == 'fw':
            self.fwPlans.append(data)

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
        # Have to redo how i loop through this data to add the options.
        if data[6]:
            for slab_option in data[6][0]:
                if slab_option.current() != -1 and slab_option.current() != 0:
                    self.slab_row += 1
                    self.SLABS.range('F' + str(9 + self.slab_row) + ":I" + str(9 + self.slab_row)).value = self.slabPlans[slab_option.current()]
            for fw_option in data[6][1]:
                if fw_option.current() != -1 and fw_option.current() != 0:
                    self.fw_row += 1
                    self.FW.range('F' + str(9 + self.fw_row) + ":I" + str(9 + self.fw_row)).value = self.fwPlans[fw_option.current()]    


        # Update the current row
        self.slab_row += 2
        self.fw_row += 2
    

# Main window which will have you select the start file to load.
class App(tk.Tk):
    # Constructor
    def __init__(self):
        super().__init__()

        self.title("Load Excel File")
        self.geometry("200x50")

        # Instance of the ExcelProcessor class
        self.document = None
        
        # TODO: Add entry fields for job#, subjob #, Supervisor, and pour type

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

    def __init__(self, parent):
        super().__init__(parent)

        self.title("STARTS Excel Helper")
        self.geometry("1000x500")

        lot_var= tk.StringVar()
        address_var= tk.StringVar()
        elv_var= tk.StringVar()

        # Variables to track the options
        self.opt_entries = []
        self.opt_count = 0
        
        gar_orr_options = ['LEFT', 'RIGHT']

        lot_label = tk.Label(self, text = 'Lot #', font=('calibre',10, 'bold')) 
        lot_entry = tk.Entry(self,textvariable = lot_var, font=('calibre',10,'normal'))
        
        address_label = tk.Label(self, text = 'Address', font = ('calibre',10,'bold'))
        address_entry=tk.Entry(self, textvariable = address_var, font = ('calibre',10,'normal'))

        # Todo: Add the elevation to the main window
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
        add_cust_opt_btn=tk.Button(self,text = 'Custom Option', command = self.AddCustomOption)
        add_lot_btn=tk.Button(self,text = 'Add Lot', command = lambda: self.AddLot(lot_var.get(), address_var.get().upper(),
                                                                                    elv_var.get().upper(), gar_orr_combo.get(),
                                                                                    slab_plan_combo.current(), fw_plan_combo.current()))

        # Placing the entries and button in a grid
        lot_label.grid(row=0,column=0)
        address_label.grid(row=1,column=0)
        elv_label.grid(row=2,column=0)
        gar_orr_label.grid(row=3,column=0)

        lot_entry.grid(row=0,column=1)
        address_entry.grid(row=1,column=1)
        elv_entry.grid(row=2,column=1)
        gar_orr_combo.grid(row=3,column=1)

        slab_plan_label.grid(row=0,column=3)
        slab_plan_combo.grid(row=1,column=3)
        fw_plan_label.grid(row=0,column=4)
        fw_plan_combo.grid(row=1,column=4)

        add_cust_opt_btn.grid(row=1,column=2)
        add_opt_btn.grid(row=0,column=2)
        add_lot_btn.grid(row=20,column=4, pady = 200)
        done_btn.grid(row=20,column=5, pady = 200)

    # Function to add the entry fields into excel
    def AddLot(self, lot, address, elv, gar_orr, slab_plan, fw_plan):
        # creating a list of the values to be added to the excel sheet
        values = [lot, address, gar_orr, slab_plan,
                  fw_plan, elv, self.opt_entries]

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

    def AddCustomOption(self):
        window = CustomOptionWindow(self)

    # Function to close the window
    def Close(self):
        self.destroy()

# The window for adding custom options to the list
class CustomOptionWindow(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)

        self.title("Add Custom Option") 
        self.geometry("600x200")

        # creating the variables entry fields
        plan_var= tk.StringVar()
        sage_elv_var= tk.StringVar()
        elv_var= tk.StringVar()
        cost_var= tk.StringVar()

        # creating the labels and entry fields
        plan_label = tk.Label(self, text = 'Plan', font=('calibre',10, 'bold'))
        plan_entry = tk.Entry(self,textvariable = plan_var, font=('calibre',10,'normal'))

        sage_elv_label = tk.Label(self, text = 'SAGE ELV', font = ('calibre',10,'bold'))
        sage_elv_entry=tk.Entry(self, textvariable = sage_elv_var, font = ('calibre',10,'normal'))

        elv_label = tk.Label(self, text = 'ELV', font = ('calibre',10,'bold'))
        elv_entry=tk.Entry(self, textvariable = elv_var, font = ('calibre',10,'normal'))

        cost_label = tk.Label(self, text = 'Cost', font = ('calibre',10,'bold'))
        cost_entry=tk.Entry(self, textvariable = cost_var, font = ('calibre',10,'normal'))

        # creating a button using the widget
        slab_done_btn=tk.Button(self,text = 'Add to Slab', command = lambda : self.AddNewSlabOption(plan_var.get(), sage_elv_var.get(), elv_var.get(), cost_var.get()))
        fw_done_btn=tk.Button(self,text = 'Add to FW', command = lambda : self.AddNewFWOption(plan_var.get(), sage_elv_var.get(), elv_var.get(), cost_var.get()))

        # Placing the entries and button in a grid
        plan_label.grid(row=0,column=0)
        plan_entry.grid(row=1,column=0)
        sage_elv_label.grid(row=0,column=1)
        sage_elv_entry.grid(row=1,column=1)
        elv_label.grid(row=0,column=2)
        elv_entry.grid(row=1,column=2)
        cost_label.grid(row=0,column=3)
        cost_entry.grid(row=1,column=3)
        slab_done_btn.grid(row=2,column=2)
        fw_done_btn.grid(row=2,column=3)


    # Function to add the entry fields into the list of slab options
    def AddNewSlabOption(self, plan, sage_elv, elv, cost):
        # creating a list of the values to be added to the excel sheet
        values = [plan, sage_elv, elv, cost]

        # adding the values to the excel sheet
        app.document.AddCustomOption(values, 'slab')

        # closing the window
        self.destroy()    

    # Function to add the entry fields into the list of slab options
    def AddNewFWOption(self, plan, sage_elv, elv, cost):
        # creating a list of the values to be added to the excel sheet
        values = [plan, sage_elv, elv, cost]

        # adding the values to the excel sheet
        app.document.AddCustomOption(values, 'fw')

        # closing the window
        self.destroy()  



app = App()
app.mainloop()