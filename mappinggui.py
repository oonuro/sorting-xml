import tkinter as tk
from tkinter import filedialog
from tkinter import ttk
from enginexml import Enginexml

class Mygui():

    def __init__(self, master):
        self.master = master
        master.title("xml Sorting Tool")
     
        # Calling class from enginefile.py
        self.engine = Enginexml()
        
                     
        master.resizable(False, False)
  
        self.excelfilepath = tk.StringVar()
        self.xmlfilepath = tk.StringVar()
        self.listexcel = []
        self.listlayout = []

        self.excelsheetnames = tk.StringVar()
        self.listlayoutname = tk.StringVar()
 
        # the object should be inside the frame (master) ==> (self.openfiles)

        # Frame
        self.openfiles = tk.LabelFrame(master, text= "Open File")
        self.openfiles.grid(row= 0, column= 0, sticky='EW', padx = 5, pady = 5, ipady= 5)
           
        self.mapping = tk.LabelFrame(master, text= "Mapping")
        self.mapping.grid(row= 1, column= 0, sticky="EW", padx = 5, pady = 5, ipady= 5)   
                 
        # Label
        self.filelabel = tk.Label(self.openfiles, text= "xml File: ")
        self.filelabel.grid(row= 0, column= 0, sticky= "W")

        self.layoutlabel = tk.Label(self.openfiles, text= "Layout Name: ")
        self.layoutlabel.grid(row= 1, column= 0, sticky= "W")

        self.filelabel = tk.Label(self.openfiles, text= "Excel File: ")
        self.filelabel.grid(row= 2, column= 0, sticky= "W")

        self.sheetname = tk.Label(self.openfiles, text= "Sheet Name: ")
        self.sheetname.grid(row= 3, column= 0, sticky= "W")

        # Entry
        self.openfile = tk.Entry(self.openfiles)
        self.openfile.grid(row=0, column=1, columnspan=7, sticky= "E",  padx=10, pady=5, ipadx= 120)

        self.excelfile = tk.Entry(self.openfiles)
        self.excelfile.grid(row=2, column=1, columnspan=7, sticky= "E",  padx=10, pady=5, ipadx= 120)
        
        # Buttons
        self.button1 = tk.Button(self.mapping, text = "Open xml File", command=self.selectxmlfile)
        self.button1.grid(row=0, column=1, sticky='EW',  padx=10, pady=5)

        self.button2 = tk.Button(self.mapping, text = "Open Excel File", command=self.selectexcelfile)
        self.button2.grid(row=0, column=2, sticky='EW',  padx=10, pady=5)

        self.button3 = tk.Button(self.mapping, text = "Generate Excel", command=self.datatoexcel)
        self.button3.grid(row=0, column=3, sticky="EW",  padx=10, pady=5)

        self.button3 = tk.Button(self.mapping, text = "Delete All", command=self.deletesheetdata)
        self.button3.grid(row=0, column=4, sticky="EW", padx=10, pady=5)
      
        # Combobox
        self.combolayout = ttk.Combobox(self.openfiles, textvariable = self.listlayoutname)
        self.combolayout.grid(row=1, column=1, columnspan=7, sticky= "WE",  padx=10, pady=5)

        self.combosheetname = ttk.Combobox(self.openfiles, textvariable = self.excelsheetnames)
        self.combosheetname.grid(row=3, column=1, columnspan=7, sticky= "WE",  padx=10, pady=5)

# xml file searching with filediaglog and pass the path variable to enginexml.py
    def selectxmlfile(self):
        self.openfile.delete(0, tk.END) 
        filetypes = (("xml *xml files", "*.conx"), ("All Files", "*.*"))
        self.xmlfilepath = filedialog.askopenfilename(title='xml .conx File', initialdir='/', filetypes=filetypes) 
        self.openfile.insert(tk.END, self.xmlfilepath)
        self.engine.openxmlfile(self.xmlfilepath)
        self.getlayoutname()
        return self.xmlfilepath

# Excel file searching with filediaglog and pass the path variable to enginexml.py
    def selectexcelfile(self):
        self.excelfile.delete(0, tk.END)
        filetypes = (("Excel File *vba", "*.xlsm"), ("Excel File", "*xlsx"), ("All Files", "*.*"))
        self.excelfilepath = filedialog.askopenfilename(title='Excel File', initialdir='/', filetypes=filetypes) 
        self.excelfile.insert(tk.END, self.excelfilepath)
        self.engine.openexcel(self.excelfilepath)
        self.getsheetname()
        return self.excelfilepath

# Writing data to excel 
    def datatoexcel(self):
        sheetname = self.excelsheetnames.get()
        self.engine.readxmlfile(sheetname)
    
# After open Excel get all sheetnames
    def getsheetname(self):
        self.engine.sheetnamesexcel()
        self.listexcel = self.engine.listsheetnames
        self.combosheetname.config(values=self.listexcel)
        print(self.listexcel) 
        return self.listexcel

# After open xml file get all layoutnames
    def getlayoutname(self):
        self.engine.layoutnames()
        self.listlayout = self.engine.listlayoutgui
        self.combolayout.config(values=self.listlayout)
        return self.listlayout

# Clean all data in the sheet of excel
    def deletesheetdata(self):
        name = self.excelsheetnames.get()
        self.engine.cleansheet(name)

    def deneme(self):
        sheetname = self.excelsheetnames.get()
        layoutname = self.listlayoutname.get()
        print(layoutname)
        self.engine.searchdatainexcel(sheetname, layoutname)

root = tk.Tk()
Gui = Mygui(root)
root.mainloop()
