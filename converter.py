import tkinter as tk
import threading
import random
import os
from tkinter import CENTER, END, FLAT, HORIZONTAL, NO, W, StringVar, Toplevel, filedialog, messagebox,ttk,Canvas
from PIL import Image, ImageTk
from pathlib import Path
from time import strftime
from db import Database

db = Database('converter.db')
ROOT_DIR = os.getcwd() + "\\"
SUPPLIER_LIST = []

class Application(tk.Frame):    

    def __init__(self, master):
        
        super().__init__(master)        
        self.master = master 
        
        # self.master.withdraw()
        # self.splash()
        self.master.protocol("WM_DELETE_WINDOW", self.closeApp)

        style = ttk.Style(self.master)
        self.master.tk.call("source", ROOT_DIR + "theme\\forest-dark.tcl")
        style.theme_use("forest-dark")

        w = 700 # width for the Tk root
        h = 500 # height for the Tk root
        ws = self.winfo_screenwidth() # width of the screen
        hs = self.winfo_screenheight() # height of the screen
        # calculate x and y coordinates for the Tk root window
        x = (ws/2) - (w/2)
        y = (hs/2) - (h/2)

        # set the dimensions of the screen and where it is placed
        self.master.geometry('%dx%d+%d+%d' % (w, h, x, y))
        self.master.resizable(0, 0) #
        self.master.iconbitmap("converter.ico")
        self.master.title('CWO Proforma Converter') 
        self.master.config()   
        #BACKGROUND
        # self.background_img = Image.open("forest_dark.png")
        # self.background_img = self.background_img.resize((w, h), Image.Resampling.LANCZOS)
        # self.back_image     = ImageTk.PhotoImage(self.background_img)
        # self.bg_label       = tk.Label(self.master, image = self.back_image )
        # self.bg_label.place(x=0, y=0, relwidth=1, relheight=1)
        # self.bg_label.image = self.back_image


        self.createWidget()

    def createWidget(self):        
        
        #MENU
        menubar  = tk.Menu(self.master)
        filemenu = tk.Menu(menubar, tearoff=0)
        filemenu.add_command(label="Item Code Setup", command=ItemSetup)
        filemenu.add_command(label="Supplier Setup",  command=SupplierSetup)
        filemenu.add_separator()
        filemenu.add_command(label="Exit", command=self.closeApp)        
        menubar.add_cascade(label="Masterfile", menu=filemenu) 

        toolsmenu = tk.Menu(menubar, tearoff=0)
        toolsmenu.add_command(label="Convert PSI",  command=ConvertPSI)
        toolsmenu.add_command(label="Split(PDF) Page",  command=SplitPage)
        menubar.add_cascade(label="Tools", menu=toolsmenu) 

        canvas1 = Canvas(self.master,highlightthickness=0,highlightbackground="black")
        canvas2 = Canvas(self.master,highlightthickness=0,highlightbackground="black")
        points = [
                    random.randrange(0, 300),
                    random.randrange(0, 200),
                    random.randrange(0, 300),
                    random.randrange(0, 300),
                    random.randrange(0, 200),   
                    random.randrange(0, 300),
                    random.randrange(0, 200),
                    random.randrange(0, 300),
                    random.randrange(0, 200),
                    random.randrange(0, 300)
                 ]

        canvas1.create_polygon(points, fill='white') ##217346
        canvas2.create_text(85, 40, anchor=W, font= ("Pristina", 20,'bold'),text="Cash With Order",justify='center',fill='white')
        canvas2.create_text(100, 60, anchor=W, font=("Corbel",12),text="Pro-forma Converter",justify='center',fill='white')
        canvas2.create_text(30, 90, anchor=W, font=("Corbel",10),text="1. COSMETIQUE ASIA CORPORATION",justify='left',fill='white')
        canvas2.create_text(30, 110, anchor=W, font=("Corbel",10),text="2. ECOSSENTIAL FOODS CORPORATION",justify='left',fill='white')
        canvas2.create_text(30, 130, anchor=W, font=("Corbel",10),text="3. FOOD CHOICE CORPORATION",justify='left',fill='white')
        canvas2.create_text(30, 150, anchor=W, font=("Corbel",10),text="4. FOOD INDUSTRIES, INC.",justify='left',fill='white')
        canvas2.create_text(30, 170, anchor=W, font=("Corbel",10),text="5. GREEN CROSS, INC.",justify='left',fill='white')
        canvas2.create_text(30, 190, anchor=W, font=("Corbel",10),text="6. INTELLIGENT SKIN CARE, INC.",justify='left',fill='white')
        canvas2.create_text(30, 210, anchor=W, font=("Corbel",10),text="7. JS UNITRADE MDSE., INC.",justify='left',fill='white')
        canvas2.create_text(30, 230, anchor=W, font=("Corbel",10),text="8. MEAD JOHNSON NUTRITION",justify='left',fill='white')
        canvas2.create_text(30, 250, anchor=W, font=("Corbel",10),text="9. MONDELEZ PHILIPPINES, INC.",justify='left',fill='white')
        canvas2.create_text(30, 270, anchor=W, font=("Corbel",10),text="10. SUYEN CORPORATION",justify='left',fill='white')
        
        
        canvas1.place(x=5, y=5)
        canvas2.place(x=320,y=5)
        canvas1.config(width=300,height=300)
        canvas2.config(width=322,height=300)

        self.currenttime_label = tk.Label(self.master,text='', font=('Segoe UI Light bold',16))
        self.currenttime_label.place(x=10,y=420)
        self.clock_image = Image.open("windows-clock.png")
        self.clock_image = self.clock_image.resize((12, 12), Image.Resampling.LANCZOS)
        self.clock       = ImageTk.PhotoImage(self.clock_image)
        self.clock_label = tk.Label(self.master, image = self.clock)
        self.clock_label.place(x=140,y= 420)
        self.clock_label.image = self.clock
        self.currentdate_label = tk.Label(self.master,text='',font=('Segoe UI Light bold',10))
        self.currentdate_label.place(x=10,y=450)
        self.version_label = tk.Label(self.master,text='version 1.2.4',font=('Segoe UI Light bold',10))
        self.version_label.place(x=610,y=450)
        self.master.config(menu=menubar)
        self.currentTime()

    def currentTime(self):
        self.current_time = strftime('%I:%M:%S %p')  #'%I:%M:%S %p'
        self.current_date = strftime('%A, %d %B')
        self.currenttime_label['text'] = self.current_time
        self.currentdate_label['text'] = self.current_date
        self.master.after(1000,self.currentTime)

    def closeApp(self):
        if messagebox.askokcancel('Exit','Do yo want to close the application?'):
            try:
                if submit_thread.is_alive():                
                    messagebox.showwarning('Exit','Unable to exit application while conversion is ongoing!')
                else:
                    self.master.destroy()
            except NameError:
                self.master.destroy()

    def splash(self):
        sp = Toplevel()
        w = 650
        h = 320
        ws = sp.winfo_screenwidth() 
        hs = sp.winfo_screenheight()
        x = (ws/2) - (w/2)
        y = (hs/2) - (h/2)
        sp.geometry('%dx%d+%d+%d' % (w, h, x, y))
        sp.config()   
        sp.overrideredirect(1)

        canvas1 = Canvas(sp,highlightthickness=0,highlightbackground="black")
        canvas2 = Canvas(sp,highlightthickness=0,highlightbackground="black")
        points = [
                    random.randrange(0, 300),
                    random.randrange(0, 200),
                    random.randrange(0, 300),
                    random.randrange(0, 300),
                    random.randrange(0, 200),   
                    random.randrange(0, 300),
                    random.randrange(0, 200),
                    random.randrange(0, 300),
                    random.randrange(0, 200),
                    random.randrange(0, 300)
                 ]

        canvas1.create_polygon(points, fill='white') ##217346
        canvas2.create_text(85, 40, anchor=W, font= ("Pristina", 20,'bold'),text="Cash With Order",justify='center',fill='white')
        canvas2.create_text(100, 60, anchor=W, font=("Corbel",12),text="Pro-forma Converter",justify='center',fill='white')
        canvas2.create_text(30, 90, anchor=W, font=("Corbel",10),text="1. COSMETIQUE ASIA CORPORATION",justify='left',fill='white')
        canvas2.create_text(30, 110, anchor=W, font=("Corbel",10),text="2. ECOSSENTIAL FOODS CORPORATION",justify='left',fill='white')
        canvas2.create_text(30, 130, anchor=W, font=("Corbel",10),text="3. FOOD CHOICE CORPORATION",justify='left',fill='white')
        canvas2.create_text(30, 150, anchor=W, font=("Corbel",10),text="4. FOOD INDUSTRIES, INC.",justify='left',fill='white')
        canvas2.create_text(30, 170, anchor=W, font=("Corbel",10),text="5. GREEN CROSS, INC.",justify='left',fill='white')
        canvas2.create_text(30, 190, anchor=W, font=("Corbel",10),text="6. INTELLIGENT SKIN CARE, INC.",justify='left',fill='white')
        canvas2.create_text(30, 210, anchor=W, font=("Corbel",10),text="7. JS UNITRADE MDSE., INC.",justify='left',fill='white')
        canvas2.create_text(30, 230, anchor=W, font=("Corbel",10),text="8. MEAD JOHNSON NUTRITION",justify='left',fill='white')
        canvas2.create_text(30, 250, anchor=W, font=("Corbel",10),text="9. MONDELEZ PHILIPPINES, INC.",justify='left',fill='white')
        canvas2.create_text(30, 270, anchor=W, font=("Corbel",10),text="10. SUYEN CORPORATION",justify='left',fill='white')
        
        
        canvas1.place(x=5, y=5)
        canvas2.place(x=320,y=5)
        canvas1.config(width=300,height=300)
        canvas2.config(width=322,height=300)
        
        self.progress = ttk.Progressbar(sp,orient=HORIZONTAL,length=638,mode='determinate',maximum=110)  
        self.progress.place(x=5,y=310)
        self.progress.start(15)
        sp.after(2000,lambda: self.backToMain(sp))
        # self.backToMain(sp) 

    
    def backToMain(self,sp):   
        sp.destroy()
        self.master.deiconify()

class ItemSetup(Toplevel):

    def __init__(self):
        Toplevel.__init__(self)

        self.protocol("WM_DELETE_WINDOW", self.destroy)
        self.itemid_text       = StringVar()
        self.itemcode_text     = StringVar()
        self.supplierid_text   = StringVar()
        self.suppliername_text = StringVar()

        supplierIds = db.fetch_supplierids()

        w   = 520
        h   = 400 
        ws  = self.winfo_screenwidth()
        hs  = self.winfo_screenheight()
        x   = (ws/2) - (w/2)
        y   = (hs/2) - (h/2)
        self.geometry('%dx%d+%d+%d' % (w, h, x, y))
        self.resizable(0, 0)
        self.iconbitmap("parcel-box-package.ico")
        self.title('Item Setup')      

        # Supplier ID
        supplierid_label = tk.Label(self, text='Supplier :', font=('bold', 10))
        supplierid_label.place(x=5,y=12)
        supplierid_cbo = ttk.Combobox(self, textvariable=self.supplierid_text, state='readonly', width=16, value=supplierIds)
        supplierid_cbo.bind('<<ComboboxSelected>>',self.selectSupplier) 
        supplierid_cbo.place(x=100, y=14)
        # Supplier Name
        suppliername_label = tk.Label(self, textvariable=self.suppliername_text, font=('bold', 10))
        suppliername_label.place(x=255,y=23)
        # Item ID        
        itemid_label = tk.Label(self, text='Item ID :',  font=('bold', 10))
        itemid_label.place(x=5,y=50)
        self.itemid_entry = tk.Entry(self,textvariable=self.itemid_text , relief='solid', width=20, highlightcolor= "black")
        self.itemid_entry.place(x=100,y=50)
        # Item Code
        itemcode_label = tk.Label(self, text='Item Code :', font=('bold', 10))
        itemcode_label.place(x=260,y=50)
        self.itemcode_entry = tk.Entry(self, textvariable=self.itemcode_text, relief='solid', highlightcolor= "black")
        self.itemcode_entry.place(x=344, y=50)    

        style = ttk.Style(self)
        style.map('Treeview',background=[('selected', "#217346")])
        style.configure('Treeview', rowheight=17)

        itemlist_frame = tk.Frame(self,width=500,height=300)
        itemlist_frame.pack(pady=90)
        itemlist_frame.pack_propagate(0) 
        itemlist_scroll = ttk.Scrollbar(itemlist_frame)
        itemlist_scroll.pack(side="right", fill="y")
        self.itemlist_tree = ttk.Treeview(itemlist_frame, yscrollcommand=itemlist_scroll.set, selectmode="extended")
        self.itemlist_tree.pack()
        itemlist_scroll.config(command=self.itemlist_tree.yview)
        self.itemlist_tree['columns'] = ("Item ID", "Item Code", "Supplier ID","Supplier Name")
        # Format Our Columns
        self.itemlist_tree.column("#0", width=0, stretch=NO)
        self.itemlist_tree.column("Item ID", anchor=CENTER, minwidth=40, width=50,stretch=NO)
        self.itemlist_tree.column("Item Code", anchor=CENTER, minwidth=40, width=100,stretch=NO)
        self.itemlist_tree.column("Supplier ID", anchor=CENTER, minwidth=80, width=70,stretch=NO)
        self.itemlist_tree.column("Supplier Name", anchor=CENTER, minwidth=240, width=260,stretch=NO)
        # Create Headings
        self.itemlist_tree.heading("#0", text="", anchor=W)
        self.itemlist_tree.heading("Item ID", text="Item ID", anchor=W)
        self.itemlist_tree.heading("Item Code", text="Item Code", anchor=CENTER)
        self.itemlist_tree.heading("Supplier ID", text="Supplier ID", anchor=CENTER)
        self.itemlist_tree.heading("Supplier Name", text="Supplier Name", anchor=CENTER)
        # Create Striped Row Tags
        self.itemlist_tree.tag_configure('oddrow', background="#4a4949")
        self.itemlist_tree.tag_configure('evenrow', background="#1a8a4c")
        #Buttons
        add_btn = tk.Button(self, text='Add Item', width=13, relief=FLAT, command=lambda:self.addItem(self),bg="#217346")
        add_btn.place(x=15,y=320)
        remove_btn = tk.Button(self, text='Remove Item',  width=13, relief=FLAT, command=lambda:self.removeItem(self),bg="#217346")
        remove_btn.place(x=140,y=320)
        update_btn = tk.Button(self, text='Update Item', width=13, relief=FLAT, command=lambda:self.updateItem(self),bg="#217346")
        update_btn.place(x=265,y=320)
        clear_btn = tk.Button(self, text='Clear Input', width=13, relief=FLAT, command=self.clearText,bg="#217346")
        clear_btn.place(x=388,y=320)        
        #Bind the treeview
        self.itemlist_tree.bind('<ButtonRelease-1>', self.selectItem) 
        
    def selectSupplier(self,event):
        supname = db.fetch_supplier_by_id(self.supplierid_text.get())
        self.suppliername_text.set(supname[0])
        self.clearText()
        self.populateItems()

    def clearText(self):
        self.itemid_entry.delete(0,END)
        self.itemcode_entry.delete(0,END)

    def populateItems(self):
        for record in self.itemlist_tree.get_children():
            self.itemlist_tree.delete(record)
        count = 0
        items = db.fetch_items_by_supid(self.supplierid_text.get())
        
        for i in items:
            if count % 2 == 0:
                self.itemlist_tree.insert(parent='', index='end', iid=count, text='', values=(i[0], i[1], i[2], i[3]), tags=('evenrow',))
            else:
                self.itemlist_tree.insert(parent='', index='end', iid=count, text='', values=(i[0], i[1], i[2], i[3]), tags=('oddrow',))
            count += 1

    def selectItem(self,event):
        try :
            self.itemid_entry.delete(0,END)
            self.itemcode_entry.delete(0,END)
            # Grab record Number
            selected = self.itemlist_tree.focus()
            # Grab record values
            values = self.itemlist_tree.item(selected, 'values')
            # outpus to entry boxes
            self.itemid_entry.insert(0, values[0])
            self.itemcode_entry.insert(0, values[1])
        
        except IndexError :
            pass

    def addItem(self,itemWin):
        if self.supplierid_text.get() == "" or self.itemcode_text.get() == "":
            messagebox.showerror("Required","Please fill in required field!",parent=itemWin) 
        else:
            if messagebox.askyesno("Item Setup","Add new item code?",parent=itemWin) :
                insert = db.insert_item(self.itemcode_text.get(),self.supplierid_text.get())
                if insert == "success":
                    messagebox.showinfo("Sucess","Item code is added!",parent=itemWin)
                elif insert == "duplicate":
                    messagebox.showerror("Duplicate","Item code already exists!",parent=itemWin)
            else:
                messagebox.showinfo("Info", "Item code not added!",parent=itemWin)

        self.clearText()
        self.populateItems() 

    def removeItem(self,itemWin):
        if self.itemid_text.get() != "":
            if messagebox.askyesno("Item Setup","Remove item code?",parent=itemWin) :
                dell = db.delete_item(self.itemid_text.get())
                if dell == "success":
                    messagebox.showinfo("Sucess","Item code is deleted!",parent=itemWin)
                elif dell == "failed":
                    messagebox.showerror("Failed","Failed to delete item code!",parent=itemWin)
            else:
                messagebox.showinfo("Item Setup","Item is not removed!",parent=itemWin)
        else:
            messagebox.showerror("Error","No item to delete!",parent=itemWin)
        self.clearText()
        self.populateItems() 

    def updateItem(self,itemWin):
        if self.supplierid_text.get() == "" or self.itemcode_text.get() == "" or self.itemid_text.get() == "":
             messagebox.showerror("Required","Please fill in required field!",parent=itemWin) 
        else:
            if messagebox.askyesno("Item Update","Are you sure to edit item code?",parent=itemWin):
                update = db.update_item(self.itemid_text.get(),self.itemcode_text.get(),self.supplierid_text.get())
                if update == "success" :
                    messagebox.showinfo("Sucess","Item code is updated!",parent=itemWin)
                elif update == "failed":
                    messagebox.showerror("Failed","Failed to update item!",parent=itemWin)
            else:
                messagebox.showinfo("Item Update","No changes detected!",parent=itemWin)
        self.clearText()
        self.populateItems()


class SupplierSetup(Toplevel):

    def __init__(self):
        Toplevel.__init__(self)

        self.protocol("WM_DELETE_WINDOW", self.destroy)

        supIds = db.fetch_supplierids()

        w   = 520
        h   = 400 
        ws  = self.winfo_screenwidth()
        hs  = self.winfo_screenheight()
        x   = (ws/2) - (w/2)
        y   = (hs/2) - (h/2)
        self.geometry('%dx%d+%d+%d' % (w, h, x, y))
        self.resizable(0, 0)
        self.iconbitmap("parcel-delivery.ico")
        self.title('Supplier Setup')  

        self.supid_text   = StringVar()
        self.supname_text = StringVar()

        # Supplier ID
        supplierid_label = tk.Label(self, text='Supplier ID :', font=('bold', 10))
        supplierid_label.place(x=5,y=20)
        self.supplierid_cbo = ttk.Combobox(self, textvariable=self.supid_text, width=53, value=supIds)
        self.supplierid_cbo.bind('<<ComboboxSelected>>',self.selectSupplier) 
        self.supplierid_cbo.place(x=102, y=14)
        # Supplier Name
        suppliername_label = tk.Label(self, text='Supplier Name :', font=('bold', 10))
        suppliername_label.place(x=5,y=50)
        self.suppliername_entry = tk.Entry(self, textvariable=self.supname_text, relief='solid', width=57, highlightcolor= "black")
        self.suppliername_entry.place(x=105, y=50) 

        style = ttk.Style(self)
        style.map('Treeview',background=[('selected', "#217346")])
        style.configure('Treeview', rowheight=20)

        suplist_frame = tk.Frame(self,width=500,height=500)
        suplist_frame.pack(pady=90)
        suplist_frame.pack_propagate(0) 
        suplist_scroll = ttk.Scrollbar(suplist_frame)
        suplist_scroll.pack(side="right", fill="y")
        self.suplist_tree = ttk.Treeview(suplist_frame, yscrollcommand=suplist_scroll.set, selectmode="extended")
        self.suplist_tree.pack()
        suplist_scroll.config(command=self.suplist_tree.yview)
        self.suplist_tree['columns'] = ("Supplier ID","Supplier Name")
        # Format Our Columns
        self.suplist_tree.column("#0", width=0, stretch=NO)
        self.suplist_tree.column("Supplier ID", anchor=CENTER, minwidth=90, width=150,stretch=NO)
        self.suplist_tree.column("Supplier Name", anchor=W, minwidth=300, width=330,stretch=NO)
        # Create Headings
        self.suplist_tree.heading("#0", text="", anchor=W)
        self.suplist_tree.heading("Supplier ID", text="Supplier ID", anchor=CENTER)
        self.suplist_tree.heading("Supplier Name", text="Supplier Name", anchor=CENTER)
        # Create Striped Row Tags
        self.suplist_tree.tag_configure('oddrow', background="#4a4949")
        self.suplist_tree.tag_configure('evenrow', background="#1a8a4c")

        #Buttons
        add_btn = tk.Button(self, text='Add Supplier', width=13, relief=FLAT, command=lambda:self.addSup(self),bg="#217346")
        add_btn.place(x=15,y=320)
        remove_btn = tk.Button(self, text='Remove Supplier',  width=13, relief=FLAT, command=lambda:self.removeSup(self),bg="#217346")
        remove_btn.place(x=140,y=320)
        update_btn = tk.Button(self, text='Update Supplier', width=13, relief=FLAT, command=lambda:self.updateSup(self),bg="#217346")
        update_btn.place(x=265,y=320)
        clear_btn = tk.Button(self, text='Clear Input', width=13, relief=FLAT, command=self.clearText,bg="#217346")
        clear_btn.place(x=388,y=320)        
        #Bind the treeview
        self.suplist_tree.bind('<ButtonRelease-1>', self.selectedSup) 

        self.populateSuppliers()
        
    def selectSupplier(self,event):
        supname = db.fetch_supplier_by_id(self.supid_text.get())
        self.supname_text.set(supname[0])
        self.clearText()

    def populateSuppliers(self):
        for record in self.suplist_tree.get_children():
            self.suplist_tree.delete(record)
        count = 0
        suppliers = db.fetch_all_suppliers()
        
        for s in suppliers:
            if count % 2 == 0:
                self.suplist_tree.insert(parent='', index='end', iid=count, text='', values=(s[0], s[1]), tags=('evenrow',))
            else:
                self.suplist_tree.insert(parent='', index='end', iid=count, text='', values=(s[0], s[1]), tags=('oddrow',))
            count += 1

    def selectedSup(self,event):
        try:
            self.supplierid_cbo.delete(0,END)
            self.suppliername_entry.delete(0,END)
            # Grab record Number
            selected = self.suplist_tree.focus()
            # Grab record values
            values = self.suplist_tree.item(selected, 'values')
            # outpus to entry boxes
            self.supplierid_cbo.insert(0, values[0])
            self.suppliername_entry.insert(0, values[1])
        except IndexError :
            pass
    
    def clearText(self):
        self.supplierid_cbo.delete(0,END)
        self.suppliername_entry.delete(0,END)

    def updateSup(self,supWin):
        if self.supid_text.get() == "" or self.supname_text.get() == "" :
             messagebox.showerror("Required","Please fill in required field!",parent=supWin) 
        else:
            if messagebox.askyesno("Supplier Update","Are you sure to edit supplier?",parent=supWin):
                update = db.update_supplier(self.supid_text.get(),self.supname_text.get().upper())
                if update == "success" :
                    messagebox.showinfo("Sucess","Supplier is updated!",parent=supWin)
                elif update == "failed":
                    messagebox.showerror("Failed","Failed to update supplier!",parent=supWin)
            else:
                messagebox.showinfo("Supplier Update","No changes detected!",parent=supWin)
        self.clearText()

        self.populateSuppliers()

    def removeSup(self,supWin):
        if self.supid_text.get() != "":
            if messagebox.askyesno("Supplier Setup","Remove item code?",parent=supWin) :
                dell = db.delete_supplier(self.supid_text.get())
                if dell == "success":
                    messagebox.showinfo("Sucess","Supplier is deleted!",parent=supWin)
                elif dell == "failed":
                    messagebox.showerror("Failed","Failed to delete supplier!",parent=supWin)
            else:
                messagebox.showinfo("Item Setup","Supplier is not removed!",parent=supWin)
        else:
            messagebox.showerror("Error","No supplier to delete!",parent=supWin)
        self.clearText()
        self.populateSuppliers() 

    def addSup(self,supWin):
        if self.supid_text.get() == "" or self.supname_text.get() == "":
            messagebox.showerror("Required","Please fill in required field!",parent=supWin) 
        else :
            if messagebox.askyesno("Supplier Setup","Add new supplier?",parent=supWin) :
                insert = db.insert_supplier(self.supid_text.get(),self.supname_text.get().upper())
                if insert == "success":
                    messagebox.showinfo("Sucess","Supplier is added!",parent=supWin)
                elif insert == "duplicate":
                    messagebox.showerror("Duplicate","Supplier already exists!",parent=supWin)
            else:
                messagebox.showinfo("Info", "Supplier not added!",parent=supWin)

        self.clearText()
        self.populateSuppliers()


class SplitPage(Toplevel):
    def __init__(self):
        Toplevel.__init__(self)

        self.startpage_text = StringVar()
        self.endpage_text   = StringVar()
        self.protocol("WM_DELETE_WINDOW", self.destroy)
        self.selected_supplier_list = []
        w   = 520
        h   = 400 
        ws  = self.winfo_screenwidth()
        hs  = self.winfo_screenheight()
        x   = (ws/2) - (w/2)
        y   = (hs/2) - (h/2)
        self.geometry('%dx%d+%d+%d' % (w, h, x, y))
        self.resizable(0, 0)
        self.iconbitmap("pdf.ico")
        self.title('Split Page')     

        self.supplierlist = ["COSMETIQUE", #3
                           "ECOSSENTIAL", #4
                           "FOOD CHOICE", #22
                           "FOOD INDUSTRIES", #23
                           "GREEN CROSS", #24
                           "INTELIGENT", #5
                           "JS UNITRADE", #2
                           "MEAD JOHNSON", #8
                           "MONDELEZ", #9   
                           "SUYEN"] #10 
        

        self.selected_supplier_split = tk.StringVar()         
        self.supplier_split_label    = tk.Label(self, text='SELECT SUPPLIER', font=('bold', 12))
        self.supplier_split_label.pack(pady=50)
        self.supplier_split_cbo      = ttk.Combobox(self, textvariable=self.selected_supplier_split, state='readonly', width=40, value=self.supplierlist)
        self.supplier_split_cbo.place(x=100, y= 80)
       

        # Open File
        open_btn = tk.Button(self, text='Choose File', width=29, relief=FLAT, command=lambda:self.chooseSplitFile(self),bg="#0e74e8")
        open_btn.place(x=150,y=130)
        # PDF File Path
        self.pdffile_label = tk.Label(self, text='',  font=('bold', 9))
        self.pdffile_label.place(x=80,y=160)


        # Start Page   
        startpage_label = tk.Label(self, text='START PAGE :',  font=('bold', 10))
        startpage_label.place(x=120,y=200)
        self.startpage_entry = tk.Entry(self,textvariable=self.startpage_text , relief='solid', width=20, highlightcolor= "black")
        self.startpage_entry.place(x=220,y=200)
        # End Page   
        endpage_label = tk.Label(self, text='END PAGE :',  font=('bold', 10))
        endpage_label.place(x=120,y=230)
        self.endpage_entry = tk.Entry(self,textvariable=self.endpage_text , relief='solid', width=20, highlightcolor= "black")
        self.endpage_entry.place(x=220,y=230)

        # Split button
        split_btn = tk.Button(self, text='SPLIT PAGE', width=29, relief=FLAT, command=lambda:self.splitFile(self),bg="#217346")
        split_btn.place(x=150,y=280)
      

    def chooseSplitFile(self,splitWin):
        global pdf_split_file
        pdf_split_file = filedialog.askopenfilename(parent=splitWin,initialdir="shell:MyComputerFolder",title="Select PDF File", filetypes=(("PDF Files","*.pdf"),("All Files", "*.*")) )
        self.pdffile_label.config(text=pdf_split_file)

    def splitFile(self,splitWin):
        
        if self.startpage_text.get() == "" or  self.endpage_text.get() == "" or pdf_split_file == "":
            messagebox.showerror("Required","Please fill in required field!",parent=splitWin)

        else :
            from split_pdf import pdf_splitter
            split = pdf_splitter(self.startpage_text.get(),self.endpage_text.get(),pdf_split_file, self.selected_supplier_split.get())
            if split == 1:
                messagebox.showinfo("Sucess","File is split successfully!",parent=splitWin)
            else:
                messagebox.showerror("Duplicate","Failed to split file!",parent=splitWin)


class ConvertPSI(Toplevel):

    def __init__(self):
        Toplevel.__init__(self)

        self.protocol("WM_DELETE_WINDOW", self.destroy)

        w   = 520
        h   = 400 
        ws  = self.winfo_screenwidth()
        hs  = self.winfo_screenheight()
        x   = (ws/2) - (w/2)
        y   = (hs/2) - (h/2)
        self.geometry('%dx%d+%d+%d' % (w, h, x, y))
        self.resizable(0, 0)
        self.iconbitmap("convert-files.ico")
        self.title('Convert Pro-forma Sales Invoice')  

        self.win = self
        self.batch = ["First Batch", "Second Batch"]
        self.firstBatch = ["COSMETIQUE", #3
                           "ECOSSENTIAL", #4
                           "FOOD CHOICE", #22
                           "FOOD INDUSTRIES", #23
                           "GREEN CROSS", #24
                           "INTELIGENT", #5
                           "JS UNITRADE", #2
                           "MEAD JOHNSON", #8
                           "MONDELEZ", #9   
                           "SUYEN"] #10 

        self.secondBatch = ["ALASKA","BIG E","GSMI","KSK","RECKIT"]   


        #CONVERT WIDGETS    
        self.selected_supplier = tk.StringVar() 
        self.selected_batch    = tk.StringVar()
        self.sheetnumber       = tk.StringVar()
        self.supplier_label    = tk.Label(self, text='SELECT SUPPLIER', font=('bold', 14))
        self.supplier_label.pack(pady=50)
        self.batch_cbo         = ttk.Combobox(self, textvariable=self.selected_batch, state='readonly', width=15, value=self.batch)
        self.batch_cbo.bind('<<ComboboxSelected>>',self.selectBatch) 
        self.batch_cbo.place(x=70, y= 120)
        self.supplier_cbo      = ttk.Combobox(self, textvariable=self.selected_supplier, state='readonly', width=30, value=SUPPLIER_LIST)
        self.supplier_cbo.place(x=220, y= 120)
        self.supplier_cbo.bind('<<ComboboxSelected>>', lambda event:self.openFile(self,event)) 
        # self.supplier_cbo.pack()
        self.note_label    = tk.Label(self, text='NOTE :', font=('bold', 12))
        self.note_label.place(x=70, y=320)
        self.note_text_label  = tk.Label(self, text='Please provide item mapping in the CUSTOMER MATERIAL column.', font=('bold', 10))
        self.note_text_label.place(x=70, y=340)
        # self.note_label.pack(pady=50)
        self.progress = ttk.Progressbar(self, style="green.Horizontal.TProgressbar", orient= HORIZONTAL, length = 300,mode = 'indeterminate',takefocus=True, maximum=100)         
        self.progress_label = tk.Label(self, text='', font=('', 8))
        self.progress_label.place(x=110,y=200)

    def openFile(self,convertWin,event):
        global file_path
        if self.selected_supplier.get() == "MONDELEZ" :
            file_path = filedialog.askopenfilename(parent=convertWin,initialdir="shell:MyComputerFolder",title="Select Excel File", filetypes=(("Excel Files","*.xls*"),("All Files", "*.*")) )
        elif self.selected_supplier.get() == "BIG E" or self.selected_supplier.get() == "FOOD CHOICE" or self.selected_supplier.get() == "GSMI" or self.selected_supplier.get() == "GREEN CROSS" or \
            self.selected_supplier.get() == "ALASKA" :           
            file_path = filedialog.askopenfilename(parent=convertWin,initialdir="shell:MyComputerFolder",title="Select Excel File", filetypes=(("Excel Files","*.xlsx"),("All Files", "*.*")) )
        elif self.selected_supplier.get() == "FOOD INDUSTRIES" :
            file_path = filedialog.askopenfilename(parent=convertWin,initialdir="shell:MyComputerFolder",title="Select Excel File", filetypes=(("Excel Files","*.xls"),("All Files", "*.*")) )
        elif  self.selected_supplier.get() == "MEAD JOHNSON" or self.selected_supplier.get() == "RECKIT":
            file_path = filedialog.askopenfilename(parent=convertWin,initialdir="shell:MyComputerFolder",title="Select Textfile File", filetypes=(("Text Documnents","*.txt"),("All Files", "*.*")) )

        else:
            file_path = filedialog.askopenfilename(parent=convertWin,initialdir="shell:MyComputerFolder",title="Select PDF File", filetypes=(("PDF Files","*.pdf"),("All Files", "*.*")) )

        self.startThread()   

    def selectBatch(self,event): 
        self.supplier_cbo.set("")
        if self.selected_batch.get() == "First Batch":
            SUPPLIER_LIST = self.firstBatch
        elif self.selected_batch.get() == "Second Batch":
            SUPPLIER_LIST = self.secondBatch

        self.supplier_cbo['values'] =  SUPPLIER_LIST

    def startThread(self):    
        global submit_thread    
        submit_thread = threading.Thread(target=self.processFile)
        submit_thread.daemon = True
        submit_thread.start()        
        self.runProgress()
        self.after(10, self.checkThread)

    def processFile(self):
        global getResult
        result = self.quitExcel()
        if result :
            self.progress_label.config(text='Converting ....')
            getResult = self.validateFile(file_path)              
    
    def runProgress(self):  
        self.progress.pack(pady=50)
        self.progress.start()

    def quitExcel(self):
        import win32com.client
        proceed = False
        self.progress_label.config(text='Checking open Excel instance ...')
        try: 
            excel = win32com.client.GetActiveObject("Excel.Application")
            
        except:
            excel = ""  
        
        if excel != "":
            import psutil

            self.progress_label.config(text='Closing Excel instance ...')
            for proc in psutil.process_iter():
                if proc.name().lower() == "excel.exe":
                    proc.kill()
            proceed = True
        else :
            proceed = True

        self.progress_label.config(text='')
        return proceed


    def checkThread(self):
        if submit_thread.is_alive():
            self.progress.step()
            self.after(10, self.checkThread)
        else:
            self.progress.stop()
            self.progress.pack_forget()  

            self.progress_label.config(text='')
            if getResult == 1:               
                messagebox.showinfo('Success', 'File is converted successfully!',parent=self.win)
            elif getResult == 2:
                messagebox.showerror('Error', 'File is not a valid PSI!',parent=self.win)
            elif getResult == 3:
                messagebox.showerror('Error', 'Invalid file extension!',parent=self.win)
            elif getResult == 4:
                messagebox.showerror('Required', 'No file selected!',parent=self.win)
            else:#0
                messagebox.showerror('Error', 'Converted file not found!',parent=self.win) 

    def validateFile(self,file_path):
        file_ext = ""
        if self.selected_supplier.get() == "MONDELEZ" :
            file_ext = ['.xlsb','xlsx']
        elif self.selected_supplier.get() == "BIG E" or self.selected_supplier.get() == "FOOD CHOICE" or self.selected_supplier.get() == "GSMI" or self.selected_supplier.get() == "GREEN CROSS" or \
            self.selected_supplier.get() == "ALASKA" :
            file_ext = ['.xlsx', '.XLSX']
        elif self.selected_supplier.get() == "FOOD INDUSTRIES" : 
            file_ext = ['.xls']
        elif self.selected_supplier.get() == "MEAD JOHNSON" or self.selected_supplier.get() == "RECKIT" : 
            file_ext = ['.txt', '.TXT']
        else:
            file_ext = ['.pdf','.PDF']

        if file_path == "":
                result = 4                
        else:
            
            if file_path.endswith(tuple(file_ext)):     
                result = self.readPdfBySupplier(self.selected_supplier.get(),file_path)    
            else :                
                result = 3          

            self.supplier_cbo.set("")
            file_path = ""
            return result 

        
    def readPdfBySupplier(self,supplier,file_path):
        xlsx = 0        
        if supplier == "JS UNITRADE":
            from read_pdf_js import read_pdf
            xlsx = read_pdf(file_path) 
                  
        elif supplier == "SUYEN":
            import read_pdf_suyen as suyen #encrypted pdf+
            decrypted_file = suyen.decrypt_pdf(file_path)
            path = Path(decrypted_file)
            if path.is_file():
                xlsx = suyen.read_pdf(decrypted_file)
            else:
                messagebox.showerror('Error', decrypted_file)

        elif supplier == "MONDELEZ":
            import read_mondelez as mon
            ext = Path(file_path).suffix
            if ext == ".xlsb":
                converted_file = mon.convert_to_xlsx(file_path) #.xlsb file
                path = Path(converted_file)
                if path.is_file():
                    xlsx = mon.read_mondelez(converted_file,1)
                else:
                    messagebox.showerror('Error', converted_file)
            elif ext == ".xlsx" :
                xlsx = mon.read_mondelez(file_path,2)

        elif supplier == "INTELIGENT":         
            from read_pdf_intelligent import read_pdf 
            xlsx = read_pdf(file_path)  

        elif supplier == "COSMETIQUE":
            from read_pdf_cosmetique import read_pdf    
            xlsx = read_pdf(file_path)

        elif supplier == "ECOSSENTIAL":
            from read_pdf_eco import read_pdf
            xlsx = read_pdf(file_path)
        
        elif supplier == "FOOD CHOICE":
            from read_food_choice  import read_xlsx
            xlsx = read_xlsx(file_path)

        elif supplier == "FOOD INDUSTRIES":
            import read_food_industries as fi
            converted_file = fi.convert_to_xlsx(file_path) #.xls
            path = Path(converted_file)
            if path.is_file():
                xlsx = fi.read_foodin(converted_file)
            else:
                messagebox.showerror('Error', converted_file)       

        elif supplier == "GREEN CROSS":
            from read_green_cross import read_xlsx
            xlsx = read_xlsx(file_path)
        
        elif supplier == "MEAD JOHNSON":
            from read_mead_johnson import textfile_to_xlsx
            xlsx = textfile_to_xlsx(file_path)
            

        ## START SECOND BATCH
        elif supplier == "KSK":
            from read_ksk import read_pdf
            xlsx = read_pdf(file_path)

        elif supplier == "BIG E":
            from read_big_e import read_xlsx
            xlsx = read_xlsx(file_path)

        elif supplier == "ALASKA":
            from read_alaska import read_xlsx
            xlsx = read_xlsx(file_path)

        elif supplier == "GSMI":
            from read_gsmi import read_xlsx
            xlsx = read_xlsx(file_path)

        elif supplier == "RECKIT":
            from read_reckit import textfile_to_xlsx
            xlsx = textfile_to_xlsx(file_path)
            
        ## END SECOND BATCH
        
        elif supplier == "ALECO":   
            # from read_pdf_aleco import read_pdf
            # xlsx = read_pdf(file_path)
            pass

       


        
            

        else:
            messagebox.showerror('Error', 'Unknown supplier!')

        return xlsx

    def inputSheetNumber(self):
        w   = 350
        h   = 120 
        ws  = self.winfo_screenwidth()
        hs  = self.winfo_screenheight()
        x   = (ws/2) - (w/2)
        y   = (hs/2) - (h/2)
        my_w_child=Toplevel(self.win) 
        my_w_child.geometry('%dx%d+%d+%d' % (w, h, x, y))
        my_w_child.resizable(0, 0)
        my_w_child.iconbitmap("convert-files.ico")
        my_w_child.title('Convert Pro-forma Sales Invoice')
        
        # Open File
        open_btn = tk.Button(self, text='Choose File', width=29, relief=FLAT, command=lambda:self.chooseSplitFile(self),bg="#0e74e8")
        open_btn.place(x=150,y=130)
        # PDF File Path
        self.pdffile_label = tk.Label(self, text='',  font=('bold', 9))
        self.pdffile_label.place(x=80,y=160)

        #Input Sheet Number       
        input_label = tk.Label(my_w_child, text='Input Sheet Number to Convert :',  font=('bold', 12))
        input_label.place(x=10,y=15)
        input_sheetno_entry = tk.Entry(my_w_child,textvariable=self.sheetnumber, relief='solid', width=47, highlightcolor= "white")
        input_sheetno_entry.place(x=10,y=40)

        

window = tk.Tk()

app = Application(master=window)
app.mainloop()