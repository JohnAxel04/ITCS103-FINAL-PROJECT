import tkinter as tk
import openpyxl as op
from tkinter import ttk,messagebox
import datetime
from datetime import datetime
def valid():
    entry = productEntry.get()
    qty = productquatity.get()
    price = productPrice.get()
    exp = productExpiry.get()

    if not entry or not qty or not price or not exp:
        messagebox.showerror("Invalid Input","Input must not be empty")
        return False
    if not qty.isdigit() or not price.isdigit():
        messagebox.showerror("Invalid Input","Quantity,Price,Expiry Date must be a number")
        return False
    
    try:
        expiry_date = datetime.strptime(exp, "%Y-%m-%d",).date()
        today = datetime.today().date()

        if expiry_date < today:
            messagebox.showerror("Expired","Item already expired")
            return False
        
    except:
        messagebox.showerror("Error","User format YYYY-MM-DD")
        return False

    return True

def save():
    if not valid():
        return
    entry = productEntry.get()
    qty = int(productquatity.get())
    price = int(productPrice.get())
    exp = productExpiry.get()
    opt = typeOption.get()
    wrk = op.load_workbook("Inventory.xlsx")
    sheet = wrk.active
    totalprice = qty * price
    id = sheet.max_row
    sheet.append([id,entry,qty,price,exp,opt,totalprice])
    wrk.save("Inventory.xlsx")
    messagebox.showinfo("Successfull","Item added successfully")
    refresh()
def refresh():
    work = op.load_workbook("Inventory.xlsx")
    sheet = work.active
    for i in table.get_children():
        table.delete(i)
    for row in sheet.iter_rows(min_row=2,values_only=True):
        table.insert("",tk.END,values=row)
def focus(event):
    selected = table.focus()
    values = table.item(selected,"values")
    
    if values:
        productEntry.delete(0,tk.END)
        productquatity.delete(0,tk.END)
        productPrice.delete(0,tk.END)
        productExpiry.delete(0,tk.END)
        typeOption.current(0)

        typeopt = values[5]
        productEntry.insert(0,values[1])
        productquatity.insert(0,values[2])
        productPrice.insert(0,values[3])
        productExpiry.insert(0,values[4])
        typeOption.set(typeopt)

def update():
    select = table.focus()
    values = table.item(select,"values")

    if not select:
        messagebox.showerror("error","Select to update")
        return
    if not valid():
        return
    
    id = values[0]
    entry = productEntry.get()
    qty = int(productquatity.get())
    price = int(productPrice.get())
    exp = productExpiry.get()
    opt = typeOption.get()
    total = qty * price

    work = op.load_workbook("Inventory.xlsx")
    sheet = work.active
    for i in sheet.iter_rows(min_row=2):
        if str(i[0].value) == str(id):
            i[1].value = entry
            i[2].value = qty
            i[3].value = price
            i[4].value = exp
            i[5].value = opt
            i[6].value = total
    work.save("Inventory.xlsx")
    messagebox.showinfo("Success","Info successfully updated")
    refresh()
def delete():
    select = table.focus()
    values = table.item(select,"values")

    if not select:
        messagebox.showerror("error","Select to update")
        return
    
    confirm = messagebox.askyesnocancel("Delete","Want to delete?")
    if not confirm:
        return
    id = values[0]
    work = op.load_workbook("Inventory.xlsx")
    sheet = work.active
    for i,row in enumerate(sheet.iter_rows(min_row=2),start=2):
        if str(row[0].value) == str(id):
            sheet.delete_rows(i)

    productEntry.delete(0,tk.END)
    productquatity.delete(0,tk.END)
    productPrice.delete(0,tk.END)
    productExpiry.delete(0,tk.END)
    typeOption.current(0)
    work.save("Inventory.xlsx")
    messagebox.showinfo("Delete","Info successfully deleted")
    refresh()


window = tk.Tk()
window.title("Inventory System")
mainlabel = tk.Label(window,text="Inventory System",font=("Poppins",15,"bold"))
mainlabel.grid(columnspan=5)
frame = tk.Frame(window,background="#888")
frame.grid(row=1,columnspan=6,padx=5,pady=8)

productEntry = tk.Entry(frame)
productEntry.grid(row=0,column=0,padx=5,pady=(8,2))
productquatity = tk.Entry(frame)
productquatity.grid(row=0,column=1,padx=5,pady=(8,2))
productPrice = tk.Entry(frame)
productPrice.grid(row=0,column=2,padx=5,pady=(8,2))
productExpiry = tk.Entry(frame)
productExpiry.grid(row=0,column=3,padx=5,pady=(8,2))
options = ['Box','Pack','Sachet',"Galon"]
typeOption = ttk.Combobox(frame,values=options,state="readonly")
typeOption.current(0)
typeOption.grid(column=4,row=0,padx=5,pady=(8,2))

NameLabel = tk.Label(frame,text="Product Name",background="#888",foreground="white")
NameLabel.grid(row=1,column=0)
QuantityLabel = tk.Label(frame,text="Quantity",background="#888",foreground="white") 
QuantityLabel.grid(row=1,column=1) 
PriceLabel = tk.Label(frame,text="Price",background="#888",foreground="white")
PriceLabel.grid(row=1,column=2)
ExpiryLabel = tk.Label(frame,text="Expiry Date",background="#888",foreground="white")
ExpiryLabel.grid(row=1,column=3)
TypeLabel = tk.Label(frame,text="Type of Packaging",background="#888",foreground="white")
TypeLabel.grid(row=1,column=4)
def savebt(event):
    SaveBtn['background'] = "black"
btnframe = tk.Frame(window)
btnframe.grid(row=2,columnspan=4)
SaveBtn = tk.Button(btnframe,text="Submit",command=save)
SaveBtn.grid(column=1,row=0,padx=10,pady=(1,10))
SaveBtn.bind("<<Enter>>",savebt)
UpdateBtn = tk.Button(btnframe,text="Update",command=update)
UpdateBtn.grid(column=0,row=0,padx=10,pady=(1,10))
DeleteBtn = tk.Button(btnframe,text="Delete",command=delete)
DeleteBtn.grid(column=2,row=0,padx=10,pady=(1,10))

table = ttk.Treeview(window,columns=("ID","Product Name","Quantity","Price","Expiry","Type of packaging","Total Price"),show="headings")
for row in ("ID","Product Name","Quantity","Price","Expiry","Type of packaging","Total Price"):
    table.heading(row, text=row)
table.grid(columnspan=4,row=3,padx=10)
table.bind("<<TreeviewSelect>>",focus)
refresh()
window.mainloop()