from docxtpl import DocxTemplate
import tkinter
from tkinter import ttk
from docxtpl import DocxTemplate
import datetime
from tkinter import messagebox
pip install docxtpl

doc = DocxTemplate("invoice_template.docx")


invoice_list = [["", "", "", ""],
                ["", "", "", ""]]
                


doc.render({"name":"name", 
            "phone":"phone",
            "invoice_list": invoice_list,
            "subtotal":0,
            "Biltycharges":0,
            "total":0})
doc.save("new_invoice.docx")
def clear_item():
    qty_spinbox.delete(0, tkinter.END)
    qty_spinbox.insert(0, "1")
    desc_entry.delete(0, tkinter.END)
    price_spinbox.delete(0, tkinter.END)
    price_spinbox.insert(0, "0.0")
    Wei_spinbox.delete(0,tkinter.END)
    Wei_spinbox.insert(0,"0.0")
    bil_spinbox.delete(0,tkinter.END)
    bil_spinbox.insert(0,"0.0")
    

invoice_list = []
def add_item():
    qty = float(qty_spinbox.get())
    desc = desc_entry.get()
    price = float(price_spinbox.get())
    weight= float(Wei_spinbox.get())
    biltycharges = float(bil_spinbox.get())
    total = (weight * price) + biltycharges 
    invoice_item = [qty, desc, price,weight,biltycharges,total]
    tree.insert('',0, values=invoice_item)
    clear_item()
    
    invoice_list.append(invoice_item)

    
def new_invoice():
    first_name_entry.delete(0, tkinter.END)
    Address_name_entry.delete(0, tkinter.END)
    phone_entry.delete(0, tkinter.END)
    clear_item()
    tree.delete(*tree.get_children())
    
    invoice_list.clear()
    
def generate_invoice():
    doc = DocxTemplate("invoice_template.docx")
    name = first_name_entry.get()
    address = Address_name_entry.get()
    phone = phone_entry.get()
    subtotal = subtotal = sum(item[3] * item[2] for item in invoice_list) 
    biltycharges = float(bil_spinbox.get())
    total = subtotal + biltycharges
    
    doc.render({"name":name, 
            "phone":phone,
            "invoice_list": invoice_list,
            "subtotal":subtotal,
            "biltycharges":biltycharges,
            "total":total})
    
    doc_name = "new_invoice" + name + datetime.datetime.now().strftime("%Y-%m-%d-%H%M%S") + ".docx"
    doc.save(doc_name)
    
    messagebox.showinfo("Invoice Complete", "Invoice Complete")
    
    new_invoice()


    

window = tkinter.Tk()
window.title("Fast Express delivery")

frame = tkinter.Frame(window)
frame.pack(padx=10, pady=10)

first_name_label = tkinter.Label(frame, text="First Name")
first_name_label.grid(row=0, column=0)
Address_name_label = tkinter.Label(frame, text="Address")
Address_name_label.grid(row=0, column=1)

first_name_entry = tkinter.Entry(frame)
Address_name_entry = tkinter.Entry(frame)
first_name_entry.grid(row=1, column=0)
Address_name_entry.grid(row=1, column=1)

phone_label = tkinter.Label(frame, text="Phone")
phone_label.grid(row=0, column=2)
phone_entry = tkinter.Entry(frame)
phone_entry.grid(row=1, column=2)

qty_label = tkinter.Label(frame, text="Qty")
qty_label.grid(row=2, column=0)
qty_spinbox = tkinter.Spinbox(frame, from_=1, to=100)
qty_spinbox.grid(row=3, column=0)

desc_label = tkinter.Label(frame, text="Description")
desc_label.grid(row=2, column=1)
desc_entry = tkinter.Entry(frame)
desc_entry.grid(row=3, column=1)

Wei_label = tkinter.Label(frame, text="Weight")
Wei_label.grid(row=0, column=4)
Wei_spinbox = tkinter.Spinbox(frame, from_=0, to=1000)
Wei_spinbox.grid(row=1, column=4)

bil_label = tkinter.Label(frame, text="Bilty")
bil_label.grid(row=2, column=4)
bil_spinbox = tkinter.Spinbox(frame, from_=0, to=1000)
bil_spinbox.grid(row=3, column=4)

price_label = tkinter.Label(frame, text="Price")
price_label.grid(row=2, column=2)
price_spinbox = tkinter.Spinbox(frame, from_=0, to=1000)
price_spinbox.grid(row=3, column=2)

add_item_button = tkinter.Button(frame, text = "Add item", command = add_item)
add_item_button.grid(row=4, column=2, pady=5)

columns = ('qty', 'desc', 'price', 'weight','BiltyCharges','total' )
tree = ttk.Treeview(frame, columns=columns, show="headings")
tree.heading('qty', text='Qty')
tree.heading('desc', text='Description')
tree.heading('price', text='Price')
tree.heading('weight', text='Weight')
tree.heading('BiltyCharges', text="BiltyCharges")
tree.heading('total', text="Total")

    
tree.grid(row=5, column=0, columnspan=3, padx=20, pady=10)


save_invoice_button = tkinter.Button(frame, text="Generate Invoice", command=generate_invoice)
save_invoice_button.grid(row=6, column=0, columnspan=3, sticky="news", padx=20, pady=5)
new_invoice_button = tkinter.Button(frame, text="New Invoice", command=new_invoice)
new_invoice_button.grid(row=7, column=0, columnspan=3, sticky="news", padx=20, pady=5)


window.mainloop()


