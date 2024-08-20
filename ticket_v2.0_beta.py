from tkinter import *
from tkinter import StringVar
import tkinter.messagebox as msgbox
import tkinter.ttk
import tkinter.font
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment
import pandas as pd
from pandastable import Table, TableModel
import re
from datetime import datetime
import json

# Define the font properties
font = Font(name='굴림체', size=9, bold=False)
nm_font = Font(name='맑은 고딕', size=11, bold=False)

with open('user_data_int.json', 'r', encoding='utf-8') as file:
    user_data_int = json.load(file)

class TicketNumber():
    def __init__(self, ticket_number):
        self._ticket_number = ticket_number

    @property
    def ticket_number(self):
        return self._ticket_number

    @ticket_number.setter
    def ticket_number(self, value):
        self._ticket_number = value
        ticket_var.set(self._ticket_number-1)  # Update the StringVar

class NonMember():
    def __init__(self, non_member_row):
        self.non_member_row = non_member_row

count = TicketNumber(1)
nm_row = NonMember(2)

def excel_loader(mont, day):
    # Retrieve month and date
    date = f'{day:02}'
    month = f'{mont:02}'

    # Set filename
    file_name = f'전체이용자(편집)_{month}_{date}'

    try:
        # Load the workbook and sheet
        workbook = load_workbook(f'{file_name}.xlsx')
        sheet = workbook['회원']
        print("Successfully Loaded Excel")
        return workbook, sheet, file_name
    except FileNotFoundError:
        msgbox.showerror("Error", "Excel file not found.")
        return None, None, None
    
def jsonLoader(mont, day):
    # Retrieve month and date
    date = f'{day:02}'
    month = f'{mont:02}'

    # Set filename
    file_name = f'식권_판매정보_{month}_{date}.json'

    try:
        with open(file_name, 'r', encoding='utf-8') as file:
            data = json.load(file)
            if len(data) != 0:
                for d in data:
                    treeview.insert('', '0', text='', values=(f'{d['user_id']} {d['user_name']} {d['status']} {d['price']} {d['ticket_number']}'))
            return file_name
    except:
        with open(file_name, 'w', encoding='utf-8') as file:
            json.dump([], file)
        return file_name


def warn():
    msgbox.showwarning("경고", "회원번호 불일치")

def output(text):
    msgbox.showinfo("Output", f'{text}')

tk = Tk()
tk.title("식권 판매 프로그램")
tk.geometry("800x800")

# Create the DataFrame viewer at the top
frame = Frame(tk)
frame.pack(fill=BOTH, expand=True)

# Define the font for the Treeview
treeview_font = tkinter.font.Font(family="Arial", size=13)

style = tkinter.ttk.Style()
style.configure("Treeview", font=treeview_font)  # Set the font for the items
style.configure("Treeview.Heading", font=treeview_font)  # Set the font for the headings

treeview = tkinter.ttk.Treeview(
    frame,            
    columns=['UserID', 'Name', 'Status', 'price', 'ticketNumber'],
    displaycolumns=['UserID', 'Name', 'Status', 'price', 'ticketNumber'], show='headings')
treeview.pack(side=TOP, fill=BOTH, expand=True, padx=10, pady=10)

treeview.column("UserID", width=100, anchor="center")
treeview.heading("UserID", text="회원번호", anchor="center")

treeview.column("Name", width=100, anchor="center")
treeview.heading("Name", text="이름", anchor="center")

treeview.column("Status", width=150, anchor="center")
treeview.heading("Status", text="생활구분", anchor="center")

treeview.column("price", width=100, anchor="center")
treeview.heading("price", text="가격", anchor="center")

treeview.column("ticketNumber", width=100, anchor="center")
treeview.heading("ticketNumber", text="식권", anchor="center")

# Create a frame to hold the entry and buttons horizontally
input_frame = Frame(tk)
input_frame.pack(pady=10)

entry = Entry(input_frame)
entry.grid(row=0, column=0, padx=5)

jsonName = jsonLoader(datetime.now().month, datetime.now().day)

def search(event=None): # Add event parameter to handle binding
    id = entry.get()
    jsonSellTicket(id)
    entry.delete(0, 'end')

'''def refund():
    

def add_user(userID, name, status):
    

def add_non_member(name, reason_for_visit):'''

def jsonSellTicket(userID):
    global jsonName
    data = user_data_int.get(userID)
    name = data[0]
    status = data[1]
    price = data[2]
    
    new_data = {
            "user_id": userID,
            "user_name": name,
            "status": status,
            "price": price,
            "ticket_number": '1' 
        }
    
    treeview.insert('', '0', text='', values=(f'{userID} {name} {status} {price} {count.ticket_number}'))
    
    # Open the JSON file and load the data
    with open(jsonName, 'r', encoding='utf-8') as file:
        data = json.load(file)

    # Append new data to the list
    data.append(new_data)

    # Write the updated data back to the file
    with open(jsonName, 'w', encoding='utf-8') as file:
        json.dump(data, file, indent=4, ensure_ascii=False)

def jsonUpdate(userID, name, status, price):
    global jsonName
    new_data = {
            "user_id": userID,
            "user_name": name,
            "status": status,
            "price": price,
            "ticket_number": count.ticket_number 
        }

    treeview.insert('', '0', text='', values=(f'{userID} {name} {status} {price} {count.ticket_number}'))
    
    # Open the JSON file and load the data
    with open(jsonName, 'r', encoding='utf-8') as file:
        data = json.load(file)

    # Append new data to the list
    data.append(new_data)

    # Write the updated data back to the file
    with open(jsonName, 'w', encoding='utf-8') as file:
        json.dump(data, file, indent=4, ensure_ascii=False)

def reset_ticket_number():
    reset_value = int(entry.get())
    count.ticket_number = reset_value

'''def save_file():
    dir = f'C:\\Users\\Administrator\\Desktop\\식권2024\\8월\\{file_name}.xlsx'
    workbook.save(dir)
    print("FILE SAVED SUCCESSFULLY!!!")'''

def addUserPopUp():
    addUserWindow = Toplevel(tk)
    addUserWindow.geometry('500x300')
    addUserWindow.title('신규회원추가')
    userID_text = Entry(addUserWindow)
    name_text = Entry(addUserWindow)
    status_text = Entry(addUserWindow)
    
    Label(addUserWindow, text='회원번호').grid(row=0, column=0, padx=5)
    userID_text.grid(row=0, column=1, padx=5)
    Label(addUserWindow, text='이름').grid(row=1, column=0, padx=5)
    name_text.grid(row=1, column=1, padx=5)
    Label(addUserWindow, text='생활구분').grid(row=2, column=0, padx=5)
    status_text.grid(row=2, column=1, padx=5)

    addUserButton = Button(addUserWindow, text='신규회원추가', command=lambda: add_user(userID_text.get(), name_text.get(), status_text.get()))
    addUserButton.grid(row=3, column=2, padx=5)

def nonUserPopUp():
    nonUserWindow = Toplevel(tk)
    nonUserWindow.geometry('500x300')
    nonUserWindow.title('비회원 판매')

    name_text = Entry(nonUserWindow)
    reason_for_visit_text = Entry(nonUserWindow)

    Label(nonUserWindow, text='이름').grid(row=0, column=0, padx=5)
    name_text.grid(row=0, column=1, padx=5)
    Label(nonUserWindow, text='방문용무').grid(row=1, column=0, padx=5)
    reason_for_visit_text.grid(row=1, column=1, padx=5)

    addNonUserButton = Button(nonUserWindow, text='비회원 판매', command=lambda: add_non_member(name_text.get(), reason_for_visit_text.get()))
    addNonUserButton.grid(row=2, column=2, padx=5)

# Create the StringVar after initializing the root window
ticket_var = StringVar()

search_button = Button(input_frame, text='검색', command=search)
search_button.grid(row=0, column=1, padx=5)

total_tickets_sold = Label(input_frame, textvariable=ticket_var, font=('Helvetica', 25, 'bold'))
total_tickets_sold.grid(row=1)

# Bind the Enter key to the search function
entry.bind("<Return>", search)

tk.mainloop()