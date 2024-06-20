import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
from sqlalchemy import create_engine
import pyodbc


def Get_Database(server):
    Driver = "DRIVER={ODBC Driver 17 for SQL Server};"
    Server = f"SERVER=SF3-AP-PH\\{server};"
    Database = "DATABASE=master;" 
    connectionString = pyodbc.connect(Driver + Server + Database + "Trusted_Connection=yes;")
    get_query = "SELECT name FROM sys.databases WHERE database_id > 4"
    command = connectionString.cursor()
    command.execute(get_query)
    db_rows = command.fetchall()
    databases = [row[0] for row in db_rows if row[0]]  
    return databases

def update_db_combobox_state(event):
    if not server_combobox.get():  
        db_combobox.config(state="disabled")
        textbox_query.config(state="disabled")  # Disable the textbox
    else:
        db_combobox.config(state="readonly")
        textbox_query.config(state="normal")  # Enable the textbox
        server = server_combobox.get()
        databases = Get_Database(server)
        db_combobox.set('')
        db_combobox['values'] = databases

def button_clicked():
    try:
        root = tk.Tk()
        root.withdraw() 
        file_path = filedialog.asksaveasfilename(defaultextension=".csv",
                                                  filetypes=[("CSV files", "*.csv")],
                                                  title="Save CSV File As")
        if not file_path: 
            return
        
        connection_string = f"mssql+pyodbc://SF3-AP-PH\{server_combobox.get()}/{db_combobox.get()}?driver=ODBC+Driver+17+for+SQL+Server"

        engine = create_engine(connection_string)

        get_pallet =  textbox_query.get("1.0", "end-1c")
        df = pd.read_sql(sql=get_pallet, con=engine)

        df.to_csv(file_path, index=False)
        messagebox.showinfo("Extraction Completed", f"CSV file saved at:\n{file_path}")
        server_combobox.set('')
        db_combobox.set('')
        textbox_query.set('')
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")

servers = ["APSQLSERVER01", "APSQLSERVER02"]
root = tk.Tk()
root.title("SQL TO CSV")
root.geometry('500x500')
root.eval('tk::PlaceWindow . center')

label_server = tk.Label(root, text="Select SQL Server : ")
label_server.place(x=1, y=10)
server_combobox = ttk.Combobox(root, values=servers)
server_combobox.place(x=110, y=10)
server_combobox.bind("<<ComboboxSelected>>", update_db_combobox_state)  

label_db = tk.Label(root, text="Select Database : ")
label_db.place(x=1, y=50)
db_combobox = ttk.Combobox(root, state="disable")
db_combobox.place(x=110, y=50)

label_query = tk.Label(root, text="Paste the query : ")
label_query.place(x=1, y=90)
textbox_query = tk.Text(root, wrap="word", state="disabled")  # Initially disabled
textbox_query.place(x=110, y=90, width=350, height=350)

btn_extract = tk.Button(root, text="EXTRACT", command=button_clicked, bg="green", fg="white", height=1, width=10, font=("Calibri", 12, "bold"))
btn_extract.place(x=400, y=460 )

root.mainloop()
