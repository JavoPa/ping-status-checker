import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox
import subprocess
import threading
import os

def ping(ip):
    try:
        output = subprocess.check_output(["ping", "-n", "1", ip], stderr=subprocess.STDOUT, universal_newlines=True)
        if "TTL=" in output:
            return True
        else:
            return False
    except subprocess.CalledProcessError:
        return False

def update_status(tree, ip, item):
    if ping(ip):
        tree.item(item, tags=('online',))
    else:
        tree.item(item, tags=('offline',))

def load_data():
    try:
        df = pd.read_excel('ips.xlsx')
    except FileNotFoundError:
        messagebox.showerror("Error", "El archivo 'ips.xlsx' no existe.")
        return

    for index, row in df.iterrows():
        item = tree.insert('', 'end', text=row['IP'], values=(row['Description'], 'Checking...'))
        threading.Thread(target=update_status, args=(tree, row['IP'], item)).start()

root = tk.Tk()
root.title("IP Status Checker")

tree = ttk.Treeview(root, columns=('Description', 'Status'), show='headings')
tree.heading('Description', text='Description')
tree.heading('Status', text='Status')
tree.pack(fill=tk.BOTH, expand=True)

tree.tag_configure('online', background='green')
tree.tag_configure('offline', background='red')

load_data()

root.mainloop()