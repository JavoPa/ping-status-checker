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
        tree.set(item, column='Status', value='Online')
    else:
        tree.item(item, tags=('offline',))
        tree.set(item, column='Status', value='Offline')

def load_data():
    try:
        df = pd.read_excel('ips.xlsx')
    except FileNotFoundError:
        messagebox.showerror("Error", "El archivo 'ips.xlsx' no existe.")
        return

    for index, row in df.iterrows():
        item = tree.insert('', 'end', values=(row['IP'], row['Description'], 'Checking...'))
        update_status_periodically(tree, row['IP'], item)

def update_status_periodically(tree, ip, item):
    threading.Thread(target=lambda: update_status(tree, ip, item)).start()
    tree.after(5000, update_status_periodically, tree, ip, item) # 5000 millisegundos = 5 segundos

root = tk.Tk()
root.title("IP Status Checker")

tree = ttk.Treeview(root, columns=('IP', 'Description', 'Status'), show='headings')
tree.heading('IP', text='IP')
tree.heading('Description', text='Descripción')
tree.heading('Status', text='Estado')
tree.pack(fill=tk.BOTH, expand=True)

tree.tag_configure('online', background='green')
tree.tag_configure('offline', background='red')

load_data()

root.mainloop()