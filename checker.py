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
    if item in tree.get_children():
        if ping(ip):
            tree.after(0, lambda: update_tree_item(tree, item, 'online', 'Online'))
        else:
            tree.after(0, lambda: update_tree_item(tree, item, 'offline', 'Offline'))

def update_tree_item(tree, item, tag, status):
    if item in tree.get_children():
        tree.item(item, tags=(tag,))
        tree.set(item, column='Status', value=status)

def create_default_xlsx():
    df = pd.DataFrame(columns=['IP', 'Description'])
    df.to_excel('ips.xlsx', index=False)

def load_data():
    if not os.path.exists('ips.xlsx'):
        create_default_xlsx()
        messagebox.showinfo("Información", "El archivo 'ips.xlsx' no existía y ha sido creado. Por favor, añada las IPs y descripciones y vuelva a cargar la aplicación.")
        return

    df = pd.read_excel('ips.xlsx')

    for index, row in df.iterrows():
        item = tree.insert('', 'end', values=(row['IP'], row['Description'], 'Checking...'))
        update_status_periodically(tree, row['IP'], item)

def update_status_periodically(tree, ip, item):
    threading.Thread(target=lambda: update_status(tree, ip, item)).start()
    tree.after(5000, update_status_periodically, tree, ip, item) # 5000 milisegundos = 5 segundos

def refresh_data():
    for item in tree.get_children():
        tree.delete(item)
    load_data()

root = tk.Tk()
root.title("IP Status Checker")

tree = ttk.Treeview(root, columns=('IP', 'Description', 'Status'), show='headings')
tree.heading('IP', text='IP')
tree.heading('Description', text='Descripción')
tree.heading('Status', text='Estado')
tree.pack(fill=tk.BOTH, expand=True)

tree.tag_configure('online', background='green')
tree.tag_configure('offline', background='red')

refresh_button = tk.Button(root, text="Actualizar", command=refresh_data)
refresh_button.pack(pady=10)

load_data()

root.mainloop()