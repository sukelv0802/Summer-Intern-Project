import tkinter as tk
from tkinter import ttk
import ttkbootstrap as ttks
from ttkbootstrap.constants import *
from tkinter import messagebox
import time
import serial
import pandas as pd
import re

def export_to_excel():
    try:
        data = [tree.item(item, "values") for item in tree.get_children()]
        df = pd.DataFrame(data, columns=["Timestamp", "Data"])
        df.to_excel("data.xlsx", index=False)
        tk.messagebox.showinfo("Success", "Data has been successfully exported to data.xlsx")
    except Exception as e:
        tk.messagebox.showerror("Error", f"Failed to export data: {str(e)}")

# To configure the serial connection
baudrate = 115200
serialConnection = serial.Serial("COM10", baudrate)

threshold_value = None
def set_threshold():
    global threshold_value
    try:
        threshold_value = int(threshold_entry.get())
    except ValueError:
        print("Please enter a valid integer value for the threshold.")

update_flag = False
def start_update():
    global update_flag
    if not update_flag:     # A condition to prevent starting update when it's already running, was breaking without this
        update_flag = True
        update_text()

def stop_update():
    global update_flag
    update_flag = False

def update_text():
    global threshold_value, update_flag
    if update_flag:
        data = serialConnection.readline()
        if data == b"EOF":
            serialConnection.close()
        else:
            value = data.decode('utf-8').strip()
            timestamp = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
            try:
                numeric_value = int(value.split(':')[-1].strip())
                if threshold_value is not None and numeric_value > threshold_value:
                    tree.insert('', 'end', values=(timestamp, value), tags=('highlight',))
                else:
                    tree.insert('', 'end', values=(timestamp, value))
            except ValueError:
                tree.insert('', 'end', values=(timestamp, value))
        root.after(500, update_text)
        
def treeview_sort_column(tv, col, reverse):
    l = [(tv.set(k, col), k) for k in tv.get_children('')]
    l.sort(reverse=reverse)

    for index, (val, k) in enumerate(l):
        tv.move(k, '', index)

    tv.heading(col, command=lambda: treeview_sort_column(tv, col, not reverse))

def on_motion(event):
    region = tree.identify("region", event.x, event.y)
    if region == "heading":
        tree.config(cursor="hand2")
    else:
        tree.config(cursor="")

root = ttks.Window(themename='darkly')

tree = ttk.Treeview(root)
tree['columns'] = ('Timestamp', 'Data')
tree.heading('Timestamp', text='Timestamp', command=lambda: treeview_sort_column(tree, 'Timestamp', False))
tree.heading('Data', text='Data', command=lambda: treeview_sort_column(tree, 'Data', False))
tree.column('Timestamp', width=150, anchor=CENTER)
tree.column('Data', width=650, anchor=W)
tree.tag_configure('highlight', background='yellow', foreground='black')



options_bar = ttks.Frame(root, style='darkly', width=400)
options_bar.pack(fill=tk.Y, side=tk.LEFT)

threshold_label = ttks.Label(options_bar, text="Threshold")
threshold_label.pack()
threshold_label.config(font=("Arial", 20, "bold"))
threshold_entry = ttks.Entry(options_bar)
threshold_entry.pack()
confirm_button = ttks.Button(options_bar, text="Confirm", command=set_threshold)
confirm_button.pack()

button_frame = ttks.Frame(root, style='darkly')
button_frame.pack(side=tk.BOTTOM, fill=tk.X)
start_button = ttks.Button(button_frame, text="Start", command=start_update)
start_button.pack(side=tk.LEFT, padx=5, pady=5)
stop_button = ttks.Button(button_frame, text="Stop", command=stop_update)
stop_button.pack(side=tk.LEFT, padx=5, pady=5)
export_button = ttks.Button(button_frame, text="Export to Excel", command=export_to_excel)
export_button.pack(side=tk.LEFT, padx=5, pady=5)

tree.pack(side='top', fill='both', expand=True)
tree.bind("<Motion>", on_motion)

root.mainloop()
