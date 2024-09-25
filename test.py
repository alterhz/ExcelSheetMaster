import tkinter as tk
from tkinter import ttk

def create_gui():
    root = tk.Tk()

    # 创建可选列表（Combobox）
    options = ["Option 1", "Option 2", "Option 3"]
    combo_box = ttk.Combobox(root, values=options)
    combo_box.pack()

    # 创建页签名称（Label）
    tab_label = tk.Label(root, text="页签名称")
    tab_label.pack()

    # 创建工作簿名称（Entry）
    workbook_entry = tk.Entry(root)
    workbook_entry.pack()

    root.mainloop()

create_gui()