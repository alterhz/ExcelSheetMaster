import tkinter as tk
from tkinter import messagebox


def open_child_window():
    """打开子窗口的函数"""
    child_window = tk.Toplevel(root)
    child_window.title("子窗口")
    # 设置子窗口尺寸为100*30
    child_window.geometry("400x30")

    # 禁用子窗口的最大化、最小化和关闭按钮
    child_window.resizable(False, False)
    child_window.protocol("WM_DELETE_WINDOW", lambda: None)
    child_window.attributes('-toolwindow', True)

    # 在子窗口中添加一个“最大化”按钮
    maximize_button = tk.Button(child_window, text="最大化", command=lambda: maximize(root, child_window))
    maximize_button.pack()


def maximize(root, child_window):
    # 显示主窗体
    root.deiconify()

    # 销毁子窗体
    child_window.destroy()


def on_minimize():
    print("我就是不最小化")

    open_child_window()
    # 隐藏主窗口
    root.withdraw()


root = tk.Tk()
root.title("主窗口")
root.geometry("800x600")

# 创建打开子窗口的按钮
open_button = tk.Button(root, text="最小化", command=open_child_window)
open_button.pack()

root.mainloop()
