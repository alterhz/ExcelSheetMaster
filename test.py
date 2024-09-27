import tkinter as tk

def do_something1():
    print("功能 1 被触发")

def do_something2():
    print("功能 2 被触发")

def add_button():
    global button_count
    button_count += 1
    new_button = tk.Button(toolbar, text=f"新增功能 {button_count}", command=lambda: print(f"新增功能 {button_count} 被触发"))
    new_button.grid(row=button_count - 1, column=button_count - 1, padx=2, pady=2)

def remove_all_buttons():
    for widget in toolbar.winfo_children():
        if isinstance(widget, tk.Button):
            widget.destroy()

def toggle_toolbar():
    global toolbar
    if toolbar.winfo_viewable():
        toolbar.grid_remove()
    else:
        toolbar.grid(row=0, column=0, sticky='nsew')

root = tk.Tk()

# 创建工具栏框架
toolbar = tk.Frame(root, bd=1, relief='raised')

button1 = tk.Button(toolbar, text="功能 1", command=do_something1)
button1.grid(row=0, column=0, padx=2, pady=2)

button2 = tk.Button(toolbar, text="功能 2", command=do_something2)
button2.grid(row=0, column=1, padx=2, pady=2)

add_button_btn = tk.Button(root, text="添加按钮", command=add_button)
add_button_btn.grid(row=1, column=0)

remove_all_btn = tk.Button(root, text="删除所有按钮", command=remove_all_buttons)
remove_all_btn.grid(row=1, column=1)

toggle_button = tk.Button(root, text="显示/隐藏工具栏", command=toggle_toolbar)
toggle_button.grid(row=2, column=0)

# 配置网格布局参数
root.grid_rowconfigure(0, weight=0)
root.grid_columnconfigure(0, weight=1)

button_count = 2

# 初始显示工具栏
toolbar.grid(row=0, column=0, sticky='nsew')

root.mainloop()