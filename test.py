import tkinter as tk

root = tk.Tk()

# 创建一个状态栏标签
status_bar = tk.Label(root, text="状态栏信息", bd=1, relief=tk.SUNKEN, anchor=tk.W)
status_bar.pack(side=tk.BOTTOM, fill=tk.X)

# 更新状态栏信息的函数
def update_status(message):
    status_bar.config(text=message)

# 模拟按钮点击更新状态栏
def button_click():
    update_status("按钮被点击了！")

button = tk.Button(root, text="点击我", command=button_click)
button.pack()

root.mainloop()