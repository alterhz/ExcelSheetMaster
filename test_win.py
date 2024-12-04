import tkinter as tk

FONT_12 = ("微软雅黑", 12)

mini_window = None


def get_mini_window():
    global mini_window
    if mini_window is None:
        """打开子窗口的函数"""
        child_window = tk.Toplevel(root)
        child_window.title("子窗口")
        # 设置子窗口尺寸为100*30
        child_window.geometry("350x50")

        # 禁用子窗口的最大化、最小化和关闭按钮
        child_window.resizable(False, False)
        child_window.protocol("WM_DELETE_WINDOW", lambda: None)
        child_window.attributes('-toolwindow', True)
        # 设置窗口最上方
        child_window.attributes('-topmost', True)

        # 搜索文本框
        search_entry = tk.Entry(child_window, width=30, font=FONT_12)
        # 使用grid布局，将文本框放置在第0列，设置一定的内边距使其看起来更协调
        search_entry.grid(row=0, column=0, padx=5, pady=5)

        # 在子窗口中添加一个“搜索”按钮
        maximize_button = tk.Button(child_window, text="搜索", font=FONT_12,
                                    command=lambda: mini_search(root, child_window, search_entry))
        # 将按钮也放置在第0行，第1列，设置内边距与文本框统一
        maximize_button.grid(row=0, column=1, padx=5, pady=5)
        mini_window = child_window
    return mini_window


def mini_search(root, child_window, search_entry=None):
    # 显示主窗体
    root.deiconify()
    # 获取文本框内容
    search_text = search_entry.get()
    print(f'搜索内容：{search_text}')
    # 隐藏
    child_window.withdraw()


root = tk.Tk()
root.title("主窗口")
root.geometry("300x200")


# 定义显示第二个窗口的函数
def show_mini_window():
    root.withdraw()  # 隐藏主窗口
    get_mini_window().deiconify()  # 显示子窗口


btn = tk.Button(root, text="点击切换窗口", command=show_mini_window)
btn.pack()

root.mainloop()
