import logging
import os.path
import tkinter as tk
import tkinter.ttk as ttk
from tkinter import filedialog, messagebox
from tkinter.font import Font

import cache_utils
from cache_utils import compute_cache_data, get_all_sheet_names
from config import ConfigHandler
from excel_utils import open_excel_sheet
from logger_utils import init_logging_basic_config

init_logging_basic_config()

root = tk.Tk()
root.title("Excel 页签搜索工具")

# 禁用最大化按钮
root.resizable(False, False)

# 创建菜单栏
menu_bar = tk.Menu(root)

# 创建设置菜单
setting_menu = tk.Menu(menu_bar, tearoff=0)

FONT_12 = ("微软雅黑", 12)
FONT_14 = ("微软雅黑", 14)


def open_path_window():
    path_window = tk.Toplevel(root)
    path_window.grab_set()
    path_window.title("设置路径")

    # 创建路径输入框及标签
    path_label = tk.Label(path_window, text="路径：", font=FONT_12)
    path_label.grid(row=0, column=0)
    path_entry = tk.Entry(path_window, width=50, font=FONT_12)
    path_entry.grid(row=0, column=1)

    def select_path_from_button():
        selected_path = select_path()
        if selected_path:
            path_entry.delete(0, tk.END)
            path_entry.insert(0, selected_path)

    select_path_button = tk.Button(path_window, text="选择路径", command=select_path_from_button,
                                   font=FONT_12)
    select_path_button.grid(row=0, column=2)

    # 创建名称输入框及标签
    name_label = tk.Label(path_window, text="名称：", font=FONT_12)
    name_label.grid(row=1, column=0)
    name_entry = tk.Entry(path_window, width=50, font=FONT_12)
    name_entry.grid(row=1, column=1)

    # 创建描述输入框及标签
    description_label = tk.Label(path_window, text="描述：", font=FONT_12)
    description_label.grid(row=2, column=0)
    description_entry = tk.Text(path_window, height=5, width=50, font=FONT_12)
    description_entry.grid(row=2, column=1)

    # 创建是否包含子目录的复选框及标签
    include_subdirs_var = tk.BooleanVar()
    include_subdirs_checkbox = tk.Checkbutton(path_window, text="包含子目录", variable=include_subdirs_var,
                                              font=FONT_12)
    include_subdirs_checkbox.grid(row=3, column=0, columnspan=2)

    def save_path():
        new_path = path_entry.get()
        sheet_name = name_entry.get()
        description = description_entry.get("1.0", tk.END).strip()
        include_subdirs = include_subdirs_var.get()
        # 判断路径是否存在
        if not os.path.exists(new_path):
            # 弹窗提示
            messagebox.showerror("警告", "路径不存在")
            return
        if cache_utils.get_path_data(new_path):
            # 弹窗提示
            messagebox.showerror("警告", "已添加过该路径")
            return
        if cache_utils.exist_sheet(sheet_name):
            # 弹窗提示
            messagebox.showerror("警告", "页签名称已存在")
            return
        cache_utils.set_config_value("usePath", new_path)
        cache_utils.set_path_data(new_path, sheet_name, include_subdirs, description)
        compute_cache_data()
        logging.info(f"更改路径为：{new_path}，名称：{sheet_name}，描述：{description}，是否包含子目录：{include_subdirs}")
        path_window.destroy()

    save_button = tk.Button(path_window, text="保存", command=save_path, font=FONT_12)
    save_button.grid(row=4, column=0, columnspan=2)


def change_path():
    open_path_window()


# 添加更改路径选项
setting_menu.add_command(label="添加目录", command=change_path, font=FONT_12)

# 将设置菜单添加到菜单栏
menu_bar.add_cascade(label="设置", menu=setting_menu)

# 将菜单栏添加到窗口
root.config(menu=menu_bar)

# 配置
cache_utils.create_config_sheet_header()


def select_path():
    root = tk.Tk()
    root.withdraw()
    path = filedialog.askdirectory()
    return path


use_path = cache_utils.get_config_value("usePath")
if cache_utils.get_path_data(use_path):
    compute_cache_data()
else:
    open_path_window()

# 创建可选列表（Combobox）
options = ["页签名称", "工作簿名称"]
combo_box = ttk.Combobox(root, values=options, state='readonly', font=FONT_14, width=10)
combo_box.current(0)
combo_box.grid(row=0, column=0)

# 创建文本框
entry = tk.Entry(root, width=50, font=FONT_14)
entry.grid(row=0, column=1)
entry.focus_set()  # 默认激活文本框

# 创建一个框架用于包裹按钮，模拟外边距
button_frame = tk.Frame(root)
button_frame.grid(row=0, column=2, padx=5, pady=3)


# 创建按钮
def search():
    tree.delete(*tree.get_children())
    search_text = entry.get().strip().lower()
    sheet_names = get_all_sheet_names()
    search_type_index = combo_box.current()
    print(f"搜索类型：{search_type_index}，搜索内容：{search_text}")
    for item in sheet_names:
        excel_name = item["name"]
        sheet_name = item["sheet_name"]
        if search_type_index == 0:
            # 页签搜索
            if search_text in sheet_name.lower():
                tree.insert('', tk.END, values=(sheet_name, excel_name))
        else:
            # 工作簿搜索
            if search_text in excel_name.lower():
                tree.insert('', tk.END, values=(sheet_name, excel_name))


button = tk.Button(button_frame, text="模糊搜索", command=search, font=FONT_14, padx=15, pady=5)
button.pack()

# 创建一个框架用于包裹 Treeview
tree_frame = tk.Frame(root)
tree_frame.grid(row=1, column=0, columnspan=3, padx=5, pady=3)

style = ttk.Style()
# 修改 Treeview 的字体大小
style.configure("Treeview", font=FONT_12)
style.configure("Treeview.Heading", font=("微软雅黑", 12, "bold"))

# 创建 Treeview
tree = ttk.Treeview(tree_frame, columns=('Sheet Name', 'Excel Name'), show='headings')
tree.heading('Sheet Name', text='页签名称')
tree.heading('Excel Name', text='工作簿名称')
tree.column('Sheet Name', width=500, anchor='center')
tree.column('Excel Name', width=400, anchor='center')
# 设置 Treeview 的高度为 10 行（可根据实际需求调整）
tree.configure(height=30)
# 创建垂直滚动条
v_scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
# 将 Treeview 的 yview 方法与垂直滚动条关联
tree.configure(yscrollcommand=v_scrollbar.set)
tree.pack(fill=tk.BOTH, expand=True)


def on_double_click(event):
    item = tree.selection()
    if item:
        values = tree.item(item, "values")
        sheet_name = values[0]
        excel_name = use_path + "/" + values[1]
        open_excel_sheet(excel_name, sheet_name)
        print(f"打开页签 {sheet_name}，工作簿 {excel_name}。")


tree.bind("<Double-1>", on_double_click)


# 绑定回车键事件到按钮的点击事件
def on_enter(event):
    search()
    return 'break'  # 阻止回车键的默认换行操作


entry.bind('<Return>', on_enter)

root.mainloop()
