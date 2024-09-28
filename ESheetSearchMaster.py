import logging
import multiprocessing
import os.path
import subprocess
import time
import tkinter as tk
import tkinter.ttk as ttk
from functools import partial
from tkinter import filedialog, messagebox

import cache_utils
from cache_utils import compute_cache_data, get_all_sheet_names
from excel_utils import open_excel_sheet
from logger_utils import init_logging_basic_config

FONT_12 = ("微软雅黑", 12)
FONT_BOLD_12 = ("微软雅黑", 12, "bold")
TITLE = "Excel页签搜索大师"

root = tk.Tk()
root.title(TITLE)


def select_path():
    root = tk.Tk()
    root.withdraw()
    path = filedialog.askdirectory()
    return path


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
        refresh_switch_dir()
        logging.info(f"更改路径为：{new_path}，名称：{sheet_name}，描述：{description}，是否包含子目录：{include_subdirs}")
        path_window.destroy()

    save_button = tk.Button(path_window, text="保存", command=save_path, font=FONT_12)
    save_button.grid(row=4, column=0, columnspan=2)


def refresh_switch_dir():
    t1 = time.time()
    compute_cache_data()
    refresh_menu_switch_dir()
    refresh_toolbar()
    path = cache_utils.get_config_value("usePath")
    root.title(f"{TITLE} - 设置路径：{path}")
    search()
    t2 = time.time()
    logging.debug(f"刷新目录耗时：{t2 - t1:.2f} 秒。")


def change_path_window():
    new_window = tk.Toplevel(root)
    new_window.grab_set()
    new_window.title("新窗口设置")

    # 创建下拉选择列表及标签
    options_label = tk.Label(new_window, text="名称：", font=FONT_12)
    options_label.grid(row=0, column=0)
    all_path_data = cache_utils.get_all_path_data()
    options = [row["path"] for row in all_path_data]
    cb_path = ttk.Combobox(new_window, values=options, state='readonly', font=FONT_12, width=50)
    cb_path.grid(row=0, column=1)

    # 为 Combobox 添加选择事件
    def on_combobox_select(event):
        selected_option = cb_path.get()
        logging.info(f"选择了：{selected_option}")
        path_data = cache_utils.get_path_data(selected_option)
        entry_sheet_name.delete(0, tk.END)
        entry_sheet_name.insert(0, path_data["sheet_name"])
        # 设置是否包含子目录的复选框
        var_ckb_include_sub_dirs.set(path_data["includeSubDir"])
        # 设置描述文本框
        txt_desc.delete("1.0", tk.END)
        txt_desc.insert("1.0", path_data["desc"])

    cb_path.bind("<<ComboboxSelected>>", on_combobox_select)

    # 创建第一个文本框及标签
    text1_label = tk.Label(new_window, text="页签名称：", font=FONT_12)
    text1_label.grid(row=1, column=0)
    entry_sheet_name = tk.Entry(new_window, width=50, font=FONT_12)
    entry_sheet_name.grid(row=1, column=1)

    # 创建第二个勾选框及标签
    text2_label = tk.Label(new_window, text="是否包含子目录：", font=FONT_12)
    text2_label.grid(row=2, column=0)
    var_ckb_include_sub_dirs = tk.BooleanVar()
    text2_checkbox = tk.Checkbutton(new_window, variable=var_ckb_include_sub_dirs, font=FONT_12)
    # text2_checkbox禁用编辑
    text2_checkbox["state"] = "disabled"
    text2_checkbox.grid(row=2, column=1)

    # 创建第三个文本框及标签
    text3_label = tk.Label(new_window, text="描述：", font=FONT_12)
    text3_label.grid(row=3, column=0)
    txt_desc = tk.Text(new_window, height=5, width=50, font=FONT_12)
    txt_desc.grid(row=3, column=1)

    # 获取当前路径
    use_path = cache_utils.get_config_value("usePath")
    cb_path.set(use_path)
    on_combobox_select(None)

    def confirm_selection():
        selected_path = cb_path.get()
        cache_utils.set_config_value("usePath", selected_path)
        # 加载
        refresh_switch_dir()
        new_window.destroy()

    confirm_button = tk.Button(new_window, text="确认选择", command=confirm_selection, font=FONT_12)
    confirm_button.grid(row=4, column=0, columnspan=2)


def search():
    global tree, entry, combo_box
    tree.delete(*tree.get_children())
    search_text = entry.get().strip().lower()
    sheet_names = get_all_sheet_names()
    search_type_index = combo_box.current()
    values_to_insert = []
    for item in sheet_names:
        excel_name = item["name"]
        sheet_name = item["sheet_name"]
        if search_type_index == 0:
            # 页签搜索
            if search_text in sheet_name.lower():
                values_to_insert.append((sheet_name, excel_name))
        else:
            # 工作簿搜索
            if search_text in excel_name.lower():
                values_to_insert.append((sheet_name, excel_name))
    for values in values_to_insert:
        tree.insert('', tk.END, values=values)
    # 更新状态栏，显示搜索结果数量
    status_bar.config(text=f"搜索到 {len(values_to_insert)} 条结果。")


def change_use_path(path):
    cache_utils.set_config_value("usePath", path)
    refresh_switch_dir()


def refresh_menu_switch_dir():
    global switch_menu
    # 点击选择目录删除所有目录
    while switch_menu.index('end') is not None and switch_menu.index('end') >= 0:
        switch_menu.delete(switch_menu.index('end'))
    for row in cache_utils.get_all_path_data():
        # 添加命令和参数，以便在点击菜单项时执行相应操作
        command_with_arg = partial(change_use_path, row["path"])
        command_text = "[" + row["sheet_name"] + "] " + row["path"]
        if row["path"] == cache_utils.get_config_value("usePath"):
            command_text += " ✔"
        switch_menu.add_command(label=command_text, command=command_with_arg)


def refresh_toolbar():
    column_index = 0
    for widget in toolbar.winfo_children():
        if isinstance(widget, tk.Button):
            widget.destroy()
    for row in cache_utils.get_all_path_data():
        # 添加命令和参数，以便在点击菜单项时执行相应操作
        command_with_arg = partial(change_use_path, row["path"])
        command_text = row["sheet_name"]
        if row["path"] == cache_utils.get_config_value("usePath"):
            command_text += " ✔"
        button1 = tk.Button(toolbar, text=command_text, command=command_with_arg)
        button1.grid(row=row_index, column=column_index, padx=2, pady=2)
        column_index += 1


if __name__ == '__main__':
    init_logging_basic_config()

    # Pyinstaller fix
    multiprocessing.freeze_support()

    cache_utils.start_back_thread()

    # 禁用最大化按钮
    root.resizable(False, False)

    # 创建菜单栏
    menu_bar = tk.Menu(root)

    # 创建设置菜单
    setting_menu = tk.Menu(menu_bar, tearoff=0)

    # 添加更改路径选项
    setting_menu.add_command(label="添加目录", command=open_path_window)
    setting_menu.add_command(label="查看目录", command=change_path_window)

    # 将设置菜单添加到菜单栏
    menu_bar.add_cascade(label="设置", menu=setting_menu)

    # # 添加动态菜单项
    switch_menu = tk.Menu(menu_bar, tearoff=0)
    menu_bar.add_cascade(label="切换目录", menu=switch_menu)

    # 将菜单栏添加到窗口
    root.config(menu=menu_bar)

    # 行索引
    row_index = 0

    # 创建工具栏框架
    toolbar = tk.Frame(root, bd=1, relief='raised')

    # 配置网格布局参数
    root.grid_rowconfigure(0, weight=0)
    root.grid_columnconfigure(0, weight=1)

    # 初始显示工具栏
    toolbar.grid(row=row_index, column=0, sticky='nsew', columnspan=3)

    row_index += 1
    # 创建工具栏框架
    toolbar2 = tk.Frame(root, bd=1, relief='raised')


    def run_tortoise_update():
        # 弹窗确认是否更新
        usePath = cache_utils.get_config_value("usePath")
        if messagebox.askokcancel("SVN更新确认", f"确定要更新下面的目录吗？\n{usePath}"):
            subprocess.Popen(["TortoiseProc.exe", "/command:update", f"/path:{usePath}", "/closeonend:0"])


    btnSvnUpdate = tk.Button(toolbar2, text="SVN更新", command=run_tortoise_update)
    btnSvnUpdate.grid(row=row_index, column=1, padx=2, pady=2)


    def run_tortoise_commit():
        usePath = cache_utils.get_config_value("usePath")
        subprocess.Popen(["TortoiseProc.exe", "/command:commit", f"/path:{usePath}", "/closeonend:0"], shell=False)


    btnSvnCommit = tk.Button(toolbar2, text="SVN提交", command=run_tortoise_commit)
    btnSvnCommit.grid(row=row_index, column=2, padx=2, pady=2)

    # 初始显示工具栏
    toolbar2.grid(row=row_index, column=0, sticky='nsew', columnspan=3)

    row_index += 1
    # 创建可选列表（Combobox）
    options = ["页签名称", "工作簿名称"]
    combo_box = ttk.Combobox(root, values=options, state='readonly', font=FONT_BOLD_12, width=10)
    combo_box.current(0)
    combo_box.grid(row=row_index, column=0)

    # 创建文本框
    entry = tk.Entry(root, width=65, font=FONT_BOLD_12)
    entry.grid(row=row_index, column=1)
    entry.focus_set()  # 默认激活文本框

    # 创建一个框架用于包裹按钮，模拟外边距
    button_frame = tk.Frame(root)
    button_frame.grid(row=row_index, column=2, padx=5, pady=3)

    # 创建按钮
    button = tk.Button(button_frame, text="模糊搜索", command=search, font=FONT_BOLD_12, padx=15, pady=5)
    button.pack()

    row_index += 1
    # 创建一个框架用于包裹 Treeview
    tree_frame = tk.Frame(root)
    tree_frame.grid(row=row_index, column=0, columnspan=3, padx=5, pady=3)

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

    row_index += 1
    # 创建状态栏标签
    status_bar = tk.Label(root, text="状态栏信息", bd=1, relief=tk.SUNKEN)
    status_bar.grid(row=row_index, column=0, sticky=tk.W + tk.E, columnspan=3)


    def on_double_click(event):
        item = tree.selection()
        if item:
            values = tree.item(item, "values")
            sheet_name = values[0]
            use_path = cache_utils.get_config_value("usePath")
            excel_name = use_path + "/" + values[1]
            open_excel_sheet(excel_name, sheet_name)
            print(f"打开页签 {sheet_name}，工作簿 {excel_name}。")


    tree.bind("<Double-1>", on_double_click)


    def on_enter(event):
        search()
        return 'break'  # 阻止回车键的默认换行操作


    entry.bind('<Return>', on_enter)


    def on_window_load():
        # 先加载窗口，再初始化数据
        use_path = cache_utils.get_config_value("usePath")
        if cache_utils.get_path_data(use_path):
            refresh_switch_dir()
        else:
            open_path_window()


    root.after_idle(on_window_load)

    running = True

    thread_idle = False

    def heartbeat():
        cache_utils.run_thread()
        if cache_utils.is_all_empty():
            global thread_idle
            if not thread_idle:
                thread_idle = True
                status_bar.config(text="Excel页签加载完毕...")
        else:
            thread_idle = False
            status_bar.config(
                text="正在加载Excel页签, 待加载Excel数量：" + str(cache_utils.get_waiting_run_excel_count()))

        if running:
            root.after(100, heartbeat)


    heartbeat()


    def on_close():
        global running
        running = False
        # 停止后台线程
        cache_utils.stop_back_thread()
        # 关闭缓存文件
        cache_utils.close_cache()

        root.after(100, root.destroy)

    root.protocol("WM_DELETE_WINDOW", on_close)

    root.mainloop()


