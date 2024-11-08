import logging
import multiprocessing
import os.path
import subprocess
import sys
import time
import tkinter as tk
import tkinter.ttk as ttk
from functools import partial
from tkinter import filedialog, messagebox

import psutil as psutil

import cache_utils
from cache_utils import compute_cache_data, get_all_sheet_names
from excel_utils import open_excel_sheet
from logger_utils import init_logging_basic_config

FONT_12 = ("微软雅黑", 12)
FONT_BOLD_12 = ("微软雅黑", 12, "bold")
TITLE = "Excel页签搜索大师"

last_search_text = ""

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

    def save_path():
        new_path = path_entry.get()
        sheet_name = name_entry.get()
        # sheet_name 不能为空
        if not sheet_name or sheet_name.isspace():
            # 弹窗提示
            messagebox.showerror("警告", "名称不能为空")
            return
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
        cache_utils.set_path_data(new_path, sheet_name, False, "")
        refresh_switch_dir()
        logging.info(f"更改路径为：{new_path}，名称：{sheet_name}")
        path_window.destroy()

    save_button = tk.Button(path_window, text="添加目录", command=save_path, font=FONT_12)
    save_button.grid(row=4, column=0, columnspan=2, pady=5)


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
        if path_data:
            entry_sheet_name.delete(0, tk.END)
            entry_sheet_name.insert(0, path_data["sheet_name"])

    cb_path.bind("<<ComboboxSelected>>", on_combobox_select)

    # 创建第一个文本框及标签
    text1_label = tk.Label(new_window, text="页签名称：", font=FONT_12)
    text1_label.grid(row=1, column=0)
    entry_sheet_name = tk.Entry(new_window, width=50, font=FONT_12)
    entry_sheet_name.grid(row=1, column=1)

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

    def delete_selection():
        # 先删除目录，然后获取第一个路径
        selected_path = cb_path.get()
        # 删除页签
        sheet_name = cache_utils.get_path_sheet_name(selected_path)
        cache_utils.remove_cache_sheet(sheet_name)
        # 获取sheet_name
        cache_utils.del_path_data(selected_path)
        first_path = cache_utils.get_first_path()
        if first_path is None:
            cache_utils.set_config_value("usePath", "")
            open_path_window()
        else:
            cache_utils.set_config_value("usePath", first_path)
            refresh_switch_dir()

        new_window.destroy()

    confirm_button = tk.Button(new_window, text="切换目录", command=confirm_selection, font=FONT_12)
    confirm_button.grid(row=4, column=0, pady=5)
    delete_button = tk.Button(new_window, text="删除目录", command=delete_selection, font=FONT_12)
    delete_button.grid(row=4, column=1, pady=5)


def get_second_part(s: str):
    parts = s.split('|')
    if len(parts) > 1:
        return parts[1]
    else:
        return ""


def search():
    global tree, entry_search, combo_box
    tree.delete(*tree.get_children())
    search_text = entry_search.get().strip().lower()
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

    # 按照 sheet_name 进行排序，并将空字符串排在最后面
    def custom_sort_key(full_sheet_name):
        second_part = get_second_part(full_sheet_name[0])
        if second_part == "":
            return 1, full_sheet_name[0]
        else:
            return 0, second_part

    values_to_insert.sort(key=custom_sort_key)
    first_match_index = None
    for index, values in enumerate(values_to_insert):
        item = tree.insert('', tk.END, values=values)
        if len(search_text) > 0 and get_second_part(values[0]).lower() == search_text.lower():
            # 记录第一个匹配项的索引
            first_match_index = item

    if first_match_index is not None:
        # 选中第一个匹配项
        tree.selection_set(first_match_index)
        tree.see(first_match_index)

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


def add_svn_toolbar():
    # 创建工具栏框架
    toolbar2 = tk.Frame(root, bd=1, relief='raised')
    column_num = 0

    def run_tortoise_update():
        # 弹窗确认是否更新
        usePath = cache_utils.get_config_value("usePath")
        if messagebox.askokcancel("SVN更新确认", f"确定要更新下面的目录吗？\n{usePath}"):
            subprocess.Popen(["TortoiseProc.exe", "/command:update", f"/path:{usePath}", "/closeonend:0"])

    btnSvnUpdate = tk.Button(toolbar2, text="SVN更新", command=run_tortoise_update)
    column_num += 1
    btnSvnUpdate.grid(row=row_index, column=column_num, padx=2, pady=2)

    # svn cleanup
    def run_tortoise_cleanup():
        usePath = cache_utils.get_config_value("usePath")
        subprocess.Popen(["TortoiseProc.exe", "/command:cleanup", f"/path:{usePath}", "/closeonend:0"], shell=False)

    btnSvnCleanup = tk.Button(toolbar2, text="SVN清理", command=run_tortoise_cleanup)
    column_num += 1
    btnSvnCleanup.grid(row=row_index, column=column_num, padx=2, pady=2)

    # svn revert
    def run_tortoise_revert():
        usePath = cache_utils.get_config_value("usePath")
        subprocess.Popen(["TortoiseProc.exe", "/command:revert", f"/path:{usePath}", "/closeonend:0"], shell=False)

    btnSvnRevert = tk.Button(toolbar2, text="SVN还原", command=run_tortoise_revert)
    column_num += 1
    btnSvnRevert.grid(row=row_index, column=column_num, padx=2, pady=2)

    def run_tortoise_commit():
        usePath = cache_utils.get_config_value("usePath")
        subprocess.Popen(["TortoiseProc.exe", "/command:commit", f"/path:{usePath}", "/closeonend:0"], shell=False)

    btnSvnCommit = tk.Button(toolbar2, text="SVN提交", command=run_tortoise_commit)
    column_num += 1
    btnSvnCommit.grid(row=row_index, column=column_num, padx=2, pady=2)

    def open_dir():
        usePath = cache_utils.get_config_value("usePath")
        os.startfile(usePath)

    btnSvnCommit = tk.Button(toolbar2, text="打开目录", command=open_dir)
    column_num += 1
    btnSvnCommit.grid(row=row_index, column=column_num, padx=10, pady=2)
    # 初始SVN工具栏
    toolbar2.grid(row=row_index, column=0, sticky='nsew', columnspan=3)


def open_selected_excel():
    item = tree.selection()
    if item:
        values = tree.item(item, "values")
        sheet_name = values[0]
        use_path = cache_utils.get_config_value("usePath")
        excel_name = use_path + "/" + values[1]
        open_excel_sheet(excel_name, sheet_name)
        print(f"打开页签 {sheet_name}，工作簿 {excel_name}。")


def on_up():
    selected_item = tree.selection()
    if selected_item:
        current_index = tree.index(selected_item[0])
        if current_index > 0:
            new_index = current_index - 1
            tree.selection_set(tree.get_children()[new_index])
            tree.see(tree.get_children()[new_index])


def on_down():
    selected_item = tree.selection()
    if selected_item:
        current_index = tree.index(selected_item[0])
        if current_index < len(tree.get_children()) - 1:
            new_index = current_index + 1
            tree.selection_set(tree.get_children()[new_index])
            tree.see(tree.get_children()[new_index])
    else:
        if tree.get_children():
            tree.selection_set(tree.get_children()[0])
            tree.see(tree.get_children()[0])


if __name__ == '__main__':
    # Pyinstaller fix
    multiprocessing.freeze_support()

    init_logging_basic_config()


    def close_same_exe():
        current_script = os.path.basename(__file__)
        this_exe_name = os.path.splitext(current_script)[0] + '.exe'
        for proc in psutil.process_iter():
            try:
                if proc.name() == this_exe_name and proc.pid != os.getpid():
                    proc.kill()
                    logging.info(f"关闭其他同名应用程序：{this_exe_name}")
            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                pass


    close_same_exe()

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

    # 初始显示工具栏
    toolbar.grid(row=row_index, column=0, sticky='nsew', columnspan=3)

    row_index += 1
    add_svn_toolbar()

    row_index += 1
    # 创建可选列表（Combobox）
    options = ["页签名称", "工作簿名称"]
    combo_box = ttk.Combobox(root, values=options, state='readonly', font=FONT_BOLD_12, width=10)
    combo_box.current(0)
    combo_box.grid(row=row_index, column=0)

    # 创建文本框
    entry_search = tk.Entry(root, width=80, font=FONT_BOLD_12)
    entry_search.grid(row=row_index, column=1)
    entry_search.focus_set()  # 默认激活文本框

    # 创建一个框架用于包裹按钮，模拟外边距
    button_frame = tk.Frame(root)
    button_frame.grid(row=row_index, column=2, padx=5, pady=3)

    # 创建按钮
    button = tk.Button(button_frame, text="模糊搜索", command=search, font=FONT_BOLD_12, padx=0, pady=5, width=15)
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
    tree.column('Excel Name', width=600, anchor='center')
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
        open_selected_excel()


    tree.bind("<Double-1>", on_double_click)

    def on_enter(event):
        open_selected_excel()

    # 绑定回车事件
    tree.bind("<Return>", on_enter)

    def on_enter(event):
        global last_search_text
        if last_search_text == entry_search.get().strip().lower():
            open_selected_excel()
        else:
            last_search_text = entry_search.get().strip().lower()
            search()
        return 'break'  # 阻止回车键的默认换行操作

    entry_search.bind('<Return>', on_enter)

    def on_up_key(event):
        on_up()


    def on_down_key(event):
        on_down()


    entry_search.bind("<Up>", on_up_key)
    entry_search.bind("<Down>", on_down_key)


    def on_up_key(event):
        on_up()


    def on_down_key(event):
        on_down()


    tree.bind("<Up>", on_up_key)
    tree.bind("<Down>", on_down_key)


    def on_window_load():
        # 先加载窗口，再初始化数据
        use_path = cache_utils.get_config_value("usePath")
        if cache_utils.get_path_data(use_path):
            refresh_switch_dir()
        else:
            # 获取第一个路径
            first_path = cache_utils.get_first_path()
            if first_path:
                cache_utils.set_config_value("usePath", first_path)
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

        root.destroy()
        # root.after(100, root.destroy)


    root.protocol("WM_DELETE_WINDOW", on_close)


    def center_window(root):
        # 获取屏幕宽度和高度
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()
        # 确保窗口布局已经完成，再获取窗口的实际宽度和高度
        root.update_idletasks()
        # 获取窗口宽度和高度
        window_width = root.winfo_width()
        window_height = root.winfo_height()
        # 计算窗口左上角的坐标
        x = (screen_width - window_width) / 2
        y = (screen_height - window_height) / 2
        logging.debug(
            f"屏幕宽度：{screen_width}，屏幕高度：{screen_height}，窗口宽度：{window_width}，窗口高度：{window_height}")
        root.geometry('+%d+%d' % (x, y))


    center_window(root)
    root.mainloop()
