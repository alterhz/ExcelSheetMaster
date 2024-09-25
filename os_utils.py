import os


def get_tree_file_names(directory, extension):
    """
    获取指定目录下的所有文件(包括子目录)
    :param extension: 文件后缀（扩展名）
    :param directory: 指定目录
    :return: 文件名列表
    """
    file_names = []  # 用于存储文件名的列表

    # 遍历目录中的所有文件和子目录
    for root, dirs, files in os.walk(directory):
        # 将当前目录下的所有文件添加到列表中
        for file in files:
            if file.endswith(extension):
                file_names.append(os.path.join(root, file))

    return file_names


def get_current_file_names(directory, extension):
    """
    获取指定目录下的所有文件(不包括子目录)
    :param directory: 指定目录
    :param extension: 文件后缀（扩展名）
    :return: 文件名列表
    """
    file_names = []  # 用于存储文件名的列表

    # 获取指定目录下的所有文件和子目录
    items = os.listdir(directory)

    # 遍历所有文件和子目录
    for item in items:
        # 构建文件的完整路径
        item_path = os.path.join(directory, item)
        # 检查是否是文件
        if os.path.isfile(item_path) and not item.startswith("~$"):
            # 检查文件是否以指定后缀结尾
            if item.endswith(extension):
                file_names.append(item_path)

    return file_names

# 获取目录名字
def get_child_directory_names(directory):
    """
    获取指定目录下的所有文件夹(不包括子目录)
    :param directory: 指定目录
    :return: 文件夹名列表
    """
    directory_names = []  # 用于存储文件夹名的列表

    # 获取指定目录下的所有文件和子目录
    items = os.listdir(directory)

    # 遍历所有文件和子目录
    for item in items:
        # 构建文件的完整路径
        item_path = os.path.join(directory, item)
        # 检查是否是文件夹
        if os.path.isdir(item_path):
            directory_names.append(item)

    return directory_names


def read_file_to_list():
    """
    此函数用于读取指定文件内容并逐行保存到列表中。
    返回值为包含文件内容每一行的列表。
    """
    file_path = "filelist.txt"
    lines = []
    try:
        with open(file_path, 'r') as file:
            for line in file:
                lines.append(line.strip())
    except FileNotFoundError:
        return lines
    return lines

def get_filename_from_path(file_path):
    """
    此函数接收一个完整的文件路径，返回其中的文件名。
    参数：
        file_path：完整的文件路径字符串。
    返回值：
        文件名字符串。
    """
    import os
    return os.path.basename(file_path)

if __name__ == "__main__":
    # 测试函数
    directory_path = "./国际版翻译"
    file_names_list = get_current_file_names(directory_path, ".xlsx")
    for file_name in file_names_list:
        print(file_name)
