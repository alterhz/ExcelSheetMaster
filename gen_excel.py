import os
import random
import string
from openpyxl import Workbook

# 创建 config 目录，如果已存在则不创建
config_dir = 'config'
if not os.path.exists(config_dir):
    os.mkdir(config_dir)

# 随机有意义的单词列表，可以根据需要扩展
words = ['apple', 'banana', 'cherry', 'date', 'elderberry', 'fig', 'grape', 'honeydew', 'kiwi', 'lemon',
         'mango', 'orange', 'peach', 'quince', 'raspberry', 'strawberry', 'tomato', 'watermelon', 'zucchini',
         'broccoli', 'carrot', 'cabbage', 'potato', 'onion', 'garlic', 'spinach', 'lettuce', 'asparagus',
         'cauliflower', 'eggplant', 'green_beans', 'peas', 'corn', 'beet', 'radish', 'cucumber', 'pumpkin',
         'squash', 'mushroom', 'pepper', 'celery', 'artichoke', 'avocado', 'cantaloupe', 'honeydew_melon',
         'kumquat', 'lychee', 'nectarine', 'papaya', 'persimmon', 'plum', 'pomegranate', 'tangerine',
         'apricot', 'blueberry', 'cranberry', 'durian', 'figs', 'gooseberry', 'guava', 'huckleberry', 'jackfruit',
         'jujube', 'loganberry', 'loquat', 'mulberry', 'olive', 'passion_fruit', 'pear', 'pineapple', 'rambutan',
         'starfruit', 'ugli_fruit']

for i in range(500):
    # 生成随机文件名
    filename = random.choice(words) + '.xlsx'
    wb = Workbook()
    num_sheets = random.randint(1, 10)
    for _ in range(num_sheets):
        # 生成随机页签名
        sheet_name = random.choice(words)
        wb.create_sheet(sheet_name)
    wb.save(os.path.join(config_dir, filename))