"""
    物料盘点数据排序
    1.对‘规格型号’列中的字段按[0402]这种规格排序
    2.对‘规格型号’列中的字段按‘ΩKMV’排序大小排序
    3.不满足1、2排序要求的物料按'物料名称'列排序
    注：文件名必须另存为csv格式、文件名不能含有中文
"""

import numpy as np
import pandas as pd
import re
import os
import time
import xlrd


# 获取大小
def get_size(model):
    model_size = re.findall(r'-\d*.?\d+[ΩKMV]', model)
    try:
        model_size = model_size[0][1:]
        if 'Ω' in model_size:
            model_size = 'a' + model_size
        elif 'K' in model_size:
            model_size = 'b' + model_size
        elif 'M' in model_size:
            model_size = 'c' + model_size
        elif 'V' in model_size:
            model_size = 'd' + model_size
        else:
            print('ERROR: '+ model_size)
    except:
        model_size = np.nan
    return model_size

# 获取id
def get_id(model):
    model_num = re.findall(r'-\[\d+\W*\d*\]', model)
    try:
        model_num = model_num[0][2:-1]
    except:
        model_num = np.nan
    return model_num

# 从用户的文档目录中获取文件
# user_dir = os.environ.get('USERPROFILE')
user_dir = os.getcwd()
docs_dir = os.path.join(user_dir, 'doc')

# 输入文件（带整理文件）
file_path = os.path.join(docs_dir, '仓库盘点.xls')

# 输出文件（整理好了的文件）
out_file_name = os.path.join(docs_dir, '仓库盘点_自动化整理.xls')

# 打开输入文件
content = xlrd.open_workbook(file_path, encoding_override='gbk')

if os.path.isfile(file_path):
    df = pd.read_excel(content, header=3, engine='xlrd')
else:
    print("错误。找不到文件：%s，请检查目录和文件名是否正确。" % file_path)
    exit(1)
    raise FileNotFoundError("找不到文件：%s，请检查目录和文件名是否正确。" % file_path)

# Sep_statics_79stock.csv
# df = pd.read_csv(file_path, encoding='gbk', headers=3)

df = df.dropna(axis=0, how='all')
df = df.set_index(keys='序号')

# 增加两列用于排序
model_size = df['规格型号'].apply(get_size)
model_num = df['规格型号'].apply(get_id)
df.insert(loc=7, column = 'model_size', value=model_size)
df.insert(loc=8, column = 'model_id', value=model_num)

# 排序
df2 = df.sort_values(by=['model_id', 'model_size', '物料名称'])

# 去除增加的列
df2.pop('model_id')
df2.pop('model_size')


# 写入excel文件
writer = pd.ExcelWriter(out_file_name)
df2.to_excel(writer,index=True)
writer.save()
print('成功生成整理好的新文件，新文件路径为：%s'%out_file_name)
