# -*- coding:utf-8 _*-
# @time     2021/11/22 11:42 上午
# @explain  arithmetic
import tkinter as tk
import tkinter.filedialog as td
# 文件选择
file_name = td.askopenfilename(
    title='打开文件',
    initialdir='~/desktop/',
    filetypes=[('py文件', '*.py'), ('文档', '*.doc')]
)
print(file_name)