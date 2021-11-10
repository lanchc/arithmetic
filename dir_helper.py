# -*- coding:utf-8 _*-
# @time     2021/11/10 1:47 下午
# @explain  arithmetic
import os
import shutil
import datetime


# 获取目录
def project_root_dir():
    return os.path.abspath(os.path.dirname(__file__))


# 模版文件
def template_path(file_name='template.docx'):
    return project_root_dir() + "/" + str(file_name)


# 保存文档最终位置
def docx_rename():
    time_form = '%Y-%m-%d-%H%M%S'
    time_name = datetime.datetime.now().strftime(time_form)
    return project_root_dir() + '/dist/' + time_name + '.docx'


# 检查创建文件夹
def checkup_dir(path):
    if os.path.exists(path):
        shutil.rmtree(path)
    os.mkdir(path)
