# -*- coding:utf-8 -*-

import os
import shutil
import time

import PySimpleGUI as sg
from pdf2docx import Converter
from xmind2testcase.zentao import xmind_to_zentao_csv_file

# 存储路径识别
isExists=os.path.exists('D:\@@@')
if not isExists:
    os.mkdir('D:\@@@')
else:
    pass

# 获取当前时间
time_tup=time.localtime(time.time())
format_time="%Y-%m-%d_%H%M%S"
cur_time=time.strftime(format_time,time_tup)

def pdf_conversion():

    pdf_file = address
    wordfile_path = 'D://@@@//wordfile_{}.docx'.format(cur_time)

    cv = Converter(pdf_file)
    cv.convert(wordfile_path, start=0, end=None)
    cv.close()

def xmind_conversion():

    xmind_file = address
    csvfile_path = 'D://@@@'

    old_csvfile = xmind_to_zentao_csv_file(xmind_file)
    old_csvfile_path = os.path.split(old_csvfile)
    new_csvfile = old_csvfile_path[0] + '\csvfile_{}.docx'.format(cur_time)
    os.rename(old_csvfile, new_csvfile)
    shutil.move(new_csvfile, csvfile_path)


layout = [
    [sg.Radio('PDF转word', 'RADIO1', key='_RADIO11_', default=True), sg.Radio('Xmind转csv', 'RADIO1', key='_RADIO12_')],
    [sg.Input(key = '_ADDRESS_', font='微软雅黑', size=(20, 1)), sg.FileBrowse('选择文件', font='微软雅黑', size=(8, 1))],
    [sg.Button('文件转换', key = '_CONFIRM_', font='微软雅黑', size=(8, 1)), sg.Button('打开', key='_FOLDER_', size=(8, 1)), sg.Exit('退出', key = '_EXIT_', font='微软雅黑', size=(8, 1))]
]
# 定义窗口，窗口名称
window = sg.Window('文件转换工具',layout,font='微软雅黑')
# 自定义窗口进行数值回显
while True:
    event,values = window.read()
    if event == '_CONFIRM_':
        if values.get('_RADIO11_','True'):
            address = values['_ADDRESS_']
            pdf_conversion = pdf_conversion()
            sg.popup('转换已完成', font='微软雅黑')
        elif values.get('_RADIO12_','True'):
            address = values['_ADDRESS_']
            xmind_conversion = xmind_conversion()
            sg.popup('转换已完成', font='微软雅黑')
        else:
            pass
    elif event == '_FOLDER_':
        os.startfile(r"D:\\@@@")
    elif event in ['_EXIT_',None]:
        break
    else:
        pass