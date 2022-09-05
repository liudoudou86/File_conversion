# -*- coding:utf-8 -*-


import os
import time
from typing import Sized
from PySimpleGUI.PySimpleGUI import TRANSPARENT_BUTTON
from pdf2docx import Converter
import PySimpleGUI as sg


def file_conversion():

    isExists=os.path.exists('D:\###')
    if not isExists:
        os.mkdir('D:\###')
    else:
        pass

    time_tup=time.localtime(time.time()) # 获取当前时间
    format_time="%Y-%m-%d_%H-%M-%S"
    cur_time=time.strftime(format_time,time_tup)

    pdf_file = address
    docx_file = 'D://###//newfile_{}.docx'.format(cur_time)

    cv = Converter(pdf_file)
    cv.convert(docx_file, start=0, end=None)
    cv.close()

layout = [
    [sg.Input(key = '_ADDRESS_', font='微软雅黑', size=(20, 1)), sg.FileBrowse('选择文件', font='微软雅黑', size=(8, 1))],
    [sg.Button('文件转换', key = '_CONFIRM_', font='微软雅黑', size=(8, 1)), sg.Button('打开', key='_FOLDER_', size=(8, 1)), sg.Exit('退出', key = '_EXIT_', font='微软雅黑', size=(8, 1))]
]
# 定义窗口，窗口名称
window = sg.Window('PDF转Word工具',layout,font='微软雅黑')
# 自定义窗口进行数值回显
while True:
    event,values = window.read()
    if event == '_CONFIRM_':
        address = values['_ADDRESS_']
        file_conversion = file_conversion()
        sg.popup('转换已完成', font='微软雅黑')
    elif event == '_FOLDER_':
        os.startfile(r"D:\###")
    elif event in ['_EXIT_',None]:
        break
    else:
        pass