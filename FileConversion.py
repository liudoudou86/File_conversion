# -*- coding:utf-8 -*-

import os
import shutil
import time

import PySimpleGUI as sg
from pdf2docx import Converter
from docx2pdf import convert
from xmind2testcase.zentao import xmind_to_zentao_csv_file

# 获取当前时间
def get_current_time():
    time_tup = time.localtime(time.time())
    format_time = "%Y-%m-%d_%H%M%S"
    return time.strftime(format_time, time_tup)

# 确保输出目录存在
def ensure_output_dir(output_dir=None):
    if output_dir is None:
        # 使用用户文档目录作为默认位置
        user_docs = os.path.join(os.path.expanduser('~'), 'Documents')
        output_dir = os.path.join(user_docs, 'FileOutput')
    
    try:
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
        return output_dir
    except Exception as e:
        # 如果创建目录失败，回退到当前目录
        fallback_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'output')
        if not os.path.exists(fallback_dir):
            os.makedirs(fallback_dir)
        sg.popup_error(f'无法创建指定目录: {str(e)}\n已使用备用目录: {fallback_dir}', font='微软雅黑')
        return fallback_dir

def pdf_to_word(address):
    try:
        output_dir = ensure_output_dir()
        cur_time = get_current_time()
        wordfile_path = os.path.join(output_dir, f'wordfile_{cur_time}.docx')

        cv = Converter(address)
        cv.convert(wordfile_path, start=0, end=None)
        cv.close()
        return True
    except Exception as e:
        sg.popup_error(f'PDF转Word失败: {str(e)}', font='微软雅黑')
        return False

def word_to_pdf(address):
    try:
        output_dir = ensure_output_dir()
        cur_time = get_current_time()
        pdffile_path = os.path.join(output_dir, f'pdffile_{cur_time}.pdf')

        convert(address, pdffile_path)
        return True
    except Exception as e:
        sg.popup_error(f'Word转PDF失败: {str(e)}', font='微软雅黑')
        return False

def xmind_conversion(address):
    try:
        output_dir = ensure_output_dir()
        cur_time = get_current_time()

        old_csvfile = xmind_to_zentao_csv_file(address)
        old_csvfile_path = os.path.split(old_csvfile)
        new_csvfile = os.path.join(old_csvfile_path[0], f'csvfile_{cur_time}.csv')
        os.rename(old_csvfile, new_csvfile)
        shutil.move(new_csvfile, output_dir)
        return True
    except Exception as e:
        sg.popup_error(f'Xmind转CSV失败: {str(e)}', font='微软雅黑')
        return False

def main():
    layout = [
        [sg.Radio('PDF转word', 'RADIO1', key='_RADIO11_', default=True), 
         sg.Radio('Word转pdf', 'RADIO1', key='_RADIO12_')],
        [sg.Radio('Xmind转csv', 'RADIO1', key='_RADIO13_')],
        [sg.Input(key='_ADDRESS_', font='微软雅黑', size=(20, 1)), 
         sg.FileBrowse('选择文件', font='微软雅黑', size=(8, 1))],
        [sg.Button('文件转换', key='_CONFIRM_', font='微软雅黑', size=(8, 1)),
         sg.Button('打开目录', key='_FOLDER_', size=(8, 1)), 
         sg.Exit('退出', key='_EXIT_', font='微软雅黑', size=(8, 1))]
    ]
    
    # 定义窗口，窗口名称
    window = sg.Window('文件转换工具', layout, font='微软雅黑')
    
    # 自定义窗口进行数值回显
    while True:
        event, values = window.read()
        
        if event == '_CONFIRM_':
            address = values.get('_ADDRESS_', '')
            
            if not address:
                sg.popup_error('请选择文件', font='微软雅黑')
                continue
                
            if values.get('_RADIO11_'):
                if pdf_to_word(address):
                    sg.popup('转换已完成', font='微软雅黑')
            elif values.get('_RADIO12_'):
                if word_to_pdf(address):
                    sg.popup('转换已完成', font='微软雅黑')
            elif values.get('_RADIO13_'):
                if xmind_conversion(address):
                    sg.popup('转换已完成', font='微软雅黑')
        
        elif event == '_FOLDER_':
            output_dir = ensure_output_dir()
            os.startfile(output_dir)
        
        elif event in ['_EXIT_', None]:
            break

if __name__ == "__main__":
    main()
