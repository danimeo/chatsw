import re
import os


def get_files(directory, ext, with_dot=True):
    path_list = []
    for paths in [[os.path.join(dirpath, name).replace('/', '\\') for name in filenames if
                   name.endswith(('.' if with_dot else '') + ext)] for
                  dirpath, dirnames, filenames in os.walk(directory)]:
        path_list.extend(paths)
    return path_list


def convert(text):
    text = re.sub('Set swApp = Application.SldWorks', '', text)
    text = re.sub(r'myModelView\.FrameState \= [\.\w]+\n', '', text)
    text = re.sub('Sub \w+(\w*)', '', text)
    text = re.sub('Set\s+(\w+)\s+=\s+(\w+)', r'\1 = \2', text)
    text = re.sub(r'Dim \w+ As \w+', '', text)
    text = re.sub('\' .+\n', '', text)

    text = re.sub(r'([^=()\n ]+) ([^=\n()]+)\n', r'\1(\2)\n', text)
    text = re.sub(r'(\n[^=()\n ]+)\n', r'\1()\n', text)
    text = re.sub('([0-9]+)#', r'\1', text)
    # text = re.sub(r'(myDimension.SystemValue = .+)\n', r'\1\npyautogui.hotkey("enter")\n', text)

    text = '''
import win32com.client
from pythoncom import Nothing

swApp = win32com.client.Dispatch("SldWorks.Application")
swApp.Visible = True
''' + text
    return text


vba_dir = r'F:\datasets\SW数据集\宏（自己录制）\20230408_01'
output_dir = r'F:\datasets\SW数据集\宏（自己录制）\20230408_01\py'
vba_files = get_files(vba_dir, 'bas')

for vba_file in vba_files:
    with open(vba_file, 'rt') as f:
        text = f.read()
    text = convert(text)
    with open(os.path.join(output_dir, '.'.join(os.path.basename(vba_file).split('.')[:-1]) + '.py'), 'wt', encoding='utf-8') as f:
        f.write(text)
