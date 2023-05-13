import os
import csv
import pandas as pd

import win32com.client
from pythoncom import Nothing, VT_BYREF, VT_I4
from win32com.client import VARIANT
from swconst import constants

swApp = win32com.client.Dispatch("SldWorks.Application")
swApp.Visible = True
Part = swApp.ActiveDoc

seg_type_list = [
    constants.swSketchLINE, 
    constants.swSketchARC, 
    constants.swSketchELLIPSE, 
    constants.swSketchSPLINE, 
    constants.swSketchTEXT, 
    constants.swSketchPARABOLA,
]
feat_type_list = ["ProfileFeature", 'Boss/Extrusion', 'Cut']
type_list = seg_type_list + feat_type_list[1:]


def get_files(directory, ext, with_dot=True):
    path_list = []
    for paths in [[os.path.join(dirpath, name).replace('/', '\\') for name in filenames if
                   name.endswith(('.' if with_dot else '') + ext)] for
                  dirpath, dirnames, filenames in os.walk(directory)]:
        path_list.extend(paths)
    return path_list


def preprocess(root):
    paths = get_files(root, 'txt')

    for pth in paths:
        last_id = 0
        with open(pth, 'r', encoding='utf-8') as f:
            reader = csv.reader(f, )
            next(reader)
            for row in reader:
                id = int(row[0])
                typ = int(row[1])
                print(row)
                if typ == 0:
                    if last_id != id:
                        Part.SketchManager.InsertSketch(True)
                    data = [float(d) * 1e-4 for d in (row[2], row[3], 0, row[4], row[5], 0, )]
                    Part.SketchManager.CreateLine(*data)
                    if last_id != id:
                        Part.SketchManager.InsertSketch(True)
                elif typ == 1:
                    if last_id != id:
                        Part.SketchManager.InsertSketch(True)
                    data = [float(d) * 1e-4 for d in (row[6], row[7], 0, row[2], row[3], 0, row[4], row[5], 0, row[8])]
                    Part.SketchManager.CreateArc(*data)
                    if last_id != id:
                        Part.SketchManager.InsertSketch(True)
                elif typ == type_list.index("Boss/Extrusion"):
                    if last_id != id and last_id < 7:
                        Part.SketchManager.InsertSketch(True)
                    data = [float(d) * 1e-4 if isinstance(d, float) else int(d) for d in row[2:12]]
                    Part.FeatureManager.FeatureExtrusion3(*row[2:25])
                    if last_id != id and last_id < 7:
                        Part.SketchManager.InsertSketch(True)
                elif typ == type_list.index("Cut"):
                    if last_id != id and last_id < 7:
                        Part.SketchManager.InsertSketch(True)
                        data = [float(d) * 1e-4 if isinstance(d, float) else int(d) for d in row[2:29]]
                    Part.FeatureManager.FeatureCut4(*row[2:29])
                    if last_id != id and last_id < 7:
                        Part.SketchManager.InsertSketch(True)

                last_id = id


preprocess(r'F:\datasets\SW数据集\提取数据\20230415_01')
