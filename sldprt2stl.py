import os
import re
import win32com
from pythoncom import Nothing, VT_BYREF, VT_I4
from win32com.client import VARIANT
from swconst import constants


seq = []


Errors=VARIANT(VT_BYREF | VT_I4, -1)
Warnings=VARIANT(VT_BYREF | VT_I4, -1)


def traverse_features(swApp, Part, thisFeat, isTopLevel, seq, traversed_features, active_sketch_name, id, pth, out_dir):
    curFeat = thisFeat
    while curFeat is not None:
        name = curFeat.Name
        typeName = curFeat.GetTypeName
        print(curFeat.Name, curFeat.GetTypeName)

        if name not in [feat[1] for feat in traversed_features] and typeName == "ProfileFeature":
            feat_type_name = "SKETCH"
            Part.Extension.SelectByID2(name, "SKETCH", 0, 0, 0, False, 0, Nothing, 0)
            id += 1

            Part.SketchManager.InsertSketch(True)
            sketch = Part.GetActiveSketch2
            if sketch is not None:
                sketchSegments = sketch.GetSketchSegments
                if sketchSegments is not None:
                    for seg in sketchSegments:
                        # filename = os.path.join(out_dir, '.'.join(os.path.basename(pth).split('.')[:-1]) + '_{}_{}'.format(curFeat.Name, seg.Name) + '.stl')
                        # Part.Extension.SaveAs(filename, constants.swSaveAsCurrentVersion, constants.swSaveAsOptions_Silent, Nothing, Errors, Warnings)
                        
                        data = [seg.ConstructionGeometry]
                        data += [0] * (28 - len(data))
                        
                        label = input('请输入标签：')
                        traversed_features.append((id, curFeat.Name, "SKETCHSEGMENT", seg.GetName, data, label))
        elif all(s not in curFeat.GetTypeName for s in ("Folder", "AnnotationViewFeat", "DetailCabinet", "Light", "OriginProfileFeature")):
            feat_type_name = 'BODYFEATURE'
            id += 1
            swExtrusion = curFeat.GetDefinition
            if curFeat.GetTypeName in ('Boss', 'Extrusion'):
                data = [True, swExtrusion.FlipSideToCut
                        , swExtrusion.ReverseDirection
                        , swExtrusion.GetEndCondition(True)
                        , swExtrusion.GetEndCondition(False)
                        , swExtrusion.GetDepth(True)
                        , swExtrusion.GetDepth(False)
                        , swExtrusion.GetDraftWhileExtruding(True)
                        , swExtrusion.GetDraftWhileExtruding(False)
                        , not swExtrusion.GetDraftOutward(True)
                        , not swExtrusion.GetDraftOutward(False)
                        , swExtrusion.GetDraftAngle(True)
                        , swExtrusion.GetDraftAngle(False)
                        , swExtrusion.GetReverseOffset(True)
                        , swExtrusion.GetReverseOffset(False)
                        , swExtrusion.GetTranslateSurface(True)
                        , swExtrusion.GetTranslateSurface(False)
                        , swExtrusion.Merge
                        , swExtrusion.FeatureScope
                        , swExtrusion.AutoSelect
                        , 0, 0, False]
            elif typeName == 'Cut':
                data = [True, swExtrusion.FlipSideToCut
                        , swExtrusion.ReverseDirection
                        , swExtrusion.GetEndCondition(True)
                        , swExtrusion.GetEndCondition(False)
                        , swExtrusion.GetDepth(True)
                        , swExtrusion.GetDepth(False)
                        , swExtrusion.GetDraftWhileExtruding(True)
                        , swExtrusion.GetDraftWhileExtruding(False)
                        , not swExtrusion.GetDraftOutward(True)
                        , not swExtrusion.GetDraftOutward(False)
                        , swExtrusion.GetDraftAngle(True)
                        , swExtrusion.GetDraftAngle(False)
                        , swExtrusion.GetReverseOffset(True)
                        , swExtrusion.GetReverseOffset(False)
                        , swExtrusion.GetTranslateSurface(True)
                        , swExtrusion.GetTranslateSurface(False)
                        , swExtrusion.NormalCut
                        , swExtrusion.FeatureScope
                        , swExtrusion.AutoSelect
                        , swExtrusion.AssemblyFeatureScope
                        , swExtrusion.AutoSelectComponents
                        , swExtrusion.PropagateFeatureToParts
                        , 0, 0, False, swExtrusion.OptimizeGeometry]
            else:
                data = []
            data += [0] * (28 - len(data))
            label = input('请输入标签：')
            traversed_features.append((id, curFeat.Name, feat_type_name, curFeat.GetTypeName, data, label))

            


        subfeat = curFeat.GetFirstSubFeature

        while subfeat is not None:
            traverse_features(swApp, Part, subfeat, False, seq, traversed_features, active_sketch_name, id, pth, out_dir)
            nextSubFeat = subfeat.GetNextSubFeature
            subfeat = nextSubFeat
            nextSubFeat = None

        subfeat = None


        if isTopLevel:
            nextFeat = curFeat.GetNextFeature
        else:
            nextFeat = None
            
        
        curFeat = nextFeat
        nextFeat = None


def get_files(directory, ext, with_dot=True):
    path_list = []
    for paths in [[os.path.join(dirpath, name).replace('/', '\\') for name in filenames if
                   name.endswith(('.' if with_dot else '') + ext)] for
                  dirpath, dirnames, filenames in os.walk(directory)]:
        path_list.extend(paths)
    return path_list



import numpy as np
 
 
 
def stl_get(stl_path):
    points=[]
    f = open(stl_path)
    lines = f.readlines()
    prefix='vertex'
    num=3
    for line in lines:
        #print (line)
 
        if line.strip().startswith(prefix):
            values = line.strip().split()
            #print(values[1:4])
            if num%3==0:
                points.append([values[1],values[2],values[3]])
                num=0
            num+=1
        #print(type(line))
    points=np.array(points,dtype='float64')
    f.close()

    normals = []
    for i, p0 in enumerate(points):
        v1 = points[(i+1) % len(points)] - p0
        v2 = points[(i+2) % len(points)] - p0
        normals.append(np.cross(v1, v2))

    print(points.shape, len(normals))
  
    # # 将点云和法向量保存为OBJ格式文件
    # with open(r'F:\datasets\SW数据集\提取数据\20230415_05\point_f_assembly.obj', 'w') as f:
    #     for i, p in enumerate(points):
    #         f.write(f'v {p[0]} {p[1]} {p[2]}\n')
    #         n = normals[i]
    #         f.write(f'vn {n[0]} {n[1]} {n[2]}\n')
    
    #     # 将点云转换为面信息
    #     for i in range(0, len(points), 3):
    #         f.write(f'f {i+1}//{i+1} {i+2}//{i+2} {i+3}//{i+3}\n')
    return points, normals
 
# import open3d as o3d
# stl_path=r"F:\datasets\SW数据集\提取数据\20230415_05\111_3.STL"
# stl_get(stl_path)
# mesh = o3d.io.read_triangle_mesh(r'F:\datasets\SW数据集\提取数据\20230415_05\point_f_assembly.obj')
# o3d.visualization.draw_geometries([mesh], window_name="obj")
# exit()

traversed_features = []

swApp = win32com.client.Dispatch("SldWorks.Application")
Part = swApp.ActiveDoc

out_dir = r""

files = get_files(r'F:\机械模型库\SW', 'sldprt')
files += get_files(r'F:\机械模型库\SW', 'SLDPRT')
print(files)
out_dir = r'F:\datasets\SW数据集\提取数据\20230415_05'
if not os.path.exists(out_dir):
    os.mkdir(out_dir)
for i, filename in enumerate(files):
    traversed_features.clear()
    swApp.OpenDoc6(filename,1,1,"",Errors,Warnings)
    Part = swApp.ActiveDoc
    try:
        traverse_features(swApp, Part, Part.FirstFeature, True, [], traversed_features, '', 0, filename, out_dir)
        # for i in range(len(traversed_features) - 1):
        for id, name, feat_type_name, sub_type_name, data, label in traversed_features[::-1]:
            print(id, name, feat_type_name, sub_type_name, label)
            if feat_type_name == 'SKETCHSEGMENT':
                Part.Extension.SelectByID2(feat_type_name, "SKETCH", 0, 0, 0, True, 0, Nothing, 0)
                Part.EditSketch
                Part.Extension.SelectByID2(sub_type_name, feat_type_name, 0, 0, 0, True, 0, Nothing, 0)
                # pass
            else:
                Part.EditSketch
                Part.SketchManager.InsertSketch(True)
                Part.Extension.SelectByID2(name, feat_type_name, 0, 0, 0, True, 0, Nothing, 0)
                Part.Extension.DeleteSelection2(constants.swDelete_Absorbed)
                # Part.EditDelete
                pth = os.path.join(out_dir, '.'.join(os.path.basename(filename).split('.')[:-1]) + '_{}'.format(id) + '.stl')
                boolstatus = Part.Extension.SaveAs(pth, constants.swSaveAsCurrentVersion, constants.swSaveAsOptions_CopyAndOpen, Nothing, Errors, Warnings)
                # pth = os.path.join(out_dir, '.'.join(os.path.basename(filename).split('.')[:-1]) + '_{}'.format(id) + '.obj')
                # Part.Extension.SaveAs(pth, constants.swSaveAsCurrentVersion, constants.swSaveAsOptions_Silent, Nothing, Errors, Warnings)
                # Part.EditUndo2
                # pth = os.path.join(out_dir, '.'.join(os.path.basename(filename).split('.')[:-1]) + '_{}_1'.format(id) + '.stl')
                # Part.Extension.SaveAs(pth, constants.swSaveAsCurrentVersion, constants.swSaveAsOptions_Silent, Nothing, Errors, Warnings)
                if not boolstatus:
                    continue
            with open(os.path.join(out_dir, '.'.join(os.path.basename(filename).split('.')[:-1]) + '.txt'), 'a', encoding='utf-8') as f:
                f.write(
                    label + '\t' + 
                    '.'.join(os.path.basename(filename).split('.')[:-1]) + '_{}'.format(id) + '.stl' + '\t' + 
                        sub_type_name + '\t' + str(data) + '\n')
    except Exception as e:
        print(e)
    swApp.CloseDoc("")
