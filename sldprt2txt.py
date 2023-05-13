import os
import win32com
from pythoncom import Nothing, VT_BYREF, VT_I4
from win32com.client import VARIANT
from swconst import constants

seg_type_list = [
    constants.swSketchLINE, 
    constants.swSketchARC, 
    constants.swSketchELLIPSE, 
    constants.swSketchSPLINE, 
    constants.swSketchTEXT, 
    constants.swSketchPARABOLA,
]
feat_type_list = ["ProfileFeature", 'Boss/Extrusion', 'Cut', 'HoleWzd']
type_list = seg_type_list + feat_type_list[1:]
seq = []

def traverse_features_and_save(swApp, Part, thisFeat, isTopLevel, traversed_features, active_sketch_name, id, save_txt_name, out_dir):
    curFeat = thisFeat
    while curFeat is not None:
        name = curFeat.Name
        typeName = curFeat.GetTypeName
        print(curFeat.Name, curFeat.GetTypeName)

        if typeName in feat_type_list:
            id += 1
            if typeName in ("ProfileFeature"):
                Part.Extension.SelectByID2(name, "SKETCH", 0, 0, 0, False, 0, Nothing, 0)

                Part.SketchManager.InsertSketch(True)
                sketch = Part.GetActiveSketch2
                if sketch is not None:
                    
                    v = VARIANT(VT_BYREF | VT_I4, 0)
                    ref = sketch.GetReferenceEntity(v)
                    ref.Select2(False, 0)
                    pFeat = Part.SelectionManager.GetSelectedObject5(1)
                    
                    if name not in traversed_features and pFeat is not None:
                        if v.value in (constants.swReferenceTypeFace, constants.swReferenceTypeEdge, constants.swReferenceTypeBody):
                            print('参考面：', pFeat.GetFeature.Name, pFeat.Normal)
                        else:
                            print('参考面：', pFeat.Name)
                            if not pFeat.Name:
                                continue
                    
                    sketchSegments = sketch.GetSketchSegments
                    if sketchSegments is not None:
                        # points = []
                        # for sketchSegment in sketchSegments:
                        #     if sketchSegment.GetType in seg_type_list:
                        #         p0 = (sketchSegment.GetStartPoint2.X, sketchSegment.GetStartPoint2.Y)
                        #         if p0 not in points:
                        #             points.append(p0)
                        #         p1 = (sketchSegment.GetEndPoint2.X, sketchSegment.GetEndPoint2.Y)
                        #         if p1 not in points:
                        #             points.append(p1)
                            
                        # seq.append(str(points))
                        # for sketchSegment in sketchSegments:
                        #     if sketchSegment.GetType in seg_type_list:
                        #         start_point = (sketchSegment.GetStartPoint2.X, sketchSegment.GetStartPoint2.Y)
                        #         end_point = (sketchSegment.GetEndPoint2.X, sketchSegment.GetEndPoint2.Y)
                        #         data = [
                        #             seg_type_list.index(sketchSegment.GetType),
                        #             int(sketchSegment.ConstructionGeometry),
                        #             points.index(start_point) if start_point in points else -1,
                        #             points.index(end_point) if end_point in points else -1,
                        #             sketchSegment.GetCenterPoint2.X if seg_type_list.index(sketchSegment.GetType) == 1 else 0,
                        #             sketchSegment.GetCenterPoint2.Y if seg_type_list.index(sketchSegment.GetType) == 1 else 0,
                        #         ]
                        #         seq.append('{},{},{}'.format(id, ','.join([str(d) for d in data[:4]]), ','.join(['{:.6f}'.format(d) for d in data[4:6]])))
                        
                        for sketchSegment in sketchSegments:
                            if sketchSegment.GetType in seg_type_list:
                                start_point = (sketchSegment.GetStartPoint2.X, sketchSegment.GetStartPoint2.Y)
                                end_point = (sketchSegment.GetEndPoint2.X, sketchSegment.GetEndPoint2.Y)
                                data = [
                                    seg_type_list.index(sketchSegment.GetType),
                                    int(start_point[0] * 1e4),
                                    int(start_point[1] * 1e4),
                                    int(end_point[0] * 1e4),
                                    int(end_point[1] * 1e4),
                                    int(sketchSegment.GetCenterPoint2.X * 1e4) if seg_type_list.index(sketchSegment.GetType) == 1 else 0,
                                    int(sketchSegment.GetCenterPoint2.Y * 1e4) if seg_type_list.index(sketchSegment.GetType) == 1 else 0,
                                    int(sketchSegment.ConstructionGeometry),
                                ]
                                data += [0] * (28 - len(data))
                                seq.append('{},{}'.format(id, ','.join([str(int(d * 1e4)) if isinstance(d, float) else str(int(d)) for d in data])))
            elif typeName in ('Boss', 'Extrusion'):
                swExtrusion = curFeat.GetDefinition
                if swExtrusion is not None:
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
                    data += [0] * (28 - len(data))
                    seq.append('{},{},{}'.format(id, type_list.index('Boss/Extrusion'), ','.join([str(int(d * 1e4)) if isinstance(d, float) else str(int(d)) for d in data])))
            elif typeName == 'Cut':
                swExtrusion = curFeat.GetDefinition
                if swExtrusion is not None:
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
                    data += [0] * (28 - len(data))
                    seq.append('{},{},{}'.format(id, type_list.index('Cut'), ','.join([str(int(d * 1e4)) if isinstance(d, float) else str(int(d)) for d in data])))
            elif typeName == 'HoleWzd':
                swWizHole = curFeat.GetDefinition
                if swWizHole.Type in (10, 11, 12, 13, 2, 7, 57, 58, 59, 60, 61, 62, 63, 64, 65, 66, 67, 68, 14, 15, 16, 17, 18, 19, 20, 21, 4, 9):
                    data = [
                        swWizHole.CounterBoreDiameter, 
                        swWizHole.CounterBoreDepth,
                        swWizHole.HeadClearance,
                        swWizHole.HoleFit,
                        swWizHole.CounterDrillAngle,  # ?
                        swWizHole.NearCounterSinkDiameter,
                        swWizHole.NearCounterSinkAngle,
                        swWizHole.CounterSinkDiameter,  # ?
                        swWizHole.CounterSinkAngle,  # ?
                        swWizHole.FarCounterSinkDiameter,
                        swWizHole.FarCounterSinkAngle,
                        swWizHole.OffsetDistance,
                    ]
                elif swWizHole.Type in (24, 43, 71, 78, 76, 77, 90, 79, 29, 30, 45, 44, 3, 8, ):
                    data = [
                        swWizHole.NearCounterSinkDiameter,
                        swWizHole.NearCounterSinkAngle,
                        swWizHole.HeadClearance,
                        swWizHole.HoleFit,
                        swWizHole.CounterDrillAngle,  # ?
                        swWizHole.FarCounterSinkDiameter,
                        swWizHole.FarCounterSinkAngle,
                        swWizHole.OffsetDistance,
                        swWizHole.HeadClearanceType,
                        -1,
                        -1,
                        -1,
                    ]
                elif swWizHole.Type in (22, 23, 25, 26, 27, 28):
                    data = [
                        swWizHole.HoleFit,
                        swWizHole.CounterDrillAngle,  # ?
                        swWizHole.NearCounterSinkDiameter,
                        swWizHole.NearCounterSinkAngle,
                        swWizHole.FarCounterSinkDiameter,
                        swWizHole.FarCounterSinkAngle,
                        swWizHole.OffsetDistance,
                        -1,
                        -1,
                        -1,
                        -1,
                        -1,
                    ]
                else:
                    data = [
                        swWizHole.CounterBoreDiameter, 
                        swWizHole.CounterBoreDepth,
                        swWizHole.HeadClearance,
                    ]
                data += [0] * (28 - len(data))
                seq.append('{},{},{}'.format(id, type_list.index('HoleWzd'), ','.join([str(int(d * 1e4)) if isinstance(d, float) else str(int(d)) for d in data])))
                # seq.append(f'hole_wizard({swWizHole.Type}, {swWizHole.Standard2}, {swWizHole.FastenerType2}, "{swWizHole.FastenerSize}", {swWizHole.EndCondition}, {swWizHole.Diameter}, {swWizHole.Depth}, {swWizHole.Length}, {", ".join([str(d) for d in data])}, "{swWizHole.ThreadClass}", {swWizHole.ReverseDirection}, {swWizHole.FeatureScope}, {swWizHole.AutoSelect}, {swWizHole.AssemblyFeatureScope}, {swWizHole.AutoSelectComponents}, {swWizHole.PropagateFeatureToParts})')
        

            active_sketch_name = curFeat.Name
        
        subfeat = curFeat.GetFirstSubFeature

        while subfeat is not None:
            traverse_features_and_save(swApp, Part, subfeat, False, traversed_features, active_sketch_name, id, save_txt_name, out_dir)
            nextSubFeat = subfeat.GetNextSubFeature
            subfeat = nextSubFeat
            nextSubFeat = None

        subfeat = None


        if isTopLevel:
            nextFeat = curFeat.GetNextFeature
        else:
            nextFeat = None

        traversed_features.append(curFeat.Name)
        curFeat = nextFeat
        nextFeat = None


def get_files(directory, ext, with_dot=True):
    path_list = []
    for paths in [[os.path.join(dirpath, name).replace('/', '\\') for name in filenames if
                   name.endswith(('.' if with_dot else '') + ext)] for
                  dirpath, dirnames, filenames in os.walk(directory)]:
        path_list.extend(paths)
    return path_list


swApp = win32com.client.Dispatch("SldWorks.Application")
Errors=VARIANT(VT_BYREF | VT_I4, -1)
Warnings=VARIANT(VT_BYREF | VT_I4, -1)

files = get_files(r'F:\机械模型库\SW', 'sldprt')
files += get_files(r'F:\机械模型库\SW', 'SLDPRT')
print(files)
out_dir = r'F:\datasets\SW数据集\提取数据\20230415_01'
if not os.path.exists(out_dir):
    os.mkdir(out_dir)
for filename in files:
    swApp.OpenDoc6(filename,1,1,"",Errors,Warnings)
    Part = swApp.ActiveDoc
    try:
        seq.clear()
        # seq.append('.'.join(os.path.basename(filename).split('.')[:-1]))
        seq.append(','.join(['id', 'type',] + ['p' + str(i) for i in range(1, 27 + 1)]))
        traverse_features_and_save(swApp, Part, Part.FirstFeature, True, [], '', 0, filename, out_dir)
        with open(os.path.join(out_dir, '.'.join(os.path.basename(filename).split('.')[:-1]) + '.txt'), 'wt', encoding='utf-8') as f:
            f.write('\n'.join(seq))
    except Exception as e:
        print(e)
    swApp.CloseDoc("")
