import os
import re
import win32com
from pythoncom import Nothing, VT_BYREF, VT_I4
from win32com.client import VARIANT
from swconst import constants


def initialize_seq():
    seq = [
        'import win32com.client',
        'from pythoncom import Nothing, VT_BYREF, VT_I4',
        'from win32com.client import VARIANT',
        'from swconst import constants',
        '',
        'swApp = win32com.client.Dispatch("SldWorks.Application")',
        'swApp.Visible = True',
        'Part = swApp.ActiveDoc',
        'line = Part.SketchManager.CreateLine',
        'circle = Part.SketchManager.CreateCircle',
        'arc = Part.SketchManager.CreateArc',
        'enter_sketch = lambda: Part.SketchManager.InsertSketch(True)',
        'exit_sketch = Part.SketchManager.InsertSketch',
        'get_features = Part.FeatureManager.GetFeatures',
        'select = Part.Extension.SelectByID2',
        'extrude = Part.FeatureManager.FeatureExtrusion3',
        'hole_wizard = Part.FeatureManager.HoleWizard5',
        'cut = Part.FeatureManager.FeatureCut4',
        'feat_name = ""',
        '',
    ]

    seq.append('''
def select_face(feature_name, normal):
    features = get_features(True)
    for pFeat in features[::-1]:
        if pFeat is not None and pFeat.Name == feature_name:
            for face in pFeat.GetFaces:
                if face.Normal == normal:
                    face.Select2(False, 0)
                    break
''')

    seq.append('''
def select_sketch():
    name = [feat.Name for feat in get_features(True) if feat.GetTypeName == "ProfileFeature"][-1]
    select(name, "SKETCH", 0, 0, 0, False, 0, Nothing, 0)
''')

    seq.append('''
def get_ref_feature_name():
    sketch = Part.GetActiveSketch2
    v = VARIANT(VT_BYREF | VT_I4, 0)
    ref = sketch.GetReferenceEntity(v)
    ref.Select2(False, 0)
    pFeat = Part.SelectionManager.GetSelectedObject5(1)
    if v.value in (constants.swReferenceTypeFace, constants.swReferenceTypeEdge, constants.swReferenceTypeBody):
        return pFeat.GetFeature.Name
    else:
        return pFeat.Name
''')
    return seq
            

def traverse_features(swApp, Part, thisFeat, isTopLevel, seq, traversed_features, active_sketch_name, pFaceID):
    curFeat = thisFeat
    while curFeat is not None:
        name = curFeat.Name
        typeName = curFeat.GetTypeName
        print(curFeat.Name, curFeat.GetTypeName)

        if typeName == "ProfileFeature":
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
                        pFaceID += 1
                        print('参考面：', pFeat.GetFeature.Name, pFeat.Normal)
                        seq.append('select_face(feat.Name, {})'.format(str(pFeat.Normal)))
                    else:
                        print('参考面：', pFeat.Name)
                        if not pFeat.Name:
                            continue
                        seq.append('select("{}", "PLANE", 0, 0, 0, False, 0, Nothing, 0)'.format(pFeat.Name))

                seq.append('enter_sketch()')
                
                sketchSegments = sketch.GetSketchSegments
                if sketchSegments is not None:
                    c2 = 'rels = ['
                    relations = []
                    for sketchSegment in sketchSegments:
                        if sketchSegment.GetType == constants.swSketchLINE:
                            sketchPointStart = sketchSegment.GetStartPoint2
                            sketchPointEnd = sketchSegment.GetEndPoint2

                            if sketchPointStart.X == sketchPointEnd.X and sketchPointStart.Y == sketchPointEnd.Y:
                                continue

                            data = [
                                sketchPointStart.X,
                                sketchPointStart.Y,
                                sketchPointStart.Z,
                                sketchPointEnd.X,
                                sketchPointEnd.Y,
                                sketchPointEnd.Z,
                            ]

                            seq.append('l = line({})'.format(','.join(['{}'.format(d) for d in data])))
                            if sketchSegment.ConstructionGeometry:
                                seq.append('l.ConstructionGeometry = True')
                        elif sketchSegment.GetType == constants.swSketchARC:
                            sketchPointStart = sketchSegment.GetStartPoint2
                            sketchPointEnd = sketchSegment.GetEndPoint2
                            sketchPointCenter = sketchSegment.GetCenterPoint2

                            if sketchPointStart.X == sketchPointEnd.X and sketchPointStart.Y == sketchPointEnd.Y:
                                data = [
                                    sketchPointCenter.X,
                                    sketchPointCenter.Y,
                                    sketchPointCenter.Z,
                                    sketchPointStart.X,
                                    sketchPointStart.Y,
                                    sketchPointStart.Z,
                                ]
                                seq.append('circle({})'.format(','.join(['{}'.format(d) for d in data])))
                            else:
                                data = [
                                    sketchPointCenter.X,
                                    sketchPointCenter.Y,
                                    sketchPointCenter.Z,
                                    sketchPointStart.X,
                                    sketchPointStart.Y,
                                    sketchPointStart.Z,
                                    sketchPointEnd.X,
                                    sketchPointEnd.Y,
                                    sketchPointEnd.Z,
                                    1
                                ]
                                seq.append('arc({})'.format(','.join(['{}'.format(d) for d in data])))

                        rels = sketchSegment.GetRelations
                        if rels is not None:
                            for rel in rels:
                                try:
                                    names = ['"{}"'.format(ent.GetName) for ent in rel.GetEntities]
                                    rel_type = rel.GetRelationType
                                    if rel_type == constants.swConstraintType_HORIZONTAL:
                                        rel_type = '"sgHORIZONTAL2D"'
                                    elif rel_type == constants.swConstraintType_ALONGX3D:
                                        rel_type = '"sgALONGX3D"'
                                    elif rel_type == constants.swConstraintType_HORIZPOINTS:
                                        rel_type = '"sgHORIZONTALPOINTS2D"'
                                    elif rel_type == constants.swConstraintType_ALONGXPOINTS3D:
                                        rel_type = '"sgALONGXPOINTS3D"'
                                    elif rel_type == constants.swConstraintType_VERTICAL:
                                        rel_type = '"sgVERTICAL2D"'
                                    elif rel_type == constants.swConstraintType_ALONGY3D:
                                        rel_type = '"sgALONGY3D"'
                                    elif rel_type == constants.swConstraintType_VERTPOINTS:
                                        rel_type = '"sgVERTPOINTS2D"'
                                    elif rel_type == constants.swConstraintType_ALONGYPOINTS3D:
                                        rel_type = '"sgALONGYPOINTS3D"'
                                    elif rel_type == constants.swConstraintType_ALONGZPOINTS:
                                        rel_type = '"sgALONGZPOINTS3D"'
                                    elif rel_type == constants.swConstraintType_ALONGZ:
                                        rel_type = '"sgALONGZ3D"'
                                    elif rel_type == constants.swConstraintType_COLINEAR:
                                        rel_type = '"sgCOLINEAR"'
                                    elif rel_type == constants.swConstraintType_CORADIAL:
                                        rel_type = '"sgCORADIAL"'
                                    elif rel_type == constants.swConstraintType_PERPENDICULAR:
                                        rel_type = '"sgPERPENDICULAR"'
                                    elif rel_type == constants.swConstraintType_PARALLEL:
                                        rel_type = '"sgPARALLEL"'
                                    elif rel_type == constants.swConstraintType_TANGENT:
                                        rel_type = '"sgTANGENT"'
                                    elif rel_type == constants.swConstraintType_CONCENTRIC:
                                        rel_type = '"sgCONCENTRIC"'
                                    elif rel_type == constants.swConstraintType_COINCIDENT:
                                        rel_type = '"sgCOINCIDENT"'
                                    elif rel_type == constants.swConstraintType_SYMMETRIC:
                                        rel_type = '"sgSYMMETRIC"'
                                    elif rel_type == constants.swConstraintType_ATMIDDLE:
                                        rel_type = '"sgATMIDDLE"'
                                    elif rel_type == constants.swConstraintType_ATINTERSECT:
                                        rel_type = '"sgATINTERSECT"'
                                    elif rel_type == constants.swConstraintType_ATPIERCE:
                                        rel_type = '"sgATPIERCE"'
                                    elif rel_type == constants.swConstraintType_FIXED:
                                        rel_type = '"sgFIXED"'
                                    elif rel_type == constants.swConstraintType_ANGLE:
                                        rel_type = '"sgANGLE"'
                                    elif rel_type == constants.swConstraintType_ARCANG180:
                                        rel_type = '"sgARCANG180"'
                                    elif rel_type == constants.swConstraintType_ARCANG270:
                                        rel_type = '"sgARCANG270"'
                                    elif rel_type == constants.swConstraintType_ARCANG90:
                                        rel_type = '"sgARCANG90"'
                                    elif rel_type == constants.swConstraintType_ARCANGBOTTOM:
                                        rel_type = '"sgARCANGBOTTOM"'
                                    elif rel_type == constants.swConstraintType_ARCANGLEFT:
                                        rel_type = '"sgARCANGLEFT"'
                                    elif rel_type == constants.swConstraintType_ARCANGRIGHT:
                                        rel_type = '"sgARCANGRIGHT"'
                                    elif rel_type == constants.swConstraintType_ARCANGTOP:
                                        rel_type = '"sgARCANGTOP"'
                                    elif rel_type == constants.swConstraintType_DIAMETER:
                                        rel_type = '"sgDIAMETER"'
                                    elif rel_type == constants.swConstraintType_DISTANCE:
                                        rel_type = '"sgDISTANCE"'
                                    elif rel_type == constants.swConstraintType_SAMELENGTH:
                                        rel_type = '"sgSAMELENGTH"'
                                    elif rel_type == constants.swConstraintType_OFFSETEDGE:
                                        rel_type = '"sgOFFSETEDGE"'
                                    elif rel_type == constants.swConstraintType_SNAPANGLE:
                                        rel_type = '"sgSNAPANGLE"'
                                    elif rel_type == constants.swConstraintType_SNAPGRID:
                                        rel_type = '"sgSNAPGRID"'
                                    elif rel_type == constants.swConstraintType_SNAPLENGTH:
                                        rel_type = '"sgSNAPLENGTH"'
                                    elif rel_type == constants.swConstraintType_USEEDGE:
                                        rel_type = '"sgUSEEDGE"'
                                    elif rel_type == constants.swConstraintType_MERGEPOINTS:
                                        rel_type = '"sgMERGEPOINTS"'
                                    relation = '({},{})'.format(rel_type, ','.join(names))
                                    if relation not in relations:
                                        relations.append(relation)
                                        c2 += '{},'.format(relation)
                                except AttributeError as e:
                                    print(e)
                    c2 += ']\n'
                    
                    c2 += '''
for rel in rels:
    Part.ClearSelection2(True)
    for ent_name in rel[1:]:
        Part.Extension.SelectByID2(ent_name, "SKETCHSEGMENT", 0, 0, 0, True, 0, Nothing, 0)
    Part.SketchAddConstraints(rel[0])
    '''
                seq.append(c2)
                seq.append('exit_sketch(True)')

            active_sketch_name = curFeat.Name
        elif typeName in ('Boss', 'Extrusion'):
            seq.append('select_sketch()')
            # depth = curFeat.GetDefinition.GetDepth(True)
            # if depth != 0.:
            #     seq.append(f'feat = extrude(True, False, True, constants.swEndCondBlind, 0, {depth}, {depth}, False, False, False, False, 0, 0, False, False, False, False, True, True, True, constants.swStartSketchPlane, 0, False)')
            
            swExtrusion = curFeat.GetDefinition
            if swExtrusion is not None:
                lines = f'''feat = extrude(True, {swExtrusion.FlipSideToCut
}, {swExtrusion.ReverseDirection
}, {swExtrusion.GetEndCondition(True)
}, {swExtrusion.GetEndCondition(False)
}, {swExtrusion.GetDepth(True)
}, {swExtrusion.GetDepth(False)
}, {swExtrusion.GetDraftWhileExtruding(True)
},{swExtrusion.GetDraftWhileExtruding(False)
}, {not swExtrusion.GetDraftOutward(True)
}, {not swExtrusion.GetDraftOutward(False)
}, {swExtrusion.GetDraftAngle(True)
}, {swExtrusion.GetDraftAngle(False)
}, {swExtrusion.GetReverseOffset(True)
}, {swExtrusion.GetReverseOffset(False)
}, {swExtrusion.GetTranslateSurface(True)
}, {swExtrusion.GetTranslateSurface(False)
}, {swExtrusion.Merge
}, {swExtrusion.FeatureScope
}, {swExtrusion.AutoSelect
}, 0, 0, False)'''
                seq += lines.splitlines()
        elif typeName == 'Cut':
            seq.append('select_sketch()')
            # depth = curFeat.GetDefinition.GetDepth(True)
            # if depth != 0.:
            #     seq.append(f'cut(True, False, False, 0, 0, {depth}, {depth}, False, False, False, False, 0, 0, False, False, False, False, False, False, True, True, True, False, 0, 0, False, True)')
            swExtrusion = curFeat.GetDefinition
            if swExtrusion is not None:
                lines = f'''cut(True, {swExtrusion.FlipSideToCut
}, {swExtrusion.ReverseDirection
}, {swExtrusion.GetEndCondition(True)
}, {swExtrusion.GetEndCondition(False)
}, {swExtrusion.GetDepth(True)
}, {swExtrusion.GetDepth(False)
}, {swExtrusion.GetDraftWhileExtruding(True)
}, {swExtrusion.GetDraftWhileExtruding(False)
}, {not swExtrusion.GetDraftOutward(True)
}, {not swExtrusion.GetDraftOutward(False)
}, {swExtrusion.GetDraftAngle(True)
}, {swExtrusion.GetDraftAngle(False)
}, {swExtrusion.GetReverseOffset(True)
}, {swExtrusion.GetReverseOffset(False)
}, {swExtrusion.GetTranslateSurface(True)
}, {swExtrusion.GetTranslateSurface(False)
}, {swExtrusion.NormalCut
}, {swExtrusion.FeatureScope
}, {swExtrusion.AutoSelect
}, {swExtrusion.AssemblyFeatureScope
}, {swExtrusion.AutoSelectComponents
}, {swExtrusion.PropagateFeatureToParts
}, 0, 0, False, {swExtrusion.OptimizeGeometry
})'''
                seq += lines.splitlines()
        elif typeName == 'HoleWzd':
            seq.append('select_sketch()')
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
            
            seq.append(f'hole_wizard({swWizHole.Type}, {swWizHole.Standard2}, {swWizHole.FastenerType2}, "{swWizHole.FastenerSize}", {swWizHole.EndCondition}, {swWizHole.Diameter}, {swWizHole.Depth}, {swWizHole.Length}, {", ".join([str(d) for d in data])}, "{swWizHole.ThreadClass}", {swWizHole.ReverseDirection}, {swWizHole.FeatureScope}, {swWizHole.AutoSelect}, {swWizHole.AssemblyFeatureScope}, {swWizHole.AutoSelectComponents}, {swWizHole.PropagateFeatureToParts})')
        elif typeName in ('SweepCut', 'Sweep'):
            swSweep = curFeat.GetDefinition
            if swSweep is not None:
                lines = f'''swSweep = Part.FeatureManager.CreateDefinition(constants.{'swFmSweepCut' if typeName == 'SweepCut' else 'swFmSweep'})
swSweep.TwistControlType = {swSweep.TwistControlType}
swSweep.PathAlignmentType = {swSweep.PathAlignmentType}
swSweep.AlignWithEndFaces = {swSweep.AlignWithEndFaces}
swSweep.AutoSelect = {swSweep.AutoSelect}
swSweep.StartTangencyType = {swSweep.StartTangencyType}
swSweep.BossFeature = {swSweep.BossFeature}
swSweep.ThinFeature = {swSweep.ThinFeature}
swSweep.ThinWallType = {swSweep.ThinWallType}
swSweep.MaintainTangency = {swSweep.MaintainTangency}
swSweep.Merge = {swSweep.Merge}
swSweep.MergeSmoothFaces = {swSweep.MergeSmoothFaces}
swSweep.AdvancedSmoothing = {swSweep.AdvancedSmoothing}
swSweep.TangentPropagation = {swSweep.TangentPropagation}
swSweep.PropagateFeatureToParts = {swSweep.PropagateFeatureToParts}

Part.Extension.SelectByID2({swSweep.Profile.Name}, "SKETCH", 0, 0, 0, False, 0, Nothing, 0)
pFeat = Part.SelectionManager.GetSelectedObject6(1, -1)
swSweep.Profile = pFeat

Part.Extension.SelectByID2({swSweep.Path.Name}, "SKETCH", 0, 0, 0, False, 0, Nothing, 0)
pFeat = Part.SelectionManager.GetSelectedObject6(1, -1)
swSweep.Path = pFeat

# curves = {str([(curve.CircleParams, curve.LineParams) for curve in swSweep.GuideCurves])}
Part.Extension.SelectByID2({swSweep.Path.Name}, "SKETCH", 0, 0, 0, False, 0, Nothing, 0)
pFeat = Part.SelectionManager.GetSelectedObject6(1, -1)
swSweep.GuideCurves = [pFeat]

for body in {[body.Name for body in swSweep.FeatureScopeBodies]}:
    Part.Extension.SelectByID2(body.Name, "BODYFEATURE", 0, 0, 0, True, 0, Nothing, 0)
pFeat = Part.SelectionManager.GetSelectedObject6(1, -1)
swSweep.FeatureScopeBodies = pFeat

swSweep.FeatureScope = {swSweep.FeatureScope}
swSweep.EndTangencyType = {swSweep.EndTangencyType}
swSweep.Direction = {swSweep.Direction}
swSweep.D2ReverseTwistDir = {swSweep.D2ReverseTwistDir}
swSweep.D1ReverseTwistDir = {swSweep.D1ReverseTwistDir}
swSweep.CircularProfile = {swSweep.CircularProfile}
swSweep.AutoSelectComponents = {swSweep.AutoSelectComponents}
swSweep.AssemblyFeatureScope = {swSweep.AssemblyFeatureScope}
swSweep.CircularProfileDiameter = {swSweep.CircularProfileDiameter}
swSweep.SetWallThickness(True, {swSweep.GetWallThickness(True)})
swSweep.SetWallThickness(False, {swSweep.GetWallThickness(False)})
swSweep.SetTwistAngle({swSweep.GetTwistAngle})
swSweep.SetPathAlignmentDirectionVector({swSweep.GetPathAlignmentDirectionVector})
swSweep.SetD2TwistAngle({swSweep.GetD2TwistAngle})
feat = Part.FeatureManager.CreateFeature(swSweep)'''
                seq += lines.splitlines()

        subfeat = curFeat.GetFirstSubFeature

        while subfeat is not None:
            traverse_features(swApp, Part, subfeat, False, seq, traversed_features, active_sketch_name, pFaceID)
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
Part = swApp.ActiveDoc
Errors=VARIANT(VT_BYREF | VT_I4, -1)
Warnings=VARIANT(VT_BYREF | VT_I4, -1)

seq = initialize_seq()
# traverse_features(swApp, Part, Part.FirstFeature, True, seq, [], '', 0)
# with open('macros/Macro9.py', 'wt', encoding='utf-8') as f:
#     f.write('\n'.join(seq))

files = get_files(r'F:\机械模型库\20230414_01', 'sldprt')
files += get_files(r'F:\机械模型库\20230414_01', 'SLDPRT')
print(files)
out_dir = r'F:\datasets\SW数据集\提取数据\20230414_02'
if not os.path.exists(out_dir):
    os.mkdir(out_dir)
for filename in files:
    swApp.OpenDoc6(filename,1,1,"",Errors,Warnings)
    Part = swApp.ActiveDoc
    seq = initialize_seq()
    try:
        # seq.append('.'.join(os.path.basename(filename).split('.')[:-1]))
        # seq.append(','.join(['id', 'type',] + ['p' + str(i) for i in range(1, 27 + 1)]))
        traverse_features(swApp, Part, Part.FirstFeature, True, seq, [], '', 0)
        with open(os.path.join(out_dir, '.'.join(os.path.basename(filename).split('.')[:-1]) + '.py'), 'wt', encoding='utf-8') as f:
            f.write('\n'.join(seq))
    except Exception as e:
        print(e)
    swApp.CloseDoc("")
