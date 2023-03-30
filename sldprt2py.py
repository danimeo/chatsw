import re
import win32com
from pythoncom import Nothing, VT_BYREF, VT_I4
from win32com.client import VARIANT
from swconst import constants

seq = [
    'import win32com.client',
    'from pythoncom import Nothing, VT_BYREF, VT_I4',
    'from win32com.client import VARIANT',
    'from swconst import constants',
    'swApp = win32com.client.Dispatch("SldWorks.Application")',
    'swApp.Visible = True',
    'Part = swApp.ActiveDoc',
]


def traverse_features(swApp, Part, thisFeat, isTopLevel, seq, traversed_features, active_sketch_name):
    # print(thisFeat)
    curFeat = thisFeat
    while curFeat is not None:
        name = curFeat.Name
        typeName = curFeat.GetTypeName
        print(curFeat.Name, curFeat.GetTypeName)

        if typeName == "ProfileFeature":
            Part.ClearSelection2(True)
            Part.Extension.SelectAll
            Part.Extension.SelectByID2(name, "SKETCH", 0, 0, 0, False, 0, Nothing, 0)

            Part.SketchManager.InsertSketch(True)
            sketch = Part.GetActiveSketch2
            if sketch is not None:
                
                v = VARIANT(VT_BYREF | VT_I4, 0)
                ref = sketch.GetReferenceEntity(v)
                ref.Select2(False, 0)
                pFeat = Part.SelectionManager.GetSelectedObject5(1)

                if name not in traversed_features and pFeat is not None:
                    seq.append('Part.ClearSelection2(True)')
                    if v.value in (constants.swReferenceTypeFace, constants.swReferenceTypeEdge, constants.swReferenceTypeBody):
                        normal = pFeat.Normal
                        print('参考面：', pFeat.GetFeature.Name, normal)
                        # seq.append('normal = {}'.format(str(normal)))
                        
                        seq.append('''
features = Part.FeatureManager.GetFeatures(True)
for pFeat in features[::-1]:
    if pFeat is not None and pFeat.GetTypeName in ("Boss", "Cut", "Extrusion"):
        for face in pFeat.GetFaces:
            if face.Normal == {}:
                face.Select2(False, 0)
                break
'''.format(str(normal)))
                    else:
                        print('参考面：', pFeat.Name)
                        seq.append('Part.Extension.SelectByID2("{}", "PLANE", 0, 0, 0, True, 0, Nothing, 0)'.format(pFeat.Name))
                    
                seq.append('Part.SketchManager.InsertSketch(True)')
                relations = []
                
                sketchSegments = sketch.GetSketchSegments
                if sketchSegments is not None:
                    # 遍历
                    
                    c2 = 'rels = ['
                    for sketchSegment in sketchSegments:

                        # 判断是直线时执行
                        if sketchSegment.GetType == constants.swSketchLINE:
                            sketchPointStart = sketchSegment.GetStartPoint2
                            sketchPointEnd = sketchSegment.GetEndPoint2

                            data = [
                                sketchPointStart.X,
                                sketchPointStart.Y,
                                sketchPointStart.Z,
                                sketchPointEnd.X,
                                sketchPointEnd.Y,
                                sketchPointEnd.Z,
                            ]

                            # print("0,{},{}".format(','.join(['{:.5f}'.format(d) for d in data]), c))
                            seq.append('line = Part.SketchManager.CreateLine({})'.format(','.join(['{:.6f}'.format(d) for d in data])))
                            seq.append('line.ConstructionGeometry = {}'.format(str(sketchSegment.ConstructionGeometry)))
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
                                seq.append('Part.SketchManager.CreateCircle({})'.format(','.join(['{:.6f}'.format(d) for d in data])))
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
                                # seq.append('data = {}'.format(','.join(['{:.5f}'.format(d) for d in data])))
                                seq.append('Part.SketchManager.CreateArc({})'.format(','.join(['{:.6f}'.format(d) for d in data])))

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
                # seq.append(c2)
                
            active_sketch_name = curFeat.Name
        elif typeName == 'Boss':
            seq.append('Part.SketchManager.InsertSketch(True)')
            seq.append('name = [feat.Name for feat in Part.FeatureManager.GetFeatures(True) if feat.GetTypeName == "ProfileFeature"][-1]')
            seq.append('Part.Extension.SelectByID2(name, "SKETCH", 0, 0, 0, False, 0, Nothing, 0)')
            
            featData = curFeat.GetDefinition
            depth = featData.GetDepth(True)
            seq.append(f'Part.FeatureManager.FeatureExtrusion3(True, False, True, constants.swEndCondBlind, 0, {depth}, {depth}, False, False, False, False, 0, 0, False, False, False, False, True, True, True, constants.swStartSketchPlane, 0, False)')
        elif typeName == 'Cut':
            seq.append('Part.SketchManager.InsertSketch(True)')
            seq.append('name = [feat.Name for feat in Part.FeatureManager.GetFeatures(True) if feat.GetTypeName == "ProfileFeature"][-1]')
            seq.append('Part.Extension.SelectByID2(name, "SKETCH", 0, 0, 0, False, 0, Nothing, 0)')
            
            featData = curFeat.GetDefinition
            depth = featData.GetDepth(True)
            seq.append(f'Part.FeatureManager.FeatureCut4(True, False, False, 0, 0, {depth}, {depth}, False, False, False, False, 0, 0, False, False, False, False, False, False, True, True, True, False, 0, 0, False, True)')

        

        subfeat = curFeat.GetFirstSubFeature

        while subfeat is not None:
            traverse_features(swApp, Part, subfeat, False, seq, traversed_features, active_sketch_name)
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


swApp = win32com.client.Dispatch("SldWorks.Application")
Part = swApp.ActiveDoc

traverse_features(swApp, Part, Part.FirstFeature, True, seq, [], '')
# seq.append('Part.SketchManager.InsertSketch(True)')

with open('macros/Macro4.py', 'wt', encoding='utf-8') as f:
    f.write('\n'.join(seq))