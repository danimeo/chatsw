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
    '',
    'swApp = win32com.client.Dispatch("SldWorks.Application")',
    'swApp.Visible = True',
    'Part = swApp.ActiveDoc',
    'line = Part.SketchManager.CreateLine',
    'circle = Part.SketchManager.CreateCircle',
    'arc = Part.SketchManager.CreateArc',
    'enter_sketch = Part.EditSketch',
    'exit_sketch = Part.SketchManager.InsertSketch',
    'get_features = Part.FeatureManager.GetFeatures',
    'select = Part.Extension.SelectByID2',
    'extrude = Part.FeatureManager.FeatureExtrusion3',
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
                    # c2 = 'rels = ['
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

                            # print("0,{},{}".format(','.join(['{:.5f}'.format(d) for d in data]), c))
                            seq.append('l = line({})'.format(','.join(['{:.6f}'.format(d) for d in data])))
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
                                seq.append('circle({})'.format(','.join(['{:.6f}'.format(d) for d in data])))
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
                                seq.append('arc({})'.format(','.join(['{:.6f}'.format(d) for d in data])))

    #                     rels = sketchSegment.GetRelations
    #                     if rels is not None:
    #                         for rel in rels:
    #                             try:
    #                                 names = ['"{}"'.format(ent.GetName) for ent in rel.GetEntities]
    #                                 rel_type = rel.GetRelationType
    #                                 if rel_type == constants.swConstraintType_HORIZONTAL:
    #                                     rel_type = '"sgHORIZONTAL2D"'
    #                                 elif rel_type == constants.swConstraintType_ALONGX3D:
    #                                     rel_type = '"sgALONGX3D"'
    #                                 elif rel_type == constants.swConstraintType_HORIZPOINTS:
    #                                     rel_type = '"sgHORIZONTALPOINTS2D"'
    #                                 elif rel_type == constants.swConstraintType_ALONGXPOINTS3D:
    #                                     rel_type = '"sgALONGXPOINTS3D"'
    #                                 elif rel_type == constants.swConstraintType_VERTICAL:
    #                                     rel_type = '"sgVERTICAL2D"'
    #                                 elif rel_type == constants.swConstraintType_ALONGY3D:
    #                                     rel_type = '"sgALONGY3D"'
    #                                 elif rel_type == constants.swConstraintType_VERTPOINTS:
    #                                     rel_type = '"sgVERTPOINTS2D"'
    #                                 elif rel_type == constants.swConstraintType_ALONGYPOINTS3D:
    #                                     rel_type = '"sgALONGYPOINTS3D"'
    #                                 elif rel_type == constants.swConstraintType_ALONGZPOINTS:
    #                                     rel_type = '"sgALONGZPOINTS3D"'
    #                                 elif rel_type == constants.swConstraintType_ALONGZ:
    #                                     rel_type = '"sgALONGZ3D"'
    #                                 elif rel_type == constants.swConstraintType_COLINEAR:
    #                                     rel_type = '"sgCOLINEAR"'
    #                                 elif rel_type == constants.swConstraintType_CORADIAL:
    #                                     rel_type = '"sgCORADIAL"'
    #                                 elif rel_type == constants.swConstraintType_PERPENDICULAR:
    #                                     rel_type = '"sgPERPENDICULAR"'
    #                                 elif rel_type == constants.swConstraintType_PARALLEL:
    #                                     rel_type = '"sgPARALLEL"'
    #                                 elif rel_type == constants.swConstraintType_TANGENT:
    #                                     rel_type = '"sgTANGENT"'
    #                                 elif rel_type == constants.swConstraintType_CONCENTRIC:
    #                                     rel_type = '"sgCONCENTRIC"'
    #                                 elif rel_type == constants.swConstraintType_COINCIDENT:
    #                                     rel_type = '"sgCOINCIDENT"'
    #                                 elif rel_type == constants.swConstraintType_SYMMETRIC:
    #                                     rel_type = '"sgSYMMETRIC"'
    #                                 elif rel_type == constants.swConstraintType_ATMIDDLE:
    #                                     rel_type = '"sgATMIDDLE"'
    #                                 elif rel_type == constants.swConstraintType_ATINTERSECT:
    #                                     rel_type = '"sgATINTERSECT"'
    #                                 elif rel_type == constants.swConstraintType_ATPIERCE:
    #                                     rel_type = '"sgATPIERCE"'
    #                                 elif rel_type == constants.swConstraintType_FIXED:
    #                                     rel_type = '"sgFIXED"'
    #                                 elif rel_type == constants.swConstraintType_ANGLE:
    #                                     rel_type = '"sgANGLE"'
    #                                 elif rel_type == constants.swConstraintType_ARCANG180:
    #                                     rel_type = '"sgARCANG180"'
    #                                 elif rel_type == constants.swConstraintType_ARCANG270:
    #                                     rel_type = '"sgARCANG270"'
    #                                 elif rel_type == constants.swConstraintType_ARCANG90:
    #                                     rel_type = '"sgARCANG90"'
    #                                 elif rel_type == constants.swConstraintType_ARCANGBOTTOM:
    #                                     rel_type = '"sgARCANGBOTTOM"'
    #                                 elif rel_type == constants.swConstraintType_ARCANGLEFT:
    #                                     rel_type = '"sgARCANGLEFT"'
    #                                 elif rel_type == constants.swConstraintType_ARCANGRIGHT:
    #                                     rel_type = '"sgARCANGRIGHT"'
    #                                 elif rel_type == constants.swConstraintType_ARCANGTOP:
    #                                     rel_type = '"sgARCANGTOP"'
    #                                 elif rel_type == constants.swConstraintType_DIAMETER:
    #                                     rel_type = '"sgDIAMETER"'
    #                                 elif rel_type == constants.swConstraintType_DISTANCE:
    #                                     rel_type = '"sgDISTANCE"'
    #                                 elif rel_type == constants.swConstraintType_SAMELENGTH:
    #                                     rel_type = '"sgSAMELENGTH"'
    #                                 elif rel_type == constants.swConstraintType_OFFSETEDGE:
    #                                     rel_type = '"sgOFFSETEDGE"'
    #                                 elif rel_type == constants.swConstraintType_SNAPANGLE:
    #                                     rel_type = '"sgSNAPANGLE"'
    #                                 elif rel_type == constants.swConstraintType_SNAPGRID:
    #                                     rel_type = '"sgSNAPGRID"'
    #                                 elif rel_type == constants.swConstraintType_SNAPLENGTH:
    #                                     rel_type = '"sgSNAPLENGTH"'
    #                                 elif rel_type == constants.swConstraintType_USEEDGE:
    #                                     rel_type = '"sgUSEEDGE"'
    #                                 elif rel_type == constants.swConstraintType_MERGEPOINTS:
    #                                     rel_type = '"sgMERGEPOINTS"'
    #                                 relation = '({},{})'.format(rel_type, ','.join(names))
    #                                 if relation not in relations:
    #                                     relations.append(relation)
    #                                     c2 += '{},'.format(relation)
    #                             except AttributeError as e:
    #                                 print(e)
    #                 c2 += ']\n'
                    
    #                 c2 += '''
    # for rel in rels:
    #     Part.ClearSelection2(True)
    #     for ent_name in rel[1:]:
    #         Part.Extension.SelectByID2(ent_name, "SKETCHSEGMENT", 0, 0, 0, True, 0, Nothing, 0)
    #     Part.SketchAddConstraints(rel[0])
    # '''
    #             seq.append(c2)
                seq.append('exit_sketch(True)')

            active_sketch_name = curFeat.Name
        elif typeName in ('Boss', 'Extrusion'):
            seq.append('select_sketch()')
            depth = curFeat.GetDefinition.GetDepth(True)
            if depth != 0.:
                seq.append(f'feat = extrude(True, False, True, constants.swEndCondBlind, 0, {depth}, {depth}, False, False, False, False, 0, 0, False, False, False, False, True, True, True, constants.swStartSketchPlane, 0, False)')
        elif typeName == 'Cut':
            seq.append('select_sketch()')
            depth = curFeat.GetDefinition.GetDepth(True)
            if depth != 0.:
                seq.append(f'feat = cut(True, False, False, 0, 0, {depth}, {depth}, False, False, False, False, 0, 0, False, False, False, False, False, False, True, True, True, False, 0, 0, False, True)')

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


swApp = win32com.client.Dispatch("SldWorks.Application")
Part = swApp.ActiveDoc

traverse_features(swApp, Part, Part.FirstFeature, True, seq, [], '', 0)

with open('macros/Macro7.py', 'wt', encoding='utf-8') as f:
    f.write('\n'.join(seq))