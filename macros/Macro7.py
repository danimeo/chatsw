import win32com.client
from pythoncom import Nothing, VT_BYREF, VT_I4
from win32com.client import VARIANT
from swconst import constants

swApp = win32com.client.Dispatch("SldWorks.Application")
swApp.Visible = True
Part = swApp.ActiveDoc
line = Part.SketchManager.CreateLine
circle = Part.SketchManager.CreateCircle
arc = Part.SketchManager.CreateArc
enter_sketch = Part.EditSketch
exit_sketch = Part.SketchManager.InsertSketch
get_features = Part.FeatureManager.GetFeatures
select = Part.Extension.SelectByID2
extrude = Part.FeatureManager.FeatureExtrusion3
cut = Part.FeatureManager.FeatureCut4
feat_name = ""


def select_face(feature_name, normal):
    features = get_features(True)
    for pFeat in features[::-1]:
        if pFeat is not None and pFeat.Name == feature_name:
            for face in pFeat.GetFaces:
                if face.Normal == normal:
                    face.Select2(False, 0)
                    break


def select_sketch():
    name = [feat.Name for feat in get_features(True) if feat.GetTypeName == "ProfileFeature"][-1]
    select(name, "SKETCH", 0, 0, 0, False, 0, Nothing, 0)


select("前视基准面", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
enter_sketch()
arc(0.000000,0.000000,0.000000,-0.002500,0.005454,0.000000,-0.005454,0.002500,0.000000,1.000000)
arc(0.000000,0.007500,0.000000,0.002500,0.007500,0.000000,-0.002500,0.007500,0.000000,1.000000)
arc(0.007500,0.000000,0.000000,0.007500,-0.002500,0.000000,0.007500,0.002500,0.000000,1.000000)
arc(0.000000,-0.007500,0.000000,-0.002500,-0.007500,0.000000,0.002500,-0.007500,0.000000,1.000000)
arc(-0.007500,-0.000000,0.000000,-0.007500,0.002500,0.000000,-0.007500,-0.002500,0.000000,1.000000)
arc(0.000000,0.000000,0.000000,0.005454,0.002500,0.000000,0.002500,0.005454,0.000000,1.000000)
arc(0.000000,0.000000,0.000000,0.002500,-0.005454,0.000000,0.005454,-0.002500,0.000000,1.000000)
arc(0.000000,0.000000,0.000000,-0.005454,-0.002500,0.000000,-0.002500,-0.005454,0.000000,1.000000)
l = line(-0.002500,0.007500,0.000000,-0.002500,0.005454,0.000000)
l = line(0.002500,0.007500,0.000000,0.002500,0.005454,0.000000)
l = line(0.007500,0.002500,0.000000,0.005454,0.002500,0.000000)
l = line(0.007500,-0.002500,0.000000,0.005454,-0.002500,0.000000)
l = line(0.002500,-0.007500,0.000000,0.002500,-0.005454,0.000000)
l = line(-0.002500,-0.007500,0.000000,-0.002500,-0.005454,0.000000)
l = line(-0.007500,-0.002500,0.000000,-0.005454,-0.002500,0.000000)
l = line(-0.007500,0.002500,0.000000,-0.005454,0.002500,0.000000)
exit_sketch(True)
select_sketch()
feat = extrude(True, False, True, constants.swEndCondBlind, 0, 0.003, 0.003, False, False, False, False, 0, 0, False, False, False, False, True, True, True, constants.swStartSketchPlane, 0, False)
select_face(feat.Name, (0.0, 0.0, 1.0))
enter_sketch()
circle(0.000000,0.000000,0.000000,0.003825,0.001169,0.000000)
exit_sketch(True)
select_sketch()
feat = extrude(True, False, True, constants.swEndCondBlind, 0, 0.006, 0.006, False, False, False, False, 0, 0, False, False, False, False, True, True, True, constants.swStartSketchPlane, 0, False)
select_face(feat.Name, (0.0, 0.0, 1.0))
enter_sketch()
arc(0.000000,0.000000,0.000000,-0.001323,0.001500,0.000000,0.001323,0.001500,0.000000,1.000000)
l = line(-0.001323,0.001500,0.000000,0.001323,0.001500,0.000000)
exit_sketch(True)
select_sketch()
feat = cut(True, False, False, 0, 0, 0.012, 0.012, False, False, False, False, 0, 0, False, False, False, False, False, False, True, True, True, False, 0, 0, False, True)
select_face(feat.Name, (-0.0, -0.0, -1.0))
enter_sketch()
circle(0.000000,0.000000,0.000000,0.002869,0.000877,0.000000)
exit_sketch(True)
select_sketch()
feat = extrude(True, False, True, constants.swEndCondBlind, 0, 0.001, 0.001, False, False, False, False, 0, 0, False, False, False, False, True, True, True, constants.swStartSketchPlane, 0, False)
select_face(feat.Name, (0.0, 0.0, -1.0))
enter_sketch()
arc(0.000000,0.000000,0.000000,-0.001732,0.001000,0.000000,0.001732,0.001000,0.000000,1.000000)
l = line(-0.001732,0.001000,0.000000,0.001732,0.001000,0.000000)
exit_sketch(True)
select_sketch()
feat = extrude(True, False, True, constants.swEndCondBlind, 0, 0.002, 0.002, False, False, False, False, 0, 0, False, False, False, False, True, True, True, constants.swStartSketchPlane, 0, False)
select_face(feat.Name, (-0.0, -0.0, -1.0))
enter_sketch()
exit_sketch(True)