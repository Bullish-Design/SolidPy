import csv
import time
from dataclasses import dataclass

import pandas
import pythoncom
import win32com.client
from win32com.client import Dispatch, gencache

swConst = win32com.client.gencache.EnsureModule(
    "{4687F359-55D0-4CD3-B6CF-2EB42C11F989}", 0, 29, 0
).constants  # sw2015
swCmd = win32com.client.gencache.EnsureModule(
    "{0AC1DE9F-3FBC-4C25-868D-7D4E9139CCE0}", 0, 29, 0
).constants
# # {83A33D31-27C5-11CE-BFD4-00400513BB57}
arg1 = win32com.client.VARIANT(pythoncom.VT_DISPATCH, None)
var1 = win32com.client.VARIANT(
    pythoncom.VT_VARIANT | pythoncom.VT_BYREF, None
)  # (pythoncom.VT_VARIANT | pythoncom.VT_NULL | pythoncom.VT_BYREF, None)
varVals = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_BSTR, -1)
varBool = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_BOOL, None)
varNone = win32com.client.VARIANT(pythoncom.VT_EMPTY, None)
Empty = win32com.client.VARIANT(pythoncom.VT_EMPTY, None)
newDrw = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, -1)

sw = win32com.client.Dispatch("SldWorks.Application")

# sw = sldworks.ISldWorks(DispatchEx('SldWorks.Application'))


# # -----------------------------------------------------------------------------------------------------------------------
# # ---------------------------------------------- Gameplan for this API: -----------------------------------------------------
# """

# These classes will use a csv file as an intermediary template.
# It will act as somewhat of a database during usage, but will also provide I/O to be able to edit from markdown/Dendron.
#     - Can just edit directly in dendron, call "update" from dendron, and it will update the csv file, then update the SW model from the CSV file.

# Functionality needed:
#     - Create a new part/asm/drawing from a template
#         - template will be pre-loaded with certain info, and can gradually add more templates in the future.
#     - structured as a folder with multiple csv files inside representing different aspects of the SW model:
#         - properties
#         - features
#         - feature definitions
#     - For assemblies, the part folders will sit within an assembly folder that will also have its own properties/features/feature definitions csv files.
#         - also a mates csv file to define placement and configuration of mates
#     - Eventually build into using a database instead of csv files to allow easier cross referencing and adaptability, but for now, csv files will be easier to work with.

# Properties Template:
#     - File name, revision, description, author, date, etc.
#     - Reference to the project/task that the model is associated with in Dendron
# Features Template:
#     - Feature name, feature ID, feature type (in tree order?)
# Feature Definitions Template:
#     - Prodvies the definition of each feature type, and the parameters that are needed to define it.


# """
# # -----------------------------------------------------------------------------------------------------------------------
# # -----------------------------------------------------------------------------------------------------------------------


@dataclass
class Solidworks:
    def __init__(self):
        self.instance = win32com.client.Dispatch("SldWorks.Application")

    pass


@dataclass
class Model(Solidworks):

    pass


@dataclass
class Part(Solidworks):
    # class for a solidworks part
    # inherits from Model class?
    pass


class Feature:

    pass


class DRW:
    def __init__(self, swObj):
        self.swObj = swObj
        self.model = swObj.ActiveDoc
        # self.sketchName = sketchName
        # self.name = drwName

    def create(self):
        # creates a drawing from part
        # print("  Creating drawing...")
        swModel = self.model
        # swSelMgr = swModel.SelectionManager
        # swEnt = swSelMgr.GetSelectedObject6(1, -1)
        pthName = swModel.GetPathName
        drwPath = pthName.replace(".SLDPRT", ".SLDDRW")
        print("  Creating drawing from: " + swModel.getPathName)  # getPathName?
        drawTemplate = self.swObj.GetUserPreferenceStringValue(
            swConst.swDefaultTemplateDrawing
        )
        # swDraw = swModel
        swModel = self.swObj.NewDocument(drawTemplate, swConst.swDwgPaperCsize, 0, 0)
        # swModel = self.swObj.OpenDoc6(drwPath, 3, swConst.swOpenDocOptions_Silent, '', newDrw, newDrw)
        swDraw = swModel
        # print("  Drawing created boolean " + str(swDraw))
        # swDraw = swModel
        # vSheetNames = swDraw.GetSheetNames
        # swDraw.ActivateSheet(vSheetNames[0])

        # print("  Sheet names: " + vSheetNames[0])
        # swDraw.ActivateSheet(vSheetNames[0])
        swSheet = swDraw.GetCurrentSheet
        # swSheet = swDraw.Sheet(vSheetNames[0])
        # print(self.model.getPathName)
        # print(pthName)

        # boolStat = swDraw.Create3rdAngleViews2(pthName)
        createStandardViews(swDraw, pthName)
        saveBool = swModel.Extension.SaveAs(drwPath, 0, 1, arg1, newDrw, newDrw)
        print("  Drawing saved: " + str(saveBool))
        return swDraw

    pass


class Sketch:
    def __init__(
        self, swModel, name, sketchPlane
    ):  # , refPlane, offsetDistance, distanceUnits):
        self.model = swModel
        self.name = name
        self.plane = sketchPlane
        # self.refPlane = refPlane
        # self.offsetDistance = offsetDistance
        # self.distanceUnits = distanceUnits

    def create(self):
        sketchMgr = self.model.SketchManager
        # sketchMgr.InsertSketch(False)
        selMgr = self.model.SelectionManager
        # planeName = selMgr.GetSelectedObject6(1, 0).GetSpecificFeature2
        selectedPlane = selectItemByName(self.model, self.plane, "PLANE")
        planeName = selMgr.GetSelectedObject6(1, 0).GetSpecificFeature2
        sketchMgr.InsertSketch(True)
        swSketch = sketchMgr.ActiveSketch
        swSketch.Name = self.name  # + "_TestSketch"
        # print(" Created new sketch on: " + planeName.Name)
        # sketchMgr.InsertSketch(False)
        # clearSelections(self.model)
        return


class Plane:
    def __init__(self, swModel, name, refPlane, offsetDistance, distanceUnits):
        self.model = swModel
        self.name = name
        self.refPlane = refPlane
        self.offsetDistance = offsetDistance
        self.distanceUnits = distanceUnits

    def create(self):
        # creates a plane after selecting the reference plane given in the function
        print("Creating plane")
        modelExt = self.model.Extension
        # modelExt = self.model.Extension
        featureMgr = self.model.FeatureManager
        distance = convertUnits(self.offsetDistance, self.distanceUnits)
        BoolStatus = modelExt.SelectByID2(
            self.refPlane, "PLANE", 0, 0, 0, False, 0, arg1, 0
        )
        # print(BoolStatus)
        print(distance < 0)
        if distance < 0:
            CreatePlane = featureMgr.InsertRefPlane(
                swConst.swRefPlaneReferenceConstraint_Distance
                | swConst.swRefPlaneReferenceConstraint_OptionFlip,
                distance,
                0,
                0,
                0,
                0,
            )
        else:
            CreatePlane = featureMgr.InsertRefPlane(
                swConst.swRefPlaneReferenceConstraint_Distance, distance, 0, 0, 0, 0
            )
        # print(CreatePlane)
        CreatePlane.Name = self.name
        self.model.ClearSelection2(True)
        return


    def edit(self):  # , swModel, refPlane, offsetDistance, distanceUnits):
        # edits a plane after selecting the reference plane given in the function
        try:
            modelExt = self.model.Extension
            # modelExt = self.model.Extension
            # featureMgr = self.model.FeatureManager
            swSelMgr = self.model.SelectionManager
            distance = convertUnits(self.offsetDistance, self.distanceUnits)
            BoolStatus = modelExt.SelectByID2(
                self.name, "PLANE", 0, 0, 0, False, 0, arg1, 0
            )
            # print(BoolStatus)
            Feature = swSelMgr.GetSelectedObject(1)
            swRefPlane = Feature.GetDefinition
            # swRefPlane.Distance = distance
            swRefPlane.AccessSelections(self.model, arg1)
            flip = swRefPlane.ReverseDirection  # ReversedReferenceDirection
            print(flip)
            if distance < 0:
                flip = True  # swConst.swRefPlaneReferenceConstraint_OptionFlip
            print(flip)
            swRefPlane.ReverseDirection = flip  # (swConst.swRefPlaneReference_First)
            swRefPlane.Distance = distance
            Feature.ModifyDefinition(swRefPlane, self.model, arg1)
            self.model.ClearSelection2(True)
            # print(BoolStatus)
            # CreatePlane = featureMgr.EditRefPlane(swConst.swRefPlaneReferenceConstraint_Distance, distance, 0, 0, 0, 0)
            # print(CreatePlane)
            # swModel.ClearSelection2(True)
        except:
            swRefPlane.ReleaseSelectionAccess
            print("Error editing plane: " + self.name)
        return

    def getProperties(self):
        try:
            # gets the properties of the plane
            properties = []  # "RefPlane"]
            swSelMgr = self.model.SelectionManager
            print("Getting properties of plane: " + self.name)
            BoolStatus = self.model.Extension.SelectByID2(
                self.name, "PLANE", 0, 0, 0, False, 0, arg1, 0
            )
            print(str(BoolStatus))
            if BoolStatus == False:
                print("Error selecting plane: " + self.name)
                return
            Feature = swSelMgr.GetSelectedObject6(1, 0)
            print(str(Feature))
            swRefPlane = Feature.GetDefinition
            swRefPlane.AccessSelections(self.model, arg1)
            # if swRefPlane.Selections != None:
            print(str(swRefPlane.Selections))
            try:
                properties.append(swRefPlane.Selections[0].Name)
            except:
                properties.append(swRefPlane.Selections.Name)

            properties.append(swRefPlane.Distance * 1000)
            # properties.append(swRefPlane.DistanceType)
            # properties.append(swRefPlane.DistanceUnits)
            # properties.append(swRefPlane.PlaneType)
            self.refPlane = swRefPlane.Selections[0].Name
            self.offsetDistance = swRefPlane.Distance * 1000
            # self.distanceUnits = properties[1]
            swRefPlane.ReleaseSelectionAccess()  # (self.model)
            return properties
        except:
            print("Error getting properties of plane: " + self.name)
            swRefPlane.ReleaseSelectionAccess()  # (self.model)
            return


class Dimension:
    def __init__(self, swModel, dimName, sketchName, units):
        self.model = swModel
        self.name = dimName
        self.sketchName = sketchName
        # self.dimType = dimType
        # self.value = value
        self.units = units

    def get(self):
        # gets the value of the dimension
        fullDimName = self.name + "@" + self.sketchName + "@" + self.model.GetTitle
        swDim = self.model.Parameter(fullDimName)
        # swFeat = self.model.FeatureByName(self.sketchName)
        vVal = swDim.GetSystemValue3(swConst.swThisConfiguration, Empty)
        dimVal = convertUnits(vVal[0], self.units)
        # swDispDim = swFeat.GetFirstDisplayDimension
        # while swDispDim:
        #     swDim = swDispDim.GetDimension2(Empty)
        #     if swDim.Name == self.name:
        #         vDimVals = swDim.GetValue3(swConst.swThisConfiguration, Empty)
        #         return vDimVals[0]
        #     swDispDim = swDispDim.GetNext
        return dimVal

    def set(self, value):
        # sets the value of the dimension
        fullDimName = self.name + "@" + self.sketchName + "@" + self.model.GetTitle
        swDim = self.model.Parameter(fullDimName)
        vVal = swDim.GetSystemValue3(swConst.swThisConfiguration, Empty)
        dimension = convertUnits(value, self.units)
        swDim.SetSystemValue3(dimension, swConst.swThisConfiguration, Empty)
        # vVal = swDim.GetSystemValue3(swConst.swThisConfiguration, Empty)
        # print("  Name: " + swDim.Name + "  |  Value: " + str(vVal[0]))
        return  # vVal[0]

    pass


class Line:
    def __init__(self, swModel, sketchName, lineName):
        self.model = swModel
        self.sketchName = sketchName
        self.name = lineName

    def create(self, constraintDirection, length, units):
        # length = convertUnits(length, units)
        if constraintDirection == "Horizontal":
            Xstart = -length / 2
            Ystart = 0
            Zstart = 0
            Xend = length / 2
            Yend = 0
            Zend = 0
        elif constraintDirection == "Vertical":
            Xstart = 0
            Ystart = 0
            Zstart = 0
            Xend = 0
            Yend = length
            Zend = 0
        else:
            print("Invalid constraint direction - must be 'Horizontal' or 'Vertical'")
            return
        # creates a sketch on the given plane
        selectItemByName(self.model, self.sketchName, "SKETCH")
        sketchLine = self.sketchLine(Xstart, Ystart, Zstart, Xend, Yend, Zend, units)
        clearSelections(self.model)
        return sketchLine

    def sketchLine(self, Xstart, Ystart, Zstart, Xend, Yend, Zend, units):
        # creates a line in the sketch
        sketchMgr = self.model.SketchManager
        XstartConv = convertUnits(Xstart, units)
        YstartConv = convertUnits(Ystart, units)
        ZstartConv = convertUnits(Zstart, units)
        XendConv = convertUnits(Xend, units)
        YendConv = convertUnits(Yend, units)
        ZendConv = convertUnits(Zend, units)
        sketchSegment = sketchMgr.CreateLine(
            XstartConv, YstartConv, ZstartConv, XendConv, YendConv, ZendConv
        )
        # sketchMgr_Obj.InsertSketch(False)
        return sketchSegment


class Extrude:
    # class for getting and setting extrude features, both bosses and cuts
    pass


def createStandardViews(swDraw, pathName):
    view1 = ["*Front", 0.15, 0.15, 0.1]
    view2 = ["*Left", 0.15, 0.3, 0.1]
    view3 = ["*Top", 0.4, 0.3, 0.1]

    # print("  Creating standard views...")
    for view in [view1, view2, view3]:
        # print("  Creating view: " + view[0])
        boolStat = swDraw.CreateDrawViewFromModelView3(
            pathName, view[0], view[1], view[2], view[3]
        )
        if boolStat == None:  # print("  Views created: " + str(boolStat))
            raise Exception("  Error creating views")
    # # start managers
    # #model = sw.ActiveDoc
    # modelExt = model.Extension
    # selMgr = model.SelectionManager
    # featureMgr = model.FeatureManager
    # sketchMgr = model.SketchManager
    return


def getAllProperties_toCSV(model, saveLoc, saveName):
    # Snags all properties from the model and saves them to a CSV file. Overwrites existing file each time.
    vNames, vTypes, vValues = getAllPropertyNames(model)

    with open(saveLoc + "\\" + saveName, "w", newline="") as csvfile:
        writer = csv.writer(
            csvfile, delimiter=",", quotechar='"', quoting=csv.QUOTE_MINIMAL
        )
        writer.writerow(["Name", "Type", "Value"])
        for i in range(len(vNames.value)):
            writer.writerow(
                [
                    str(vNames.value[i]),
                    convertPropertyType(vTypes.value[i]),
                    str(vValues.value[i]),
                ]
            )
    return


# def getAllFeatures_toCSV(model, saveLoc, saveName):
#     # Snags all features from the model and saves them to a CSV file. Overwrites existing file each time.
#     vNames, vTypes, vValues = getAllFeatureNames(model)

#     with open(saveLoc + "\\" + saveName, "w", newline="") as csvfile:
#         writer = csv.writer(
#             csvfile, delimiter=",", quotechar='"', quoting=csv.QUOTE_MINIMAL
#         )
#         writer.writerow(["Name", "Type", "Value"])
#         for i in range(len(vNames.value)):
#             writer.writerow(
#                 [
#                     str(vNames.value[i]),
#                     convertFeatureType(vTypes.value[i]),
#                     str(vValues.value[i]),
#                 ]
#             )
#     return


def getAllFeatureDefs_toCSV(model, saveLoc, saveName):
    # Snags all feature definitions from the model and saves them to a CSV file. Overwrites existing file each time.

    return


def insertPart(swModel, partToInsert):
    # inserts a part into the current model
    importPlanes = True
    importAxis = True
    importCThread = True
    FileName = partToInsert
    # print("  Inserting part: " + FileName)
    newFeature = swModel.InsertPart2(
        FileName,
        swConst.swInsertPartImportPlanes
        | swConst.swInsertPartImportUnabsorbedSketchs
        | swConst.swInsertPartImportAbsorbedSketchs
        | swConst.swInsertPartImportAxes,
    )
    print("  Inserted part: " + str(newFeature.Name))
    return newFeature


# # def getAllSketch


def getAllSketchDims(swModel, sketchName):
    # gets all dimensions in the sketch
    swFeat = swModel.FeatureByName(sketchName)
    swDispDim = swFeat.GetFirstDisplayDimension
    dimList = []
    while swDispDim:
        dimVal = {}
        swDim = swDispDim.GetDimension2(Empty)
        vDimVals = swDim.GetValue3(swConst.swThisConfiguration, Empty)
        print("  Name: " + swDim.Name + "  |  Value: " + str(vDimVals[0]))
        dimVal[swDim.Name] = round(vDimVals[0], 3)
        dimList.append(dimVal)
        swDispDim = swFeat.GetNextDisplayDimension(swDispDim)
    # dimMgr = swModel.ParameterManager
    # dimList = dimMgr.GetDimensions()
    # print("Dimensions: ")
    # for dim in dimList:
    #     print("  Name: " + dim.Name + "  |  Value: " + dim.ValueAsString)
    return dimList


def createNewPart(swObj, partTitle, fileLoc=""):
    defaultTemplate = swObj.GetUserPreferenceStringValue(swConst.swDefaultTemplatePart)
    CreatePart = swObj.NewDocument(defaultTemplate, 0, 0, 0)
    print("Created: " + str(CreatePart))
    model = swObj.ActiveDoc
    model.SetTitle2(partTitle)
    if fileLoc != "":
        model.SaveAs(fileLoc + "\\" + partTitle + ".SLDPRT")
    return model


def createPlane_OffsetDistance(swModel, refPlane, offsetDistance, distanceUnits):
    # creates a plane after selecting the reference plane given in the function

    modelExt = swModel.Extension
    featureMgr = swModel.FeatureManager
    distance = convertUnits(offsetDistance, distanceUnits)
    BoolStatus = modelExt.SelectByID2(refPlane, "PLANE", 0, 0, 0, False, 0, arg1, 0)
    print(BoolStatus)
    CreatePlane = featureMgr.InsertRefPlane(
        swConst.swRefPlaneReferenceConstraint_Distance, distance, 0, 0, 0, 0
    )
    print(CreatePlane)
    swModel.ClearSelection2(True)
    return


def create_RefAxis(swModel, plane1_Name, plane2_Name):
    modelExt = swModel.Extension
    featureMgr = swModel.FeatureManager
    BoolStatus1 = swModel.Extension.SelectByID2(
        plane1_Name, "PLANE", 0, 0, 0, False, 0, arg1, 0
    )
    BoolStatus2 = swModel.Extension.SelectByID2(
        plane2_Name, "PLANE", 0, 0, 0, True, 0, arg1, 0
    )
    BoolStatus3 = swModel.InsertAxis2(True)
    # Add code to rename the axis here
    # #CreatePlane.Name = self.name
    swModel.ClearSelection2(True)

    return


def resetView(swModel):
    # resets the view to the default view
    # modelExt = swModel.Extension
    swModel.ViewZoomtofit2()
    return


def converUnits_XYZ(X, Y, Z, unit):
    # converts various units to metres:
    returnX = convertUnits(X, unit)
    returnY = convertUnits(Y, unit)
    returnZ = convertUnits(Z, unit)
    return returnX, returnY, returnZ


def convertUnits(distance, unit):
    # converts various units to metres:
    if unit == "mm":
        returnDistance = distance / 1000
    if unit == "in":
        returnDistance = distance * 0.0254
    if unit == "ft":
        returnDistance = distance * (0.0254 * 12)
    return returnDistance


def toggleDimInputBox(swObj, bool):
    # toggles the dimension input box on or off
    print("Setting dim input box to: " + str(bool))
    swObj.SetUserPreferenceToggle(swConst.swInputDimValOnCreate, bool)
    return


# # plane3 = createPlane_OffsetDistance("Right Plane", 2, "ft")


def traverseFeatures(swModel):
    print("Finding Features for: " + swModel.GetTitle)
    featureMgr = swModel.FeatureManager
    allFeats = featureMgr.GetFeatures(True)
    # featList = []
    for feat in allFeats:
        print("    FeatureDef: " + str(feat.GetDefinition))
        # if feat.GetTypeName2 == "ProfileFeature":
        #     swDef = feat.CreateDefinition()
        #     print(str(swDef))
        #     #swSketch = swDef.
        # print("  Feature: {}  |  Type: {}  |  Definition: {}".format(str(feat.Name), str(feat.GetTypeName2) , str(feat.GetDefinition)))
        print(
            f"  Feature: {str(feat.Name)}  |  ID: {feat.GetID}  |  Type: {str(feat.GetTypeName2)} "
        )
    # print(str(allFeats))
    """swFeat = swModel.FirstFeature
    print("  "+ str(swFeat))
    while not swFeat == None:
        name = swFeat.Name
        swFeat = swFeat.GetNextFeature
        print("  " + name) """
    return allFeats


def getSketchSegments(swModel, sketchName):  # , sketchType):
    # returns a list of sketch segments

    # Public Enum swSketchSegments_e
    # swSketchLINE = 0
    # swSketchARC = 1
    # swSketchELLIPSE = 2
    # swSketchSPLINE = 3
    # swSketchTEXT = 4
    # swSketchPARABOLA = 5
    # End Enum

    selectItemByName(swModel, sketchName, "SKETCH")

    swSelMgr = swModel.SelectionManager
    swFeat = swSelMgr.GetSelectedObject6(1, 0)
    swSketch = swFeat.GetSpecificFeature2

    sketchConstrainedStatus = swSketch.GetConstrainedStatus
    print("    SketchConstrainedStatus: " + str(sketchConstrainedStatus))

    swModel.EditSketch
    vSketchSeg = swSketch.GetSketchSegments
    # print("    SketchSegCount: " + str(vSketchSeg.count))

    print(len(vSketchSeg))
    # for i in range(0, len(vSketchSeg)):
    # constraint = vSketchSeg[i].GetConstraints
    # GetId
    # for j, con in constraint:
    #     print("        SketchSegConstraint[" + str(i) + "] = " + str(con))
    # print("    Sketch: " + str(sketch))

    # sketchMgr = swModel.SketchManager

    # sketch = sketchMgr.GetSketchByName(sketchName)

    # if sketchType == "sketch":
    #     sketchSegs = sketch.SketchSegments
    # if sketchType == "profile":
    #     sketchSegs = sketch.ProfileSketchSegments
    return  # sketchSegs


def getAllPropertyNames(swModel):
    # returns a pandas dataframe of all properties and their values
    propertyMgr = swModel.Extension.CustomPropertyManager("")
    UseCached = False
    vNames = win32com.client.VARIANT(pythoncom.VT_VARIANT | pythoncom.VT_BYREF, None)
    vTypes = win32com.client.VARIANT(pythoncom.VT_VARIANT | pythoncom.VT_BYREF, None)
    vValues = win32com.client.VARIANT(pythoncom.VT_VARIANT | pythoncom.VT_BYREF, None)
    resolved = win32com.client.VARIANT(pythoncom.VT_VARIANT | pythoncom.VT_BYREF, None)
    linkedProps = win32com.client.VARIANT(
        pythoncom.VT_VARIANT | pythoncom.VT_BYREF, None
    )

    props = propertyMgr.GetAll3(
        vNames, vTypes, vValues, resolved, linkedProps
    )  # GetAllCustomProperties() #
    # print(str(props))
    # for i in range(0, len(vNames.value)):
    #     # print(vNames.value[i])
    #     print("  " + str(vNames.value[i]) + " | " + str(vTypes.value[i]) + " | " + str(vValues.value[i]))

    # returnTypes = []
    # for typeVal in vTypes.value:
    #     print(typeVal)
    #     typeVal = convertPropertyType(typeVal)
    # allNames = propertyMgr.GetNames #(vNames, vTypes, vValues)
    # for name in allNames:
    #     print("  " + str(name))

    # for i in range(0, len(vNames)):
    #     print("  " + str(vNames[i]) + " | " + str(vTypes[i]) + " | " + str(vValues[i]))

    # value = instance.Get6(FieldName, UseCached, ValOut, ResolvedValOut, WasResolved, LinkToProperty)
    # propVal = ValOut
    return vNames, vTypes, vValues


def getModelProperties(swModel):
    # returns a pandas dataframe of all properties and their values
    propertyMgr = swModel.Extension.CustomPropertyManager("")
    UseCached = False
    vNames = win32com.client.VARIANT(pythoncom.VT_VARIANT | pythoncom.VT_BYREF, None)
    vTypes = win32com.client.VARIANT(pythoncom.VT_VARIANT | pythoncom.VT_BYREF, None)
    vValues = win32com.client.VARIANT(pythoncom.VT_VARIANT | pythoncom.VT_BYREF, None)
    resolved = win32com.client.VARIANT(pythoncom.VT_VARIANT | pythoncom.VT_BYREF, None)
    linkedProps = win32com.client.VARIANT(
        pythoncom.VT_VARIANT | pythoncom.VT_BYREF, None
    )

    props = propertyMgr.GetAll3(
        vNames, vTypes, vValues, resolved, linkedProps
    )  # GetAllCustomProperties() #
    # print(str(props))
    # for i in range(0, len(vNames.value)):
    #     # print(vNames.value[i])
    #     print("  " + str(vNames.value[i]) + " | " + str(vTypes.value[i]) + " | " + str(vValues.value[i]))

    # returnTypes = []
    # for typeVal in vTypes.value:
    #     print(typeVal)
    #     typeVal = convertPropertyType(typeVal)
    # allNames = propertyMgr.GetNames #(vNames, vTypes, vValues)
    # for name in allNames:
    #     print("  " + str(name))

    # for i in range(0, len(vNames)):
    #     print("  " + str(vNames[i]) + " | " + str(vTypes[i]) + " | " + str(vValues[i]))

    # value = instance.Get6(FieldName, UseCached, ValOut, ResolvedValOut, WasResolved, LinkToProperty)
    # propVal = ValOut
    return vNames, vTypes, vValues


def getPropertyVal_byName(swModel, propName):
    propertyMgr = swModel.Extension.CustomPropertyManager("")
    UseCached = False
    vNames = var1
    vTypes = var1
    vValues = varVals
    ResolvedValOut = varVals
    WasResolved = var1
    LinkToProperty = var1
    propVal = propertyMgr.Get6(
        propName, UseCached, vValues, ResolvedValOut, varBool, varBool
    )
    print(vValues.value)
    # print(ResolvedValOut.value)
    print("  " + str(propVal) + "  |  " + str(vValues.value))
    return vValues.value


def convertPropertyType(propType):
    # swCustomInfoDate	64
    # swCustomInfoDouble	5
    # swCustomInfoNumber	3
    # swCustomInfoText	30
    # swCustomInfoUnknown	0
    # swCustomInfoYesOrNo	11
    if propType == "Date":
        propType = 64
    elif propType == "Double":
        propType = 5
    elif propType == "Number":
        propType = 3
    elif propType == "Text":
        propType = 30
    elif propType == "YesOrNo":
        propType = 11
    elif propType == 64:
        propType = "Date"
    elif propType == 5:
        propType = "Double"
    elif propType == 3:
        propType = "Number"
    elif propType == 30:
        propType = "Text"
    elif propType == 11:
        propType = "YesOrNo"
    else:
        raise Exception("Invalid property type: " + str(propType))
    return propType


def addProperty(swModel, propName, propType, propVal):

    print(
        "  Adding property: "
        + str(propName)
        + " | "
        + str(propType)
        + " | "
        + str(propVal)
    )
    # swCustomInfoDate	64
    # swCustomInfoDouble	5
    # swCustomInfoNumber	3
    # swCustomInfoText	30
    # swCustomInfoUnknown	0
    # swCustomInfoYesOrNo	11
    propType = convertPropertyType(propType)

    propertyMgr = swModel.Extension.CustomPropertyManager("")
    propertyMgr.Add3(propName, propType, propVal, 2)
    return


def selectPlaneByName(swModel, name):
    swModel.ClearSelection2(True)
    modelExt = swModel.Extension
    BoolStatus = modelExt.SelectByID2(name, "PLANE", 0, 0, 0, False, 0, arg1, 0)
    print("    Selecting feature: " + name + " | return val: " + str(BoolStatus))
    selectedItem = getSelectedItem(swModel)
    return selectedItem


def selectItemByName(swModel, name, itemType):
    swModel.ClearSelection2(True)
    modelExt = swModel.Extension
    BoolStatus = modelExt.SelectByID2(name, itemType, 0, 0, 0, False, 0, arg1, 0)
    # print("    Selecting feature: " + name + " | return val: " + str(BoolStatus))
    selectedItem = getSelectedItem(swModel)
    return selectedItem


def getSelectedItem(swModel):
    selMgr = swModel.SelectionManager
    selected = selMgr.GetSelectedObject6(1, 0).GetSpecificFeature2
    print("      Selected: " + str(selected.Name) + " |  ID: " + str(selected.GetID))
    return selected


def addSketch(swModel):
    sketchMgr = swModel.SketchManager
    # sketchMgr.InsertSketch(False)
    selMgr = swModel.SelectionManager
    planeName = selMgr.GetSelectedObject6(1, 0).GetSpecificFeature2
    sketchMgr.InsertSketch(True)
    swSketch = sketchMgr.ActiveSketch
    swSketch.Name = planeName.Name + "_TestSketch"
    print(" Created new sketch on: " + planeName.Name)
    return sketchMgr


def addSketchSegment(sketchMgr, sketchObj):
    return


def clearSelections(swModel):
    swModel.SketchManager.InsertSketch(False)
    swModel.ClearSelection2(True)  # (True)
    return swModel


def sketchPoint(sketchMgr_Obj, X, Y, Z, units):
    # creates a point in the sketch
    XConv = convertUnits(X, units)
    YConv = convertUnits(Y, units)
    ZConv = convertUnits(Z, units)
    sketchSegment = sketchMgr_Obj.CreatePoint(XConv, YConv, ZConv)
    # sketchMgr_Obj.InsertSketch(False)
    return sketchSegment


def sketchLine(sketchMgr_Obj, Xstart, Ystart, Zstart, Xend, Yend, Zend, units):
    # creates a line in the sketch
    XstartConv = convertUnits(Xstart, units)
    YstartConv = convertUnits(Ystart, units)
    ZstartConv = convertUnits(Zstart, units)
    XendConv = convertUnits(Xend, units)
    YendConv = convertUnits(Yend, units)
    ZendConv = convertUnits(Zend, units)
    sketchSegment = sketchMgr_Obj.CreateLine(
        XstartConv, YstartConv, ZstartConv, XendConv, YendConv, ZendConv
    )
    # sketchMgr_Obj.InsertSketch(False)
    return sketchSegment


def sketchLine_Center(sketchMgr_Obj, Xstart, Ystart, Zstart, Xend, Yend, Zend, units):
    # creates a line in the sketch
    XstartConv = convertUnits(Xstart, units)
    YstartConv = convertUnits(Ystart, units)
    ZstartConv = convertUnits(Zstart, units)
    XendConv = convertUnits(Xend, units)
    YendConv = convertUnits(Yend, units)
    ZendConv = convertUnits(Zend, units)
    sketchSegment = sketchMgr_Obj.CreateCenterLine(
        XstartConv, YstartConv, ZstartConv, XendConv, YendConv, ZendConv
    )
    # sketchMgr_Obj.InsertSketch(False)
    return sketchSegment


def sketchCircle(sketchMgr_Obj, Xcenter, Ycenter, Zcenter, radius, units):
    # creates a circle in the sketch
    distance = convertUnits(radius, units)
    Xconv = convertUnits(Xcenter, units)
    Yconv = convertUnits(Ycenter, units)
    Zconv = convertUnits(Zcenter, units)
    sketchSegment = sketchMgr_Obj.CreateCircle(Xconv, Yconv, Zconv, distance, 0, 0)
    # sketchMgr_Obj.InsertSketch(False)
    return sketchSegment


def addDimension(swModelObj, XdimLoc, YdimLoc, ZdimLoc, units):
    Xloc, Yloc, Zloc = converUnits_XYZ(XdimLoc, YdimLoc, ZdimLoc, units)
    swModelObj.AddDimension2(Xloc, Yloc, Zloc)
    return


def traverseFeaturesAndSubfeatures(swModel):
    print("Finding Features for: " + swModel.GetTitle)
    featureMgr = swModel.FeatureManager
    swFeat = swModel.FirstFeature
    # allFeats = featureMgr.GetFeatures(True)
    while swFeat == True:
        print(f"  Feature: {str(swFeat.Name)} ")
        traverseSubfeatures(swFeat)
        swFeat = swFeat.GetNextFeature


def traverseSubfeatures(swFeat):
    swSubFeat = swFeat.GetFirstSubFeature
    level = 2
    while swSubFeat == True:
        indentStr = "  " * level
        print(f"{indentStr}SubFeature: {str(swSubFeat.Name)} ")
        swDispDim = swSubFeat.GetFirstDisplayDimension
        while swDispDim == True:
            level = level + 1
            indentStr = "  " * level
            swAnn = swDispDim.GetAnnotation
            swDim = swDispDim.GetDimension
            print(
                "{}[ {} ] = {}".format(
                    indentStr, swDim.FullName, swDim.GetSystemValue2("")
                )
            )
            swDispDim = swSubFeat.GetNextDisplayDimension(swDispDim)
        swSubFeat = swSubFeat.GetNextSubFeature
    # for feat in allFeats:
    # if feat.GetTypeName2 == "ProfileFeature":
    #     swDef = feat.GetDefinition
    #     print(str(swDef))
    #     swSketch = swDef.
    # print("  Feature: {}  |  Type: {}  |  Definition: {}".format(str(feat.Name), str(feat.GetTypeName2) , str(feat.GetDefinition)))
    # print("  Feature: {}  |  Type: {} ".format(str(feat.Name), str(feat.GetTypeName2)))
    # print(str(allFeats))
    """swFeat = swModel.FirstFeature
    print("  "+ str(swFeat))
    while not swFeat == None:
        name = swFeat.Name
        swFeat = swFeat.GetNextFeature
        print("  " + name) """


def GetAllDimensions(vFeats):  # As Variant vFeats As Variant

    # Dim swDimsColl As Collection
    # Set swDimsColl = New Collection
    print("getting all display dimensions...\n")
    swDimsColl = []

    # Dim i As Integer

    for i in range(len(vFeats)):

        # Dim swFeat As SldWorks.Feature
        # Set swFeat = vFeats(i)
        swFeat = vFeats[i]
        print("Feature: " + swFeat.Name)
        # Dim swDispDim As SldWorks.DisplayDimension
        swDispDim = swFeat.GetFirstDisplayDimension
        print("  " + str(swDispDim))
        while not swDispDim is None:
            # If Not Contains(swDimsColl, swDispDim) Then
            if not swDispDim in swDimsColl:
                print("    " + str(swDispDim))
                # swDimsColl.Add swDispDim
                swDimsColl.append(swDispDim)
            # End If

            # Set swDispDim = swFeat.GetNextDisplayDimension(swDispDim)
            swDispDim = swFeat.GetNextDisplayDimension(swDispDim)

        #    If Not Contains(swDimsColl, swDispDim) Then
        #        swDimsColl.Add swDispDim
        #    End If

        #    Set swDispDim = swFeat.GetNextDisplayDimension(swDispDim)
        # Wend

    # Next

    # GetAllDimensions = CollectionToArray(swDimsColl)
    return swDimsColl


# # model = sw.ActiveDoc
# # modelExt = model.Extension
# # modelExt.SelectByID2("mysketch", "SKETCH", 0, 0, 0, False, 0, None, 0)

# # import win32com.client

# # app=win32com.client.Dispatch("SldWorks.Application")
# # doc=app.OpenDoc("c:\\Testpart.SLDPRT", 1)
# # doc.SaveAs2("c:\\Testpart.3dxml", 0, True, False)
