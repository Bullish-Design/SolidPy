import csv
import os
import random
import shutil

# import Excel_API as xl
import pythoncom

# import SW_API_FeatureParsing as SWparse
import SW_API_Functions as SWapi

# # swConst = win32com.client.gencache.EnsureModule('{4687F359-55D0-4CD3-B6CF-2EB42C11F989}', 0, 29, 0).constants # sw2015
# # swCmd = win32com.client.gencache.EnsureModule('{0AC1DE9F-3FBC-4C25-868D-7D4E9139CCE0}', 0, 29, 0).constants
# # # # {83A33D31-27C5-11CE-BFD4-00400513BB57}
# # arg1 = win32com.client.VARIANT(pythoncom.VT_DISPATCH, None)

sw = SWapi.win32com.client.Dispatch("SldWorks.Application")

# # sw = sldworks.ISldWorks(DispatchEx('SldWorks.Application'))

# # sw.Visible=1

# # insertPartName = 'testPart_forInsert2' #.SLDPRT'
# # insertPartLoc = r'C:\Users\alaureijs\Local Projects\Solidworks'

# # model = SWapi.createNewPart(sw, insertPartName, insertPartLoc)
# # model = sw.ActiveDoc


# # modelExt = model.Extension
# # selMgr = model.SelectionManager
# # featureMgr = model.FeatureManager
# # sketchMgr = model.SketchManager


# # -----------------------------------------------------------------------------------
# # turn into a series of 'template parts' that can be inserted into a new part and modified via dimension names
# # each part has mating features that can be used to mate to other parts (centres of faces, corners, etc)
# # tag drawing dimensions from imput sheet
# # -----------------------------------------------------------------------------------

# # title=model.GetTitle

# """
# fileDataSaveLoc = r"G:\My Drive\Google Drive - Projects\Code\NoteDB"
# workCSVdataSaveLoc = r"G:\My Drive\Google Drive - Work\MooreCo\swAPI"
# filePropertySaveName = "testFileProperties.csv"
# fileFeaturesSaveName = "testFileFeatures.csv"
# featureDefsSaveName = "testFeatureDefinitions.csv"
# CSVtemplateSaveLoc = workCSVdataSaveLoc
# SW_TemplateLoc = r"G:\My Drive\Google Drive - Work\MooreCo\swAPI"

# # print(title)
# circle1_radius = 5
# circle2_radius = 10
# units = "mm"

# """

units = "in"


class Model:

    # class for a solidworks model (part, assembly, drawing) - represents a wrapper around the SW object model?
    #     Or make it my own? I don't want to have to use the SW object model directly, but I also don't want to have to re-invent the wheel. <-- Brought to you by CoPilot
    # Provides access to get and set model properties
    # Provides I/o functions for model files - save, open, close, export as, create something from, add to asm, import part, etc.
    # Passed into the Feature class to get and set model features

    # Base everything off of a property within the solidworks file? Tie this ID to the model, can be autogenerated when autocreating a doc or user defined for existing files.

    def __init__(self, fileName="ActiveDoc", fileLocation=""):  # -> None:
        if fileName == "ActiveDoc":
            self.session = SWapi.win32com.client.Dispatch("SldWorks.Application")
            self.name = self.session.ActiveDoc.GetTitle
            self.templateLoc = fileLocation + "\\" + self.name
            isExist = os.path.exists(self.templateLoc)
            print("templateLoc = " + self.templateLoc)
            print(self.templateLoc)
            if not isExist:
                # Create a new directory because it does not exist
                os.makedirs(self.templateLoc)
            self.model = self.session.ActiveDoc
        else:
            try:
                self.session = SWapi.win32com.client.Dispatch("SldWorks.Application")
                self.name = fileName
                self.templateLoc = fileLocation + "\\" + self.name
                isExist = os.path.exists(self.templateLoc)
                print("templateLoc = " + self.templateLoc)
                if not isExist:
                    # Create a new directory because it does not exist
                    os.makedirs(self.templateLoc)
                swDocSpecification = self.session.GetOpenDocSpec(
                    self.templateLoc + "\\" + self.name
                )
                self.model = self.session.OpenDoc7(
                    swDocSpecification
                )  # OpenDoc6(self.templateLoc + '\\' + self.name, swConst.swDocPART, swConst.swOpenDocOptions_Silent, "", arg1, arg1)  #swConst.swDocumentTypes_e.swDocPART,
            except:
                print("File not found, creating new model")
                self.session = SWapi.win32com.client.Dispatch("SldWorks.Application")
                self.name = fileName
                self.templateLoc = fileLocation + "\\" + self.name
                isExist = os.path.exists(self.templateLoc)
                print("templateLoc = " + self.templateLoc)
                print(self.templateLoc)
                if not isExist:
                    # Create a new directory because it does not exist
                    os.makedirs(self.templateLoc)
                self.model = SWapi.createNewPart(self.session, fileName, fileLocation)

    def toggleVisible(self):
        # toggles the visibility of the model
        self.session.Visible = not self.session.Visible
        return self.session.Visible

    def getCSVfromModel(self):
        # gets all properties and features from the model and stores them in the class + CSV files
        getAllProperties_toCSV(self.model, self.templateLoc)  # , self.name)
        getAllFeatures_toCSV(self.model, self.templateLoc)  # , self.name)
        writeFeatureDefs_toCSV(self.model, self.templateLoc)  # , self.name)
        return  # self.model

    def updateModelFromCSV(self):
        # sets all properties and features from the class/CSV to the model
        addAllProperties_fromCSV(self.model, self.templateLoc)  # , self.name)
        # setAllFeatures_fromCSV(self.model, self.templateLoc) #, self.name)

        return

    def getModelProperties(self):
        # returns a pandas dataframe of all properties and their values
        propertyMgr = self.model.Extension.CustomPropertyManager("")
        UseCached = False
        vNames = SWapi.win32com.client.VARIANT(
            pythoncom.VT_VARIANT | pythoncom.VT_BYREF, None
        )
        vTypes = SWapi.win32com.client.VARIANT(
            pythoncom.VT_VARIANT | pythoncom.VT_BYREF, None
        )
        vValues = SWapi.win32com.client.VARIANT(
            pythoncom.VT_VARIANT | pythoncom.VT_BYREF, None
        )
        resolved = SWapi.win32com.client.VARIANT(
            pythoncom.VT_VARIANT | pythoncom.VT_BYREF, None
        )
        linkedProps = SWapi.win32com.client.VARIANT(
            pythoncom.VT_VARIANT | pythoncom.VT_BYREF, None
        )

        props = propertyMgr.GetAll3(
            vNames, vTypes, vValues, resolved, linkedProps
        )  # GetAllCustomProperties() #
        print(str(props))
        for i in range(0, len(vNames.value)):
            # print(vNames.value[i])
            print(
                "  "
                + str(vNames.value[i])
                + " | "
                + str(vTypes.value[i])
                + " | "
                + str(vValues.value[i])
            )


#         # returnTypes = []
#         # for typeVal in vTypes.value:
#         #     print(typeVal)
#         #     typeVal = convertPropertyType(typeVal)
#         # allNames = propertyMgr.GetNames #(vNames, vTypes, vValues)
#         # for name in allNames:
#         #     print("  " + str(name))

#         # for i in range(0, len(vNames)):
#         #     print("  " + str(vNames[i]) + " | " + str(vTypes[i]) + " | " + str(vValues[i]))

#         # value = instance.Get6(FieldName, UseCached, ValOut, ResolvedValOut, WasResolved, LinkToProperty)
#         # propVal = ValOut

#         return vNames, vTypes, vValues

#     # pass


def getAllProperties_toCSV(model, saveLoc, saveName="Properties"):
    vNames, vTypes, vValues = SWapi.getAllPropertyNames(model)
    with open(saveLoc + "\\" + saveName + ".csv", "w", newline="") as csvfile:
        writer = csv.writer(
            csvfile, delimiter=",", quotechar='"', quoting=csv.QUOTE_MINIMAL
        )
        writer.writerow(["Name", "Type", "Value"])
        for i in range(len(vNames.value)):
            writer.writerow(
                [
                    str(vNames.value[i]),
                    SWapi.convertPropertyType(vTypes.value[i]),
                    str(vValues.value[i]),
                ]
            )


def addAllProperties_fromCSV(model, saveLoc, saveName="Properties"):
    with open(saveLoc + "\\" + saveName + ".csv", newline="") as csvfile:
        reader = csv.reader(csvfile, delimiter=",", quotechar='"')
        next(reader)
        for row in reader:
            # if row[0] != 'Name':
            SWapi.addProperty(model, row[0].strip(), row[1].strip(), row[2].strip())


def getAllFeatures_toCSV(model, saveLoc, saveName="Features"):
    allFeaturesList = SWapi.traverseFeatures(model)
    for feature in allFeaturesList:
        if feature.GetTypeName2 == "ProfileFeature":

            # print()
            # SWapi.getAllSketchDims(model, feature.Name)
            # SWapi.selectItemByName(model, feature.Name, "SKETCH")
            # SWparse.ProcessSketch(sw, model) #,feature.Name)

            # selectItemByName(model, feature.Name, "BODYFEATURE")
            # SWparse.ProcessExtrude(sw, model)
            pass
    with open(saveLoc + "\\" + saveName + ".csv", "w", newline="") as csvfile:
        writer = csv.writer(
            csvfile, delimiter=",", quotechar='"', quoting=csv.QUOTE_MINIMAL
        )
        writer.writerow(["ID", "Name", "Type"])  # , 'Value'])
        for i in range(len(allFeaturesList)):
            writer.writerow(
                [
                    allFeaturesList[i].GetID,
                    allFeaturesList[i].Name,
                    allFeaturesList[i].GetTypeName2,
                ]
            )
            # writeFeatureDefs_toCSV(model, allFeaturesList[i], saveLoc, 'Feature_Definitions.csv')

            # writer.writerow([allFeaturesList[i].GetID, allFeaturesList[i].Name, allFeaturesList[i].GetTypeName2])
            # if allFeaturesList[i].GetTypeName2 == "RefPlane":
            #     if allFeaturesList[i].Name != "Top Plane" and allFeaturesList[i].Name != "Front Plane" and allFeaturesList[i].Name != "Right Plane":
            #         print("Plane: " + allFeaturesList[i].Name + "  |  Type: " + allFeaturesList[i].GetTypeName2)
            #         planeProps = SWapi.Plane(model, allFeaturesList[i].Name, "", 0, units).getProperties()
            #         refPlaneRow = ['FeatDef',{'FeatID':allFeaturesList[i].GetID}]
            #         refPlaneRow.extend(planeProps)
            #         writer.writerow(refPlaneRow)
            # if allFeaturesList[i].GetTypeName2 == "ProfileFeature":
            #     print("Sketch: " + allFeaturesList[i].Name + "  |  Type: " + allFeaturesList[i].GetTypeName2)
            #     dimList = SWapi.getAllSketchDims(model, allFeaturesList[i].Name)
            #     for dim in dimList:
            #         dimRow = ['DimVal', {'FeatID':allFeaturesList[i].GetID}]
            #         dimRow.append(dim)
            #         writer.writerow(dimRow)
    return


def updateFeatures_fromCSV(model, saveLoc, saveName="Features"):
    allFeaturesList = SWapi.traverseFeatures(model)
    with open(saveLoc + "\\" + saveName + ".csv", newline="") as csvfile:
        # Iterate through the CSV file and compare features by ID. Create new features if they don't exist.
        reader = csv.reader(csvfile, delimiter=",", quotechar='"')
        next(reader)
        for row in reader:
            pass
    pass


def writeFeatureDefs_toCSV(model, saveLoc, saveName="Feature_Definitions"):
    allFeaturesList = SWapi.traverseFeatures(model)
    with open(saveLoc + "\\" + saveName + ".csv", "w", newline="") as csvfile:
        writer = csv.writer(
            csvfile, delimiter=",", quotechar='"', quoting=csv.QUOTE_MINIMAL
        )
        writer.writerow(["FeatID", "DimName", "DimVal"])
    for i in range(len(allFeaturesList)):
        featDict = allFeaturesList[i]
        featType = featDict.GetTypeName2
        if featType == "RefPlane":
            if (
                featDict.Name != "Top Plane"
                and featDict.Name != "Front Plane"
                and featDict.Name != "Right Plane"
            ):
                planeProps = SWapi.Plane(
                    model, featDict.Name, "", 0, units
                ).getProperties()
                print(planeProps)
                if not planeProps == None:
                    refPlaneRow = [featDict.GetID]
                    refPlaneRow.extend(planeProps)
                    with open(
                        saveLoc + "\\" + saveName + ".csv", "a", newline=""
                    ) as csvfile:
                        writer = csv.writer(
                            csvfile,
                            delimiter=",",
                            quotechar='"',
                            quoting=csv.QUOTE_MINIMAL,
                        )
                        writer.writerow(refPlaneRow)
        if featType == "ProfileFeature":
            dimList = SWapi.getAllSketchDims(model, featDict.Name)
            for dim in dimList:
                dimRow = [featDict.GetID]
                dimRow.append(dim)
                with open(
                    saveLoc + "\\" + saveName + ".csv", "a", newline=""
                ) as csvfile:
                    writer = csv.writer(
                        csvfile, delimiter=",", quotechar='"', quoting=csv.QUOTE_MINIMAL
                    )
                    writer.writerow(dimRow)
                # writer.writerow(dimRow)
        if featType == "ICE":
            print("ICE")

    return


def writeGenericFeatDefs_toCSV(model, saveLoc, saveName="Feature_Defs"):
    allFeaturesList = SWapi.traverseFeatures(model)
    with open(saveLoc + "\\" + saveName + ".csv", "w", newline="") as csvfile:
        writer = csv.writer(
            csvfile, delimiter=",", quotechar='"', quoting=csv.QUOTE_MINIMAL
        )
        writer.writerow(["Feature Type", "FeatID", "DimName", "DimVal"])

    for i in range(len(allFeaturesList)):
        rowData = []
        featDict = allFeaturesList[i]
        featType = featDict.GetTypeName2
        rowData.append(featType)
        featID = featDict.GetID
        rowData.append(featID)
        print("Feature ID: " + str(featID) + "  |  Feature Type: " + str(featType))
        feature = model.FeatureByID(featID)
        print(str(feature))
        if feature.getDefinition != None:  # SWapi.varNone:
            defList = feature.getDefinition
            print(defList)
            for defItem in defList:
                print(str(defItem))
                rowData.append(defItem)
            with open(saveLoc + "\\" + saveName + ".csv", "a", newline="") as csvfile:
                writer = csv.writer(
                    csvfile, delimiter=",", quotechar='"', quoting=csv.QUOTE_MINIMAL
                )
                writer.writerow(rowData)

    return


# def getAllFeatures_fromExcel(fileDataSaveLoc, fileFeaturesSaveName):
#     swTemplate_name = fileDataSaveLoc + "\\" + fileFeaturesSaveName
#     xlTemplate = xl.xw.Book(swTemplate_name)
#     propSheet = xlTemplate.sheets["FileProperties"]
#     featSheet = xlTemplate.sheets["PartFeatures"]
#     planeSheet = xlTemplate.sheets["PartPlanes"]
#     sketchSheet = xlTemplate.sheets["PartSketches"]
#     fileProps = xl.getDF_fromSheet(propSheet)
#     print(fileProps)
#     allFeatures = xl.getDF_fromSheet(featSheet)
#     planes = xl.getDF_fromSheet(planeSheet, "B1")
#     print(planes)
#     sketches = xl.getDF_fromSheet(sketchSheet)
# def getAllFeatures_fromExcel(fileDataSaveLoc, fileFeaturesSaveName):
#     swTemplate_name = fileDataSaveLoc + "\\" + fileFeaturesSaveName
#     xlTemplate = xl.xw.Book(swTemplate_name)
#     propSheet = xlTemplate.sheets["FileProperties"]
#     featSheet = xlTemplate.sheets["PartFeatures"]
#     planeSheet = xlTemplate.sheets["PartPlanes"]
#     sketchSheet = xlTemplate.sheets["PartSketches"]
#     fileProps = xl.getDF_fromSheet(propSheet)
#     print(fileProps)
#     allFeatures = xl.getDF_fromSheet(featSheet)
#     planes = xl.getDF_fromSheet(planeSheet, "B1")
#     print(planes)
#     sketches = xl.getDF_fromSheet(sketchSheet)

#     return fileProps, allFeatures, planes, sketches
#     return fileProps, allFeatures, planes, sketches


def getAllFeature_toExcel(model, fileDataSaveLoc, fileFeaturesSaveName):
    return


def updateFeatures_fromExcel(model, fileDataSaveLoc, fileFeaturesSaveName):

    return


# def newPart_init(partName, csvFolderSaveLoc=CSVtemplateSaveLoc):
#     # Create new part templates and save them to the specified location.
#     # If no location is specified, save inside a new folder in the default location.

#     # copy and rename the template folder to the new part name:
#     if not os.path.exists(csvFolderSaveLoc + "\\" + partName):
#         shutil.copytree(CSVtemplateSaveLoc, csvFolderSaveLoc + "\\" + partName)
#         os.rename(
#             csvFolderSaveLoc + "\\" + partName, csvFolderSaveLoc + "\\" + partName
#         )
#     # copy and rename the template folder to the new part name:
#     if not os.path.exists(csvFolderSaveLoc + "\\" + partName):
#         shutil.copytree(CSVtemplateSaveLoc, csvFolderSaveLoc + "\\" + partName)
#         os.rename(
#             csvFolderSaveLoc + "\\" + partName, csvFolderSaveLoc + "\\" + partName
#         )

#     return
#     return

# """
# insertPartName = "testCircDesk"  # .SLDPRT'
# insertPartLoc = r"C:\Users\alaureijs\Local Projects\Solidworks"

# fileDataSaveLoc = r"G:\My Drive\Google Drive - Projects\Code\NoteDB"
# fileDataSaveLoc = (
#     r"C:\Users\alaureijs\OneDrive - moorecoinc\Projects\Automation\Solidworks"
# )
# fileFeaturesSaveName = "SW Part Template.csv"  # xlsm'
# partTemplateFolder = (
#     r"G:\My Drive\Google Drive - Work\MooreCo\swAPI"  # \SW_Part_Template
# )

# # model = sw.ActiveDoc #SWapi.createNewPart(sw, insertPartName, insertPartLoc)


# testModel = Model()  # activeDocName, activeDocLoc)
# print(testModel.name)
# print(testModel.templateLoc)
# print(testModel.model)
# print(testModel.session)
# # testModel.getCSVfromModel()
# # testModel.getModelProperties()
# testModel.updateModelFromCSV()
# testModel = Model()  # activeDocName, activeDocLoc)
# print(testModel.name)
# print(testModel.templateLoc)
# print(testModel.model)
# print(testModel.session)
# # testModel.getCSVfromModel()
# # testModel.getModelProperties()
# testModel.updateModelFromCSV()
# testModel.getCSVfromModel()
# writeGenericFeatDefs_toCSV(testModel.model, testModel.templateLoc, "Generic_Feat_Defs")
#  """

# # fileProps, allFeatures, planes, sketches = getAllFeatures_fromExcel(fileDataSaveLoc, fileFeaturesSaveName)

# # print("Props")
# # print(fileProps)
# # print("Feats")
# # print( allFeatures)
# # print("planes")
# # print( planes)
# # print("sketches")
# # print( sketches)

# # getAllProperties_toCSV(model, workCSVdataSaveLoc, fileFeaturesSaveName)


# ##
# # getAllFeatures_toCSV(model, workCSVdataSaveLoc, fileFeaturesSaveName)
# ##

# # features = SWapi.traverseFeatures(model)

# # print(features)

# # SWapi.GetAllDimensions(features)

# print("\n\nDone.")
# print("\n\nDone.")


def createPlanes(model, planeProps, units, edit=False):
    # print(planeProps)
    print("Creating planes...")
    print(len(planeProps))
    for i in range(len(planeProps)):
        print(i)
        # i=i-1
        lookupval = i
        lookupval = planeProps.iloc[:, :0].iloc[lookupval].name
        print(lookupval)
        planeName = lookupval  # planeProps.loc[lookupval, 'PlaneName']
        print("  " + planeName)

        # planeName = planeProps.loc[lookupval, 'PlaneName']
        # print("  " + planeName)
        planeOffsetFrom = planeProps.loc[lookupval, "OffsetFrom"]
        print("  " + planeOffsetFrom)
        planeOffsetDist = planeProps.loc[lookupval, "OffsetDistance"]
        print("  " + str(planeOffsetDist))

        # planeType = planeProps['Type'][i]
        plane = SWapi.Plane(model, planeName, planeOffsetFrom, planeOffsetDist, units)
        if edit == True:
            plane.edit()
            print(
                f"Edited '{planeName}' to '{planeOffsetDist}' {units} from '{planeOffsetFrom}'"
            )
        else:
            plane.create()
            print(
                f"Created '{planeName}' @'{planeOffsetDist}' {units} from '{planeOffsetFrom}'"
            )
    return


# # print(fileProps.loc['Units', 'Value'])
# # units = str(fileProps.loc['Units', 'Value'])
# # #print(planes)

# # createPlanes(model, planes, units) #, edit=True)


# def createRandomSkeleton():
#     planesToCreate = [5, 10, 15, 20, 25, 30]
#     planeToOffset = ["Right Plane", "Front Plane", "Top Plane"]
#     constraintDirection = ["Horizontal", "Vertical"]
#     for plane in planesToCreate:
#         planeName = "OffsetPlane_" + str(plane)
#         insertPlane = SWapi.Plane(
#             model, planeName, random.choice(planeToOffset), plane, units
#         )
#         insertPlane.create()
#         sketchToCreate = SWapi.Sketch(
#             model, "Sketch_" + planeName, insertPlane.refPlane
#         )
#         sketchToCreate.create()
#         # print(sketchToCreate.name)
#         # SWapi.clearSelections(model)
#         print("Created " + sketchToCreate.name + " on " + insertPlane.refPlane)
#         line1 = SWapi.Line(model, sketchToCreate.name, "Line_Test" + str(plane))
#         print("creating line: " + line1.name)
#         lineConstraint = random.choice(constraintDirection)
#         lineDist = random.choice(planesToCreate)
#         print(
#             "Line constraint: " + lineConstraint + "  |  Line length: " + str(lineDist)
#         )
#         line = line1.create(lineConstraint, lineDist, "mm")
# def createRandomSkeleton():
#     planesToCreate = [5, 10, 15, 20, 25, 30]
#     planeToOffset = ["Right Plane", "Front Plane", "Top Plane"]
#     constraintDirection = ["Horizontal", "Vertical"]
#     for plane in planesToCreate:
#         planeName = "OffsetPlane_" + str(plane)
#         insertPlane = SWapi.Plane(
#             model, planeName, random.choice(planeToOffset), plane, units
#         )
#         insertPlane.create()
#         sketchToCreate = SWapi.Sketch(
#             model, "Sketch_" + planeName, insertPlane.refPlane
#         )
#         sketchToCreate.create()
#         # print(sketchToCreate.name)
#         # SWapi.clearSelections(model)
#         print("Created " + sketchToCreate.name + " on " + insertPlane.refPlane)
#         line1 = SWapi.Line(model, sketchToCreate.name, "Line_Test" + str(plane))
#         print("creating line: " + line1.name)
#         lineConstraint = random.choice(constraintDirection)
#         lineDist = random.choice(planesToCreate)
#         print(
#             "Line constraint: " + lineConstraint + "  |  Line length: " + str(lineDist)
#         )
#         line = line1.create(lineConstraint, lineDist, "mm")


# # drw1 = SWapi.DRW(sw)
# # drw1.create()
# # createRandomSkeleton()

# # skeletonPart = SWapi.insertPart(model, insertPartLoc + "\\" + insertPartName + ".SLDPRT")


# # featureList = ['Overall_Footprint_Worksurface', 'Worksurface_Sketch', ]


# # Ignore full feature extraction for now. Just build parts from planes/sketches with named dimensions.
# # Then extract any named dimensions from all features and save/update them to/from a CSV file.

# # File should be pulled from a skeleton TDD sketch file anyway, then parts are just a series of extrudes.

# # Things to figure out:
# #   Midpoint relations
# #   Extrudes
# #   Revolves
# #   Cuts
# #   Create Drawing
# #   Create Assembly
# #   Import checked dimensions to drawing


# # OffsetPlane1 = SWapi.Plane(model, "OffsetPlane_1", "Right Plane", 10, units)
# # #TopPlane = SWapi.Plane(model, "TopPlane_Copy", "Top Plane", 0, units)
# # #OffsetPlane1.create()

# # print(OffsetPlane1.offsetDistance)
# # print(OffsetPlane1.refPlane)
# # #OffsetPlane1.offsetDistance = 20

# # OffsetPlane1.getProperties()
# # print(OffsetPlane1.offsetDistance)
# # print(OffsetPlane1.refPlane)

# # SWapi.create_RefAxis(model, "OffsetPlane_1", "Front Plane")
# # SWapi.clearSelections
# # sketch1 = SWapi.Sketch(model, "TestSketch", "OffsetPlane_1").create()

# # addAllProperties_fromCSV(model, fileDataSaveLoc, filePropertySaveName)
# # getAllProperties_toCSV(model, fileDataSaveLoc, filePropertySaveName)
# # getAllFeatures_toCSV(model, fileDataSaveLoc, fileFeaturesSaveName)

# # dimlist = SWapi.getAllSketchDims(model, "Sketch1")

# # propertyToAdd = [['TestProp1', 'Date', "4-13-59"],['TestProp2', 'Number', 666],['TestProp3', 'Text','This is a test']]
# # for prop in propertyToAdd:
# #     SWapi.addProperty(model, prop[0], prop[1], prop[2])

# # SWapi.getSelectedItem(model)

# # SWapi.createPlane_OffsetDistance(model, "Front Plane", 10, units)
# # vNames, vTypes, vValues = SWapi.getAllPropertyNames(model)

# # for i in range(len(vNames.value)):
# #     print("  " + str(vNames.value[i]) + " | " + SWapi.convertPropertyType(vTypes.value[i]) + " | " + str(vValues.value[i]))

# # for name in vNames:
# #     print(str(name.value, ))
# #     print(SWapi.getPropertyVal_byName(model, name))
# #     #print(SWapi.getPropertyType(model, name))
# #     #print(SWapi.getPropertyDescription(model, name))


# # #SWapi.traverseFeatures(model)
# # selected1 = SWapi.selectPlaneByName(model, "Front Plane")
# # newSketch = SWapi.addSketch(model) #, "Front Plane")
# # #print(newSketch)
# # SWapi.toggleDimInputBox(sw, False)
# # newCircle = SWapi.sketchCircle(newSketch, 0, 0, 0, circle1_radius, units)
# # SWapi.addDimension(model, circle1_radius*1.5, circle1_radius*1.5, circle2_radius*1.5, units)
# # model = SWapi.clearSelections(model)

# # model2 = sw.ActiveDoc

# # selected = SWapi.selectPlaneByName(model2, "Right Plane")
# # print(selected.Name)
# # newSketch2 = SWapi.addSketch(model2) #, "Right Plane")
# # newCircle2 = SWapi.sketchCircle(newSketch2, 2, 2, 2, circle2_radius, units)
# # SWapi.addDimension(model2, circle2_radius*1.5, circle2_radius*1.5, circle2_radius*1.5, units)
# # newCircle2 = SWapi.sketchCircle(newSketch2, 0, 5, 0, circle2_radius+1, units)
# # SWapi.addDimension(model2, circle2_radius*1.5, circle2_radius*1.5, circle2_radius*1.5, units)

# # model2 = SWapi.clearSelections(model2)
# # SWapi.toggleDimInputBox(sw, True)
# # #model2 = SWapi.clearSelections(model2)
# # # model = sw.model
# # #selMgr = model.SelectionManager
# # #print(type(selMgr))
# # #aSketch = selMgr.GetSelectedObject(1) #.GetSpecificFeature2
# # #aSketch.Name
