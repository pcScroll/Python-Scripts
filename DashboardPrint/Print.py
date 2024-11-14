# -*- coding: cp1256 -*-
import arcpy
import datetime
import uuid
import os
import zipfile
import shutil
import json,ast
import string
import operator
##import sys
##reload(sys)
##sys.setdefaultencoding('utf8')
#===========================================================================================================
arcpy.env.overwriteOutput = True
#===========================================================================================================
arcpy.AddMessage("Collecting input parameters")
#-----------------------------------------------------------------------------------------------------------
inFile = arcpy.GetParameterAsText(0)
#===========================================================================================================
arcpy.AddMessage("Loading MapDocument")
#-----------------------------------------------------------------------------------------------------------
#----------------------------------------------------------------------------------------------------------- 
# Create a folder in the scratch directory to extract zip to
#path = r"\\AD007APP06\gisapplications\arcgisdata1061\Dashboard\ArchivePath"
#zipFolder = os.path.join(arcpy.env.scratchFolder, "zipContents")
#zipFolderpath = r"\\AD007APP06\gisapplications\arcgisdata1061\GPscripts\DataExtraction\data"
#zipname =datetime.datetime.now().strftime("zip%Y%m%d%H%M%S")
zipFolder = r"\\ap007fsc05fs\GISApplications\agsadmgis\arcgisdata\DashboardPrint\zipContents"
#zipFolder = zipFolderpath + "\\" + zipname
isdir = os.path.isdir(zipFolder)
if isdir == False:
    os.mkdir(zipFolder)  
#-----------------------------------------------------------------------------------------------------------  
#Extract the zip contents  
zip2Extract = zipfile.ZipFile(inFile, 'r')  
zip2Extract.extractall(zipFolder)  
zip2Extract.close()
#===========================================================================================================
arcpy.AddMessage("Collecting result parameters")
#-----------------------------------------------------------------------------------------------------------
jsonval = []
jsonresults1 = []

MunicipalCenter = None
Zone = None
Sector = None
PlannedCategoryType = None
PlannedUsage = None
ConstructionStatus = None

Sector_count = None
PLOT_Count = None
POP_Count = None
TWATEEQUNITS_Count = None
TWATEEQCONTRACTS_Count = None
DEDLICENSES_Comm = None
DEDLICENSES_Ind = None
outseatingpermit_Count = None
foodtruckpermit_Count = None
TWATEEQUNITS_Res = None
TWATEEQUNITS_Comm = None
TWATEEQUNITS_Ind = None
NBMBulding = None
BDMSBulding = None
AsBulidMisc = None
AsBulidMain = None
AsPlannedMisc = None
AsPlannedMain = None
TWATEEQContract_Res = None
TWATEEQContract_Comm = None
TWATEEQContract_Ind = None


lang = None

for filename in os.listdir(zipFolder):
    if filename.endswith('json'):
        baseconfigjsonname = os.path.splitext(filename)[0]
        if baseconfigjsonname == "config":
            with open(zipFolder + "\\" + filename) as json_File:
                data = json.load(json_File)
                lang = data['lang']
                
                for p in data['inputs']:
                    a = p['label']
                    label = a.encode('utf-8')
                    if label == "Municipal Center":
                        b = p['value']
                        MunicipalCenter = b.encode('utf-8')
                        
                    if label == "مراكز البلدية":
                        b = p['value']
                        MunicipalCenter = b.encode('utf-8')
                       
                    if label == "Zone":
                        b = p['value']
                        Zone = b.encode('utf-8')
                    if label == "المنطقة":
                        b = p['value']
                        Zone = p['value']
                    if label == "Planned Category Type":
                        b = p['value']
                        PlannedCategoryType = b.encode('utf-8')
                    if label == "الاستخدام الرئيسي":
                        b = p['value']
                        PlannedCategoryType = b.encode('utf-8')
                    if label == "Planned Usage":
                        b = p['value']
                        PlannedUsage =  b.encode('utf-8')
                    if label == "الاستخدام الفرعي":
                        b = p['value']
                        PlannedUsage =  b.encode('utf-8')
                    if label == "Construction Status":
                        b = p['value']
                        ConstructionStatus = b.encode('utf-8')
                    if label == "حالة البناء":
                        b = p['value']
                        ConstructionStatus = b.encode('utf-8')
                    if label == "Sector":
                        b = p['value']
                        Sector = b.encode('utf-8')
                    if label == "الحوض":
                        b = p['value']
                        Sector = b.encode('utf-8')
                for p in data['results1']:
                    a = p['label']
                    label = a.encode('utf-8')
                    if label == "سكني":
                        b = p['value']
                        TWATEEQContract_Res = label + ": "+ b.encode('utf-8')
                    if label == "Residential":
                        b = p['value']
                        TWATEEQContract_Res = label + ": "+ b.encode('utf-8')
                    if label == "تجاري":
                        b = p['value']
                        TWATEEQContract_Comm = label + ": "+ b.encode('utf-8')
                    if label == "Commertial":
                        b = p['value']
                        TWATEEQContract_Comm = label + ": "+ b.encode('utf-8')
                    if label == "صناعي":
                        b = p['value']
                        TWATEEQContract_Ind = label + ": "+ b.encode('utf-8')
                    if label == "Industrial":
                        b = p['value']
                        TWATEEQContract_Ind = label + ": "+ b.encode('utf-8')
                    if label == "As Bulid Misc":
                        b = p['value']
                        AsBulidMisc = label + ": "+ b.encode('utf-8')
                    if label == "As Bulid Main":
                        b = p['value']
                        AsBulidMain = label + ": "+ b.encode('utf-8')
                    if label == "As Planned Misc":
                        b = p['value']
                        AsPlannedMisc = label + ": "+ b.encode('utf-8')
                    if label == "As Planned Main":
                        b = p['value']
                        AsPlannedMain = label + ": "+ b.encode('utf-8')
                    if label == "حسب المبنى متفرقات":
                        b = p['value']
                        AsBulidMisc = label + ": "+ b.encode('utf-8')
                    if label == "حسب المبنى الرئيسي":
                        b = p['value']
                        AsBulidMain = label + ": "+ b.encode('utf-8')
                    if label == "حسب المخطط متفرقات":
                        b = p['value']
                        AsPlannedMisc = label + ": "+ b.encode('utf-8')
                    if label == "حسب المخطط الرئيسي":
                        b = p['value']
                        AsPlannedMain = label + ": "+ b.encode('utf-8')
                for p in data['results2']:
                    a = p['label']
                    label = a.encode('utf-8')
                    if label == "سكني":
                        b = p['value']
                        TWATEEQUNITS_Res = label + ": "+ b.encode('utf-8')
                    if label == "Residential":
                        b = p['value']
                        TWATEEQUNITS_Res = label + ": "+ b.encode('utf-8')
                    if label == "تجاري":
                        b = p['value']
                        TWATEEQUNITS_Comm = label + ": "+ b.encode('utf-8')
                    if label == "Commertial":
                        b = p['value']
                        TWATEEQUNITS_Comm = label + ": "+ b.encode('utf-8')
                    if label == "صناعي":
                        b = p['value']
                        TWATEEQUNITS_Ind = label + ": "+ b.encode('utf-8')
                    if label == "Industrial":
                        b = p['value']
                        TWATEEQUNITS_Ind = label + ": "+ b.encode('utf-8')
                for s in data['results']:
                    for key in sorted(s):
                        name = (key['field']['name'])
                        val = (key['value']['val'])
                        #val = val.replace(" ", "")
                        a = name.encode('utf-8')
                        if a == "2019 / 2016 عدد السكان":
                            result = val.split('/')
                            result1 = result[0]
                            result1 = result1.strip()
                            result2 = result[1]
                            result2 = result2.strip()
                            res = result2 + "/" + result1
                            POP_Count = name + "\n" + res
                            arcpy.AddMessage(POP_Count)
                        elif a == "POPULATION 2019 / 2016":
                            result = val.split('/')
                            result1 = result[0]
                            result1 = result1.strip()
                            result2 = result[1]
                            result2 = result2.strip()
                            res = result1 + "/" + result2
                            POP_Count = name + "\n" + res
                        elif a == "الحوض":
                            Sector_count = name + ": " + val
                        elif a == "Sector":
                            Sector_count = "Sector NO" + ": " + val
                        elif a == "الأرض":
                            PLOT_Count = name + ": " + val
                        elif a == "PLOTS":
                            PLOT_Count = "PLOT NO" + ": " + val
                        elif a == "وحدات توثيق":
                            TWATEEQUNITS_Count = name + ": " + val
                        elif a == "TAWTEEQ UNITS":
                            TWATEEQUNITS_Count = name + ": " + val
                        elif a == "عقود توثيق":
                            TWATEEQCONTRACTS_Count = name + ": " + val
                        elif a == "TAWTEEQ CONTRACTS":
                            TWATEEQCONTRACTS_Count = "TWATEEQ ACTIVE CONTRACTS" + ": " + val
                        elif a == "Out Seating Permit":
                            outseatingpermit_Count = name + ": " + val
                        elif a == "تصاريح الجلسات الخارجية":
                            outseatingpermit_Count = name + ": " + val
                        elif a == "العربات المتنقلة":
                            foodtruckpermit_Count = name + ":" + val
                        elif a == "Food Truck Permit":
                            foodtruckpermit_Count = name + ": " + val
                        elif a == "تراخيص دائرة التنمية الاقتصادية الصناعية":
                            DEDLICENSES_Comm = name + ": " + val
                        elif a == "تراخيص دائرة التنمية الاقتصادية التجارية":
                            DEDLICENSES_Ind = name + ": " + val
                        elif a == "DED license Commercial":
                            DEDLICENSES_Comm = name + ": " + val
                        elif a == "DED license Industrial":
                            DEDLICENSES_Ind = name + ": " + val
                        elif a == "NBM Bulding":
                            NBMBulding = name + ": " + val
                        elif a == "خارطة الأساس للمباني":
                            NBMBulding = name + ": " + val
                        elif a == "BDMS Bulding":
                            BDMSBulding = name + ": " + val
                        elif a == "نظام إدارة بيانات المباني":
                            BDMSBulding = name + ": " + val
                        else:
                            nameval = name + ": " + val
                            jsonval.insert(len(jsonval),[nameval])
                        
#-----------------------------------------------------------------------------------------------------------
if lang == "AR":
    DocumentPath = r"\\ap007fsc05fs\GISApplications\agsadmgis\arcgisdata\DashboardPrint\A3_Template_AR.mxd"
if lang == "EN":
    DocumentPath = r"\\ap007fsc05fs\GISApplications\agsadmgis\arcgisdata\DashboardPrint\A3_Template_EN.mxd"

mxd = arcpy.mapping.MapDocument(DocumentPath)
pDataFrame = arcpy.mapping.ListDataFrames(mxd, "Layers")[0]
pSpatialReference = pDataFrame.spatialReference                      
#-----------------------------------------------------------------------------------------------------------                 
PlotsTxtElems = arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT", "PLOT_Count_")[0]

jsoncount = len(jsonval) 

##positonx = 5.301
##positiony = 27.298
##elmwidth = 4.6
##elmheight = 0.6015

positonx = PlotsTxtElems.elementPositionX
positiony = PlotsTxtElems.elementPositionY
elmwidth = PlotsTxtElems.elementWidth
elmheight = PlotsTxtElems.elementHeight
 
if jsoncount >=2:
    x = PlotsTxtElems.clone()
    x.elementPositionX = positonx + elmwidth
    positonx = x.elementPositionX
    x.text ="test"

count = 2
if jsoncount >=3:
    for x in range(jsoncount):
        y = None
        if count <> jsoncount:
            y = PlotsTxtElems.clone()
            y.elementPositionY = positiony - elmheight
            y.text ="test"
            count += 1
        if count <> jsoncount:
            y = PlotsTxtElems.clone()
            #y.elementPositionX = 9.901
            y.elementPositionX = positonx
            y.elementPositionY = positiony - elmheight
            y.text ="test"
            count += 1
        if y <> None:
            positiony = y.elementPositionY

textelm = arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT", "PLOT_Count_*")
count = 0
jsonval.sort(key=lambda x: x[0])
jsonval.sort(reverse = True)
for elm in textelm:
    if count <> jsoncount:
        fieldval = jsonval[count][0]
        a = fieldval.encode('utf-8')
        #elm.text = (str(jsonval[count][0]))
        elm.text = a
        #elm.fontSize = 10
        count += 1
#==============================================================================================================
##PlotsTxtElems1 = arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT", "Test_Count_")[0]
##positonx = PlotsTxtElems1.elementPositionX
##positiony = PlotsTxtElems1.elementPositionY
##elmwidth = PlotsTxtElems1.elementWidth
##elmheight = PlotsTxtElems1.elementHeight
##
##jsonresults1count = len(jsonresults1)
##arcpy.AddMessage(jsonresults1count)
##if jsonresults1count >=2:
##    x = PlotsTxtElems1.clone()
##    x.elementPositionX = positonx + elmwidth
##    positonx = x.elementPositionX
##    x.text ="test"
##
##count = 2
##if jsonresults1count >=3:
##    for x in range(jsonresults1count):
##        y = None
##        if count <> jsonresults1count:
##            y = PlotsTxtElems1.clone()
##            y.elementPositionY = positiony - elmheight
##            y.text ="test"
##            count += 1
##        if count <> jsonresults1count:
##            y = PlotsTxtElems1.clone()
##            #y.elementPositionX = 9.901
##            y.elementPositionX = positonx
##            y.elementPositionY = positiony - elmheight
##            y.text ="test"
##            count += 1
##        if y <> None:
##            positiony = y.elementPositionY
##textelm = arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT", "Test_Count_*")
##count = 0
##for elm in sorted(textelm):
##    if count <> jsonresults1count:
##        fieldval = jsonresults1[count][0]
##        a = fieldval.encode('utf-8')
##        #elm.text = (str(jsonval[count][0]))
##        elm.text = a
##        #elm.fontSize = 10
##        count += 1
#-----------------------------------------------------------------------------------------------------------
for elm in arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT"):
    if elm.name == "MunicipalCenter":
        if MunicipalCenter is None: elm.text = " "
        elif MunicipalCenter == "": elm.text = " "
        else: elm.text = MunicipalCenter
    if elm.name == "Zone":
        if Zone is None: elm.text = " "
        elif Zone == "": elm.text = " "
        else: elm.text = Zone
    if elm.name == "Sector":
        if Sector is None: elm.text = " "
        elif Sector == "": elm.text = " "
        else: elm.text = Sector
    if elm.name == "PlannedCategoryType":
        if PlannedCategoryType is None: elm.text = " "
        elif PlannedCategoryType == "": elm.text = " "
        else: elm.text = PlannedCategoryType
    if elm.name == "PlannedUsage":
        if PlannedUsage is None: elm.text = " "
        elif PlannedUsage == "": elm.text = " "
        else: elm.text = PlannedUsage
    if elm.name == "ConstructionStatus":
        if ConstructionStatus is None: elm.text = " "
        elif ConstructionStatus == "": elm.text = " "
        else: elm.text = ConstructionStatus
    if elm.name == "Sector_count":
        if Sector_count is None: elm.text = " "
        else: elm.text = Sector_count
    if elm.name == "PLOT_Count":
        if PLOT_Count is None: elm.text = " "
        else: elm.text = PLOT_Count
    if elm.name == "POP_Count":
        if POP_Count is None: elm.text = " "
        else: elm.text = POP_Count
    if elm.name == "TWATEEQUNITS_Count":
        if TWATEEQUNITS_Count is None: elm.text = " "
        else: elm.text = TWATEEQUNITS_Count
    if elm.name == "TWATEEQCONT_Count":
        if TWATEEQCONTRACTS_Count is None: elm.text = " "
        else: elm.text = TWATEEQCONTRACTS_Count
    if elm.name == "outseatingpermit_Count":
        if outseatingpermit_Count is None: elm.text = " "
        else: elm.text = outseatingpermit_Count
    if elm.name == "foodtruckpermit_Count":
        if foodtruckpermit_Count is None: elm.text = " "
        else: elm.text = foodtruckpermit_Count
    if elm.name == "DEDLICENSES_Ind":
        if DEDLICENSES_Ind is None: elm.text = " "
        else: elm.text = DEDLICENSES_Ind
    if elm.name == "DEDLICENSES_Comm":
        if DEDLICENSES_Comm is None: elm.text = " "
        else: elm.text = DEDLICENSES_Comm
    if elm.name == "TWATEEQUNITS_Res":
        if TWATEEQUNITS_Res is None: elm.text = " "
        else: elm.text = TWATEEQUNITS_Res
    if elm.name == "TWATEEQUNITS_Comm":
        if TWATEEQUNITS_Comm is None: elm.text = " "
        else: elm.text = TWATEEQUNITS_Comm
    if elm.name == "TWATEEQUNITS_Ind":
        if TWATEEQUNITS_Ind is None: elm.text = " "
        else: elm.text = TWATEEQUNITS_Ind
    if elm.name == "AsBulidMisc":
        if AsBulidMisc is None: elm.text = " "
        else: elm.text = AsBulidMisc
    if elm.name == "AsBulidMain":
        if AsBulidMain is None: elm.text = " "
        else: elm.text = AsBulidMain
    if elm.name == "AsPlannedMisc":
        if AsPlannedMisc is None: elm.text = " "
        else: elm.text = AsPlannedMisc
    if elm.name == "AsPlannedMain":
        if AsPlannedMain is None: elm.text = " "
        else: elm.text = AsPlannedMain
    if elm.name == "NBMBulding":
        if NBMBulding is None: elm.text = " "
        else: elm.text = NBMBulding
    if elm.name == "BDMSBulding":
        if BDMSBulding is None: elm.text = " "
        else: elm.text = BDMSBulding
    if elm.name == "twacont_comm":
        if TWATEEQContract_Res is None: elm.text = " "
        else: elm.text = TWATEEQContract_Res
    if elm.name == "twacont_ind":
        if TWATEEQContract_Comm is None: elm.text = " "
        else: elm.text = TWATEEQContract_Comm
    if elm.name == "twatcont_Res":
        if TWATEEQContract_Ind is None: elm.text = " "
        else: elm.text = TWATEEQContract_Ind
#-----------------------------------------------------------------------------------------------------------   
for filename in os.listdir(zipFolder):
    if filename.endswith('png') or filename.endswith('jpeg'):
        baseimagename = os.path.splitext(filename)[0]
        if baseimagename == "MAP":
            mapImagename = filename
            mapImagenamepath = zipFolder + "\\" + mapImagename

            mapImageraster = arcpy.sa.Raster(mapImagenamepath)

            mapImageheight = mapImageraster.height
            #arcpy.AddMessage(mapImageheight)

            mapImagewidth = mapImageraster.width
            #arcpy.AddMessage(mapImagewidth)

            mapimagewidth = (mapImagewidth * 2.54 ) / 96.0
            #arcpy.AddMessage("mapimagewidth: " + str(mapimagewidth))
            mapimageheight = (mapImageheight * 2.54 ) / 96.0
            #arcpy.AddMessage("mapimageheight:" + str(mapimageheight))

            MapImage_PIC = arcpy.mapping.ListLayoutElements(mxd, "PICTURE_ELEMENT", "MapImage")[0]
            #arcpy.AddMessage(MapImage_PIC.elementWidth)
            #arcpy.AddMessage(MapImage_PIC.elementHeight)
            MapImage_PIC.sourceImage = mapImagenamepath

            #MapImage_PIC.elementPositionX = 14.5
            #MapImage_PIC.elementPositionX = MapImage_PIC.elementPositionX
            #MapImage_PIC_y = pDataFrame.elementPositionY
            #arcpy.AddMessage("MapImage_PIC_y: " + str(MapImage_PIC_y))

            #MapImage_PIC_Height = pDataFrame.elementHeight
            #arcpy.AddMessage("MapImage_PIC_Height: " + str(MapImage_PIC_Height))

            #MapImage_PIC.elementPositionY = (MapImage_PIC_y + MapImage_PIC_Height) - mapimageheight
            #arcpy.AddMessage("MapImage_PIC_Y: " + str((MapImage_PIC_y + MapImage_PIC_Height) - mapimageheight))

            #MapImage_PIC.elementWidth = mapimagewidth
            #MapImage_PIC.elementHeight = mapimageheight
        if baseimagename == "PlotConstructionStatus":
            PlotConstructionStatusImagename = filename
            PlotConstructionStatusImagenamepath = zipFolder + "\\" + PlotConstructionStatusImagename
            PlotConstructionStatusImageraster = arcpy.sa.Raster(PlotConstructionStatusImagenamepath)
            PlotConstructionStatusImageheight = PlotConstructionStatusImageraster.height
            PlotConstructionStatusImagewidth = PlotConstructionStatusImageraster.width

            PlotConstructionStatusImagewidthcms = (PlotConstructionStatusImagewidth * 2.54 ) / 96.0
            PlotConstructionStatusImageheightcms = (PlotConstructionStatusImageheight * 2.54 ) / 96.0

            PlotConstructionStatusImage_PIC = arcpy.mapping.ListLayoutElements(mxd, "PICTURE_ELEMENT", "PlotConstructionStatusImage")[0]
            PlotConstructionStatusImage_PIC.sourceImage = PlotConstructionStatusImagenamepath

##            PlotConstructionStatusImage_PIC.elementPositionX = PlotConstructionStatusImage_PIC.elementPositionX
##            PlotConstructionStatusImage_PIC_y = pDataFrame.elementPositionY
##            PlotConstructionStatusImage_PIC_Height = pDataFrame.elementHeight
##            PlotConstructionStatusImage_PIC.elementPositionY = (PlotConstructionStatusImage_PIC_y + PlotConstructionStatusImage_PIC_Height) - PlotConstructionStatusImageheightcms
##
##            PlotConstructionStatusImage_PIC.elementWidth = PlotConstructionStatusImagewidthcms
##            PlotConstructionStatusImage_PIC.elementHeight = PlotConstructionStatusImageheightcms
        if baseimagename == "UsageTypes":
            UsageTypesImagename = filename
            UsageTypesImagenamepath = zipFolder + "\\" + UsageTypesImagename
            UsageTypesImageraster = arcpy.sa.Raster(UsageTypesImagenamepath)
            UsageTypesImageheight = UsageTypesImageraster.height
            UsageTypesImagewidth = UsageTypesImageraster.width

            UsageTypesImagewidthcms = (UsageTypesImagewidth * 2.54 ) / 96.0
            UsageTypesImageheightcms = (UsageTypesImageheight * 2.54 ) / 96.0

            #UsageTypesImage_PIC = arcpy.mapping.ListLayoutElements(mxd, "PICTURE_ELEMENT", "UsageTypes")[0]
            UsageTypesImage_PIC.sourceImage = UsageTypesImagenamepath
##            UsageTypesImage_PIC.elementPositionX = 14.4992
##            UsageTypesImage_PIC_y = pDataFrame.elementPositionY
##            UsageTypesImage_PIC_Height = pDataFrame.elementHeight
##            UsageTypesImage_PIC.elementPositionY = (UsageTypesImage_PIC_y + UsageTypesImage_PIC_Height) - UsageTypesImageheightcms

            UsageTypesImage_PIC.elementWidth = UsageTypesImagewidthcms
            UsageTypesImage_PIC.elementHeight = UsageTypesImageheightcms
#-----------------------------------------------------------------------------------------------------------
ArchivePath = r"\\ap007fsn01\GISApplications\sddpdmsdata\gpExportLayerToExcel"
archivename =datetime.datetime.now().strftime("D%Y%m%d%H%M%S")
mxd.saveACopy(ArchivePath +"\\"+ archivename +".mxd")
arcpy.mapping.ExportToPDF(mxd, ArchivePath +"\\"+ archivename + ".PDF",resolution=600,embed_fonts=True, image_quality="NORMAL")
arcpy.SetParameterAsText(1, archivename + ".PDF")

