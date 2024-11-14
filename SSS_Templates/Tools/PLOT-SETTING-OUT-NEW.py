# -*- coding: cp1256 -*-
import arcpy
import datetime
import uuid
import os,math,numpy
#===========================================================================================================
arcpy.env.overwriteOutput = True
#===========================================================================================================
arcpy.AddMessage("Collecting input parameters")
#-----------------------------------------------------------------------------------------------------------
REQNUM = arcpy.GetParameterAsText(0)
TempalteName = arcpy.GetParameterAsText(1)
#templaterows = arcpy.GetParameterAsText(4)
#Scale = arcpy.GetParameterAsText(5)
#===========================================================================================================
arcpy.AddMessage("Loading MapDocument")
#-----------------------------------------------------------------------------------------------------------

DocumentPath = r"\\ap007fsc05fs\GISApplications\agsadmgis\arcgisdata\Smartsurvey\\" + TempalteName +  ".mxd"

mxd = arcpy.mapping.MapDocument(DocumentPath)
pDataFrame = arcpy.mapping.ListDataFrames(mxd, "Layers")[0]
pSpatialReference = pDataFrame.spatialReference
#===========================================================================================================
arcpy.AddMessage("Getting Sitemap Info")
#-----------------------------------------------------------------------------------------------------------
V_APPLICATION_INFO = arcpy.mapping.ListTableViews(mxd,"V_APPLICATION_INFO_TEMP")[0]
pCursor = arcpy.da.SearchCursor(V_APPLICATION_INFO, "*" ,"REQNUM = '" + REQNUM + "'")
for row in pCursor:
        RefNo = row[0]
        Zone = row[6]
        SECTORNUMBER = row[8]
        PLOTNUMBER = row[10]
        PlotArea = row[11]
        Category = row[12]
        Surveyorname =  row[15]
        LayersPath = row[24]
        Remarks = row[22]
        OwnerName = row[27]
        ARCHIVE_PATH =  row[34]
        MAP_XMIN = row[35]
        MAP_YMIN = row[36]
        MAP_XMAX = row [37]
        MAP_YMAX = row[38]
        CONSULTANTCOMPANYNAME = row[26]
        CONTRACTORCOMPANYNAME = row[28]
        RECIPIENTCONSULTANTENGNAM = row[29]
        RECIPIENTCONTRACTORENGNAM = row[30]
        RECIPIENTCONSULTANTENGPHONENUM = row[31]
        RECIPIENTCONTRACTORENGPHONENUM = row[32]
        QCUSER = row[53]
        QCTITLE = row[54]
        SURVEY_FINISH_DATE = row[62]
        QC_FINISH_DATE = row[55]
        DEGREE = row[59]
#===========================================================================================================
arcpy.AddMessage("Adding Map Layers")
#-----------------------------------------------------------------------------------------------------------
#pDataFrame.rotation = MAP_ANGLE

S_TSUR_POINT_Layer = arcpy.mapping.Layer(LayersPath + "S_TSUR_POINT.lyr")
S_TSUR_POINT_Layer.definitionQuery = "REQNUM = '" + REQNUM + "'"
arcpy.mapping.AddLayer(pDataFrame, S_TSUR_POINT_Layer, "BOTTOM")

QiblaLayer_Layer = arcpy.mapping.Layer(LayersPath + "QiblaLayer.lyr")
QiblaLayer_Layer.definitionQuery = "REQNUM = '" + REQNUM + "'"
arcpy.mapping.AddLayer(pDataFrame, QiblaLayer_Layer, "BOTTOM")

PlotDimensions_Layer = arcpy.mapping.Layer(LayersPath + "PlotDimensions.lyr")
PlotDimensions_Layer.definitionQuery = "REQNUM = '" + REQNUM + "'"
arcpy.mapping.AddLayer(pDataFrame, PlotDimensions_Layer, "BOTTOM")

#===========================================================================================================
RoadCentreLine_Layer = arcpy.mapping.Layer(LayersPath + "RoadCentreLine.lyr")
arcpy.mapping.AddLayer(pDataFrame, RoadCentreLine_Layer, "BOTTOM")

RoadEdge_Layer = arcpy.mapping.Layer(LayersPath + "RoadEdge.lyr")
arcpy.mapping.AddLayer(pDataFrame, RoadEdge_Layer, "BOTTOM")

RoadBarrier_Layer = arcpy.mapping.Layer(LayersPath + "RoadBarrier.lyr")
arcpy.mapping.AddLayer(pDataFrame, RoadBarrier_Layer, "BOTTOM")
#===========================================================================================================
S_TSUR_PLOT_Layer = arcpy.mapping.Layer(LayersPath + "S_TSUR_PLOT.lyr")
S_TSUR_PLOT_Layer.definitionQuery = "REQNUM = '" + REQNUM + "'"
arcpy.mapping.AddLayer(pDataFrame, S_TSUR_PLOT_Layer, "BOTTOM")

Building_Layer = arcpy.mapping.Layer(LayersPath +  "Building.lyr")
arcpy.mapping.AddLayer(pDataFrame, Building_Layer, "BOTTOM")

SurveyedPlotZ40_Layer = arcpy.mapping.Layer(LayersPath + "SurveyedPlotZ40.lyr")
arcpy.mapping.AddLayer(pDataFrame, SurveyedPlotZ40_Layer, "BOTTOM")

#===========================================================================================================
Plot_OutLine_Layer = arcpy.mapping.Layer(LayersPath + "Plot_OutLine.lyr")
Plot_OutLine_Layer.definitionQuery = "not (SECTORTPSSNUMBER = '" + SECTORNUMBER + "' AND PLOTNUMBER = '" + PLOTNUMBER + "')"
arcpy.mapping.AddLayer(pDataFrame, Plot_OutLine_Layer, "BOTTOM")

SECTOR_Layer = arcpy.mapping.Layer(LayersPath + "Sector.lyr")
arcpy.mapping.AddLayer(pDataFrame, SECTOR_Layer, "BOTTOM")

ZONE_Layer = arcpy.mapping.Layer(LayersPath + "Zone.lyr")
arcpy.mapping.AddLayer(pDataFrame, ZONE_Layer, "BOTTOM")

#===========================================================================================================
arcpy.AddMessage("Setting map extent and scale")
#-----------------------------------------------------------------------------------------------------------
pExtent= arcpy.Extent(MAP_XMIN,MAP_YMIN,MAP_XMAX,MAP_YMAX)
pExtent.spatialReference = pSpatialReference
pDataFrame.extent = pExtent
#if Scale <> "":
        #pDataFrame.scale = Scale
##if DEGREE <> 0:
##        pDataFrame.rotation = DEGREE
#===========================================================================================================
arcpy.AddMessage("Getting Sitemap Coordinates")
#-----------------------------------------------------------------------------------------------------------
pCursor = arcpy.SearchCursor(r"\\ap007fsc05fs\GISApplications\agsadmgis\arcgisdata\Smartsurvey\Layers\S_TSUR_POINT.lyr","REQNUM = '" + REQNUM + "'")
#-----------------------------------------------------------------------------------------------------------
pCoordinates = []
for row in pCursor:
    obsno = row.getValue("OBSNO")
    shp =row.POINT
    easting = shp.firstPoint.X
    northing = shp.firstPoint.Y
    pCoordinates.insert(len(pCoordinates),[int(obsno),easting,northing])
#-----------------------------------------------------------------------------------------------------------
pCoordinates = sorted(pCoordinates)
#===========================================================================================================	
arcpy.AddMessage("Setting Sitemap Info")
#-----------------------------------------------------------------------------------------------------------
##pElems = arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT", "P*")
##xElems = arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT", "X*")
##yElems = arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT", "Y*")
##
##count  = len(pCoordinates)
##arcpy.AddMessage(count)
##
##pcount = 0
##for elm in pElems:
##        elmname = elm.name
##        elmname = elmname.split()
##        arcpy.AddMessage(elmname)
##        pcount += 1
##        if count < pcount:
##                elm.delete()
##xcount = 0
##for elm in xElems:
##        xcount += 1
##        if xcount > count:
##                elm.delete()
##ycount = 0
##for elm in yElems:
##        ycount += 1
##        if ycount > count:
##                elm.delete()

chktElems = arcpy.mapping.ListLayoutElements(mxd, "GRAPHIC_ELEMENT", "CHK_*")
for elm in chktElems:
        if Category == "farm":
                if elm.name <> "CHK_1":
                        elm.delete()
        elif Category == "agricultural":
                if elm.name <> "CHK_1":
                        elm.delete()
        elif Category == "pf":
                if elm.name <> "CHK_2":
                        elm.delete()
        elif Category == "cultural":
                if elm.name <> "CHK_2":
                        elm.delete()
        elif Category == "health":
                if elm.name <> "CHK_2":
                        elm.delete()
        elif Category == "archaeological":
                if elm.name <> "CHK_2":
                        elm.delete()
        elif Category == "recreational":
                if elm.name <> "CHK_2":
                        elm.delete()
        elif Category == "communication":
                if elm.name <> "CHK_2":
                        elm.delete()
        elif Category == "utility":
                if elm.name <> "CHK_2":
                        elm.delete()
        elif Category == "transportation":
                if elm.name <> "CHK_2":
                        elm.delete()
        elif Category == "religious":
                if elm.name <> "CHK_2":
                        elm.delete()
        elif Category == "diplomatic":
                if elm.name <> "CHK_2":
                        elm.delete()
        elif Category == "educational":
                if elm.name <> "CHK_2":
                        elm.delete()
        elif Category == "governmental":
                if elm.name <> "CHK_2":
                        elm.delete()
        elif Category == "temporary":
                if elm.name <> "CHK_3":
                        elm.delete()
        elif Category == "undefined":
                if elm.name <> "CHK_3":
                        elm.delete()
        elif Category == "industrial":
                if elm.name <> "CHK_4":
                        elm.delete()
        elif Category == "commercial":
                if elm.name <> "CHK_5":
                        elm.delete()
        elif Category == "commercialResidential":
                if elm.name <> "CHK_5":
                        elm.delete()
        elif Category == "residential":
                if elm.name <> "CHK_6":
                        elm.delete()
        elif Category == "private":
                if elm.name <> "CHK_6":
                        elm.delete()
for elm in arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT"):
        if elm.name == "OwnerName":
                if OwnerName is None: elm.text = " "
                else: elm.text = OwnerName	   
        if elm.name == "PlotNo":
                if PLOTNUMBER is None: elm.text = " "
                else:elm.text = PLOTNUMBER                           
        if elm.name == "SectorNo":
                if SECTORNUMBER is None: elm.text = " "
                else: elm.text = SECTORNUMBER	   
        if elm.name == "Zone":
                if Zone is None: elm.text = " "
                else: elm.text = Zone	   
        if elm.name == "PlotArea":
                if PlotArea is None: elm.text = " "
                else: elm.text = PlotArea
        if elm.name == "RefNo":
                if RefNo is None: elm.text = " "
                else: elm.text = RefNo
        if elm.name == "Surveyor":
                if Surveyorname is None: elm.text = " "
                else: elm.text = Surveyorname
        if elm.name == "CONSULTANTCOMPANYNAME":
                if CONSULTANTCOMPANYNAME is None: elm.text = " "
                else: elm.text = CONSULTANTCOMPANYNAME
        if elm.name == "CONTRACTORCOMPANYNAME":
                if CONTRACTORCOMPANYNAME is None: elm.text = " "
                else: elm.text = CONTRACTORCOMPANYNAME
        if elm.name == "RECIPIENTCONSULTANTENGNAM":
                if RECIPIENTCONSULTANTENGNAM is None: elm.text = " "
                else: elm.text = RECIPIENTCONSULTANTENGNAM
        if elm.name == "RECIPIENTCONTRACTORENGNAM":
                if RECIPIENTCONTRACTORENGNAM is None: elm.text = " "
                else: elm.text = RECIPIENTCONTRACTORENGNAM
        if elm.name == "RECIPIENTCONSULTANTENGPHONENUM":
                if RECIPIENTCONSULTANTENGPHONENUM is None: elm.text = " "
                else: elm.text = RECIPIENTCONSULTANTENGPHONENUM
        if elm.name == "RECIPIENTCONTRACTORENGPHONENUM":
                if RECIPIENTCONTRACTORENGPHONENUM is None: elm.text = " "
                else: elm.text = RECIPIENTCONTRACTORENGPHONENUM
        if elm.name == "Remarks":
                if Remarks is None: elm.text = " "
                else: elm.text = Remarks
##        if elm.name == "QCTITLE":
##                if QCTITLE is None: elm.text = " "
##                else: elm.text = QCTITLE
##        if elm.name == "QCUSER":
##                if QCUSER is None: elm.text = " "
##                else: elm.text = QCUSER
        if elm.name == "SURVEY_FINISH_DATE":
                if SURVEY_FINISH_DATE is None: elm.text = " "
                else: elm.text = SURVEY_FINISH_DATE
##        if elm.name == "QC_FINISH_DATE":
##                if QC_FINISH_DATE is None: elm.text = " "
##                else: elm.text = QC_FINISH_DATE

        #+++++++++++++++++++++++++++++++++++++++++++
 
                
                
#===========================================================================================================
arcpy.AddMessage("Setting QR Code Image")
#-----------------------------------------------------------------------------------------------------------
if not os.path.exists(os.path.dirname(ARCHIVE_PATH)):
        os.makedirs(os.path.dirname(ARCHIVE_PATH))
#===========================================================================================================
arcpy.AddMessage("Exporting SITEMAP to PDF")
#-----------------------------------------------------------------------------------------------------------
mxd.saveACopy(ARCHIVE_PATH + REQNUM + "_SR" + ".mxd")
arcpy.mapping.ExportToPDF(mxd, ARCHIVE_PATH + REQNUM + "_SR" + ".PDF",resolution=600,embed_fonts=True, image_quality="NORMAL")
arcpy.SetParameterAsText(2, ARCHIVE_PATH + REQNUM + "_SR" + ".PDF")
arcpy.SetParameterAsText(3, "Success")



