import arcpy
import datetime
import uuid
import os
#===========================================================================================================
arcpy.env.overwriteOutput = True
#===========================================================================================================
arcpy.AddMessage("Collecting input parameters")
#-----------------------------------------------------------------------------------------------------------
REQNUM = arcpy.GetParameterAsText(0)
TempalteName = arcpy.GetParameterAsText(1)
templaterows = arcpy.GetParameterAsText(4)
Scale = arcpy.GetParameterAsText(5)
Excel = arcpy.GetParameterAsText(6)
Legend = arcpy.GetParameterAsText(7)
#===========================================================================================================
arcpy.AddMessage("Loading MapDocument")
#-----------------------------------------------------------------------------------------------------------
DocumentPath = r"//AD007APP06/gisapplications/gisapps/sddsps/Templates/" + TempalteName +  ".mxd"
arcpy.AddMessage(DocumentPath)
mxd = arcpy.mapping.MapDocument(DocumentPath)
pDataFrame = arcpy.mapping.ListDataFrames(mxd, "Layers")[0]
pSpatialReference = pDataFrame.spatialReference
#===========================================================================================================
arcpy.AddMessage("Getting Sitemap Info")
#-----------------------------------------------------------------------------------------------------------
V_SM_FULL_INFO = arcpy.mapping.ListTableViews(mxd,"V_SM_FULL_INFO")[0]
pCursor = arcpy.da.SearchCursor(V_SM_FULL_INFO, "*" ,"REQNUM = '" + REQNUM + "'")
intcount = 0
for row in pCursor:
        intcount += 1
        OwnerName = row[2]
        Tenant = row[3]
        TheDate = row[4]
        TRANSACTION_DESCRIPTION_ARA = row[5]
        REQUEST_BY = row[6]
        LETTER_NO = row[7]
        LETTER_DATE = row[8]
        OWNER_LABEL = row[9]
        TENANT_LABEL = row[10]
        BENEFICIARY_LABEL = row[11]
        REMARKS = row[12]
        DRAFTSMAN_NAME = row[15]
        MAP_ANGLE = row[22]
        MAP_XMIN = row[23]
        MAP_YMIN = row[24]
        MAP_XMAX = row[25]
        MAP_YMAX = row[26]
        TMP_RENTSTARTDATE = row[27]
        TMP_RENTENDDATE = row[28]
        TMP_RENTENDDETAILS = row[29]
        MSQ_APPROVEDBY = row[30]
        MSQ_LETTERNO = row[31]
        MSQ_LLETTERDATE = row[32]
        DONATORNAME = row[33]
        MOSQUECAPACITY = row[34]
        UPD_LETTER_NO = row[35]
        UPD_LETTER_DATE = row[36]
        SURVEY_REPORT_NO = row[37]
        SURVEY_REPORT_DATE = row[38]
        CopyTo = row[39]
        ARCHIVE_PATH = row[41]
        LayersPath = row[43]
        PLOTAREAM = row[45]
        PLOTAREAF = row[46]
        CUTAREAM = row[47]
        CUTAREAF = row[48]
        EXTAREAM =row[49]
        EXTAREAF = row[50]
        PlotUse = row[51]
        UserName = row[54]
        BENEFICIARY = row[55]
        PAYEE = row[56]
        purpose = row[57]
        Zone = row[58]
        Sector = row[59]
        Plot = row[60]
        OPERATION_TYPE = row[61]
        IS_SRVC_PAYABLE = row[63]
        SMFOREN = row[70]
        REMARKS_FONTSIZE = row[72]
        LANDTYPE_FONTSIZE = row[73]
        PLOTBNDRY_DASHED = row[74]
        APPPLT_TBLFONT = row[75]
        ACTION_FONTSIZE = row[76]
        PLOT_LABEL = row[77]
        PLOTEXTN_FONTSIZE =row[80]
        ORDERBY = row[81]
        DIMENSIONS_FONTSIZE = row[82]
        SURRONDED_PLOTS_FONTSIZE = row[83]
        ZONE_FONTSIZE = row[84]
        SECTOR_FONTSIZE = row[85]
        WTYPEID = row[88]
        DMT_PRINT_USAGE= row[89]

if intcount <> 0:
        #===========================================================================================================
        arcpy.AddMessage("Adding Map Layers")
        #-----------------------------------------------------------------------------------------------------------
        pDataFrame.rotation = MAP_ANGLE

        #if Plot == "Bulk":
        V_BULKREQUEST_PLOTINTERNALIDS = arcpy.mapping.ListTableViews(mxd,"V_BULKREQUEST_PLOTINTERNALIDS")[0]
        nCursor = arcpy.da.SearchCursor(V_BULKREQUEST_PLOTINTERNALIDS,["PLOTINTERNALID"],"REQNUM = '" + REQNUM + "'")
        BulkPlotids = []
        for row in nCursor:
                BulkPlotids.append(row[0])
        BulkPlotids = ','.join(["'{}'".format(value) for value in BulkPlotids])
        
        V_SM_LAYERS_TOSHOW = arcpy.mapping.ListTableViews(mxd,"V_SM_LAYERS_TOSHOW")[0]
        searchcursor = arcpy.da.SearchCursor(V_SM_LAYERS_TOSHOW, "*" ,"REQNUM = '" + REQNUM + "'")
        lylist = []
        for row in searchcursor:
                lyname = row[2]
                lylist.append(lyname)

        satelliteexist = None
        lyname = "SATELLITE2018"
        if lyname in lylist:
                satelliteexist = lyname

        lyname = "SATELLITE2017"
        if lyname in lylist:
                satelliteexist = lyname

        lyname = "SATELLITE2019Nov"
        if lyname in lylist:
                satelliteexist = lyname

        lyname = "SATELLITE2021"
        if lyname in lylist:
                satelliteexist = lyname

        lyname = "SATELLITE2019Ortho"
        if lyname in lylist:
                satelliteexist = lyname

        lyname = "NP_Centroid"
        if lyname in lylist:
                NP_Centroid_Layer = arcpy.mapping.Layer(LayersPath + "\\" + lyname + ".lyr")
                NP_Centroid_Layer.definitionQuery = "WORKREQUEST = '" + REQNUM + "'"
                arcpy.mapping.AddLayer(pDataFrame, NP_Centroid_Layer, "BOTTOM")
        lyname = "NP_Points"
        if lyname in lylist:
                NP_Points_Layer = arcpy.mapping.Layer(LayersPath + "\\" + lyname + "_A3.lyr")
                NP_Points_Layer.definitionQuery = "WORKREQUEST = '" + REQNUM + "' AND POINT_NUMBER_VISIBLE = 1"
                arcpy.mapping.AddLayer(pDataFrame, NP_Points_Layer, "BOTTOM")
        lyname = "NP_Dimensions"
        if lyname in lylist:
                NP_Dimensions_Layer = arcpy.mapping.Layer(LayersPath + "\\" + lyname + ".lyr")
                if NP_Dimensions_Layer.supports("LABELCLASSES"):
                                if NP_Dimensions_Layer.showLabels:
                                        for lblClass in NP_Dimensions_Layer.labelClasses:
                                                if lblClass.className == "PLOT":
                                                        lblClass.expression = lblClass.expression = "\"<FNT size = '"+ str(DIMENSIONS_FONTSIZE) +"'>\" & [D_LENGTH] & \"</FNT>\""
                NP_Dimensions_Layer.definitionQuery = "WORKREQUEST = '" + REQNUM + "'AND D_LENGTH_VISIBLE = 1"
                arcpy.mapping.AddLayer(pDataFrame, NP_Dimensions_Layer, "BOTTOM")
        #===========================================================================================================
        if OPERATION_TYPE == "Plot Adjustment":
                PLOT_Adjustment = arcpy.mapping.Layer(LayersPath + "\\" + "PLOT_Adjustment.lyr")
                if Plot == "Bulk":
                        PLOT_Adjustment.definitionQuery = "(PLOTINTERNALID IN (" + BulkPlotids + "))"
                if Plot <> "Bulk":
                        #PLOT_Adjustment.definitionQuery = "SECTORTPSSNUMBER = '" + Sector + "' AND PLOTNUMBER = '" + Plot + "'"
                        PLOT_Adjustment.definitionQuery = "(PLOTINTERNALID IN (" + BulkPlotids + "))"
                arcpy.mapping.AddLayer(pDataFrame, PLOT_Adjustment, "BOTTOM")
        if OPERATION_TYPE == "Plot Shifting":
                PLOT_Adjustment = arcpy.mapping.Layer(LayersPath + "\\" + "PLOT_Adjustment.lyr")
                if Plot == "Bulk":
                        PLOT_Adjustment.definitionQuery = "(PLOTINTERNALID IN (" + BulkPlotids + "))"
                if Plot <> "Bulk":
                        #PLOT_Adjustment.definitionQuery = "SECTORTPSSNUMBER = '" + Sector + "' AND PLOTNUMBER = '" + Plot + "'"
                        PLOT_Adjustment.definitionQuery = "(PLOTINTERNALID IN (" + BulkPlotids + "))"
                arcpy.mapping.AddLayer(pDataFrame, PLOT_Adjustment, "BOTTOM")
        if OPERATION_TYPE == "Plot Extension Adjustment":
                dCursor = arcpy.da.SearchCursor(LayersPath + "\\" + "NewPlots.lyr",["SECTORTPSSNUMBER","PLOTNUMBER"],"REQNUM = '" + REQNUM + "'")
                DesignbulkPlotids = []
                for row in dCursor:
                        DesignbulkPlotids.append([row[0],row[1]])
                designids = []
                stconc = None
                for value in DesignbulkPlotids:
                        stconc = "(SECTORTPSSNUMBER = '" + value[0] + "' AND PLOTNUMBER = '" + value[1] + "')"
                        designids.append(stconc)
                designids = 'or'.join(["{}".format(value) for value in designids])
                PLOTDESIGN_Adjustment = arcpy.mapping.Layer(LayersPath + "\\" + "PLOTDESIGN_Adjustment.lyr")
                if Plot == "Bulk":
                        PLOTDESIGN_Adjustment.definitionQuery = "(" + designids + ") and EXTENSIONTYPE LIKE 'X%'" 
                if Plot <> "Bulk":
                        PLOTDESIGN_Adjustment.definitionQuery = "SECTORTPSSNUMBER = '" + Sector + "' AND PLOTNUMBER = '" + Plot + "'"
                arcpy.mapping.AddLayer(pDataFrame, PLOTDESIGN_Adjustment, "BOTTOM")
        if OPERATION_TYPE == "Gap Adjustment":
                dCursor = arcpy.da.SearchCursor(LayersPath + "\\" + "NewPlots.lyr",["SECTORTPSSNUMBER","PLOTNUMBER"],"REQNUM = '" + REQNUM + "'")
                DesignbulkPlotids = []
                for row in dCursor:
                        DesignbulkPlotids.append([row[0],row[1]])
                designids = []
                stconc = None
                for value in DesignbulkPlotids:
                        stconc = "(SECTORTPSSNUMBER = '" + value[0] + "' AND PLOTNUMBER = '" + value[1] + "')"
                        designids.append(stconc)
                designids = 'or'.join(["{}".format(value) for value in designids])
                PLOTDESIGN_Adjustment = arcpy.mapping.Layer(LayersPath + "\\" + "PLOTDESIGN_Adjustment.lyr")
                if Plot == "Bulk":
                        PLOTDESIGN_Adjustment.definitionQuery = "(" + designids + ") and EXTENSIONTYPE LIKE 'G%'"
                if Plot <> "Bulk":
                        PLOTDESIGN_Adjustment.definitionQuery = "SECTORTPSSNUMBER = '" + Sector + "' AND PLOTNUMBER = '" + Plot + "'"
                arcpy.mapping.AddLayer(pDataFrame, PLOTDESIGN_Adjustment, "BOTTOM")
                
        #===========================================================================================================)

        V_NEWPLOTS_VISIBILITY = arcpy.mapping.ListTableViews(mxd,"V_NEWPLOTS_VISIBILITY")[0]
        sCursor = arcpy.da.SearchCursor(V_NEWPLOTS_VISIBILITY, ["PLOTINTERNALID"] ,"REQNUM = '" + REQNUM + "' AND IS_VISIBLE = 0 AND FEATURE_LAYER = 'FGDABEDITZ40.PLOT'")
        PLOTINTERNALID = []

        for row in sorted(sCursor):
            PLOTINTERNALID.append(row[0])
        PLOTINTERNALID = ','.join(["'{}'".format(value) for value in PLOTINTERNALID])

        sCursor = arcpy.da.SearchCursor(V_NEWPLOTS_VISIBILITY, ["PLOTINTERNALID"] ,"REQNUM = '" + REQNUM + "' AND IS_VISIBLE = 1 AND FEATURE_LAYER = 'FGDABEDITZ40.PLOTDESIGN'")
        PlotDesignInternalID = []

        for row in sorted(sCursor):
            PlotDesignInternalID.append(row[0])

        PlotDesignInternalID = ','.join(["'{}'".format(value) for value in PlotDesignInternalID])

        sCursor = arcpy.da.SearchCursor(V_NEWPLOTS_VISIBILITY, ["PLOTINTERNALID"] ,"REQNUM = '" + REQNUM + "' AND IS_VISIBLE = 1 AND FEATURE_LAYER = 'FGDABEDITZ40.PROPOSED_PLOT'")
        PlotProposedInternalID = []

        for row in sorted(sCursor):
            PlotProposedInternalID.append(row[0])

        PlotProposedInternalID = ','.join(["'{}'".format(value) for value in PlotProposedInternalID])
        #===========================================================================================================
        lyname = "PLOT"
        if lyname in lylist:
                if OPERATION_TYPE == "Plot Adjustment":
                        PLOT_Layer = arcpy.mapping.Layer(LayersPath + "\\" + "PLOT.lyr")
                        if PLOT_Layer.supports("LABELCLASSES"):
                                if PLOT_Layer.showLabels:
                                        for lblClass in PLOT_Layer.labelClasses:
                                                lblClass.expression = lblClass.expression = "\"<FNT size = '"+ str(SURRONDED_PLOTS_FONTSIZE) +"'>\" & [PLOTNUMBER] & \"</FNT>\""
                        if PLOTINTERNALID == "":
                                if Plot == "Bulk":
                                        if BulkPlotids <> "":
                                                PLOT_Layer.definitionQuery = "not (PLOTINTERNALID IN (" + BulkPlotids + "))"
                                if Plot <> "Bulk":
                                        PLOT_Layer.definitionQuery = "not(SECTORTPSSNUMBER = '" + Sector + "' AND PLOTNUMBER = '" + Plot + "')"
                        if PLOTINTERNALID <> "":
                                if Plot == "Bulk":
                                        if BulkPlotids <> "":
                                                PLOT_Layer.definitionQuery = "not ((PLOTINTERNALID IN (" + BulkPlotids + ")) or (PLOTINTERNALID IN (" + PLOTINTERNALID + ")))"
                                        if BulkPlotids == "":
                                                PLOT_Layer.definitionQuery = "not ((PLOTINTERNALID IN (" + PLOTINTERNALID + ")))"
                                if Plot <> "Bulk":
                                         PLOT_Layer.definitionQuery = "not ((SECTORTPSSNUMBER = '" + Sector + "' AND PLOTNUMBER = '" + Plot + "') or (PLOTINTERNALID IN (" + PLOTINTERNALID + ")))"
                        arcpy.mapping.AddLayer(pDataFrame, PLOT_Layer, "BOTTOM")
                elif OPERATION_TYPE == "Plot Shifting":
                        PLOT_Layer = arcpy.mapping.Layer(LayersPath + "\\" + "PLOT.lyr")
                        if PLOT_Layer.supports("LABELCLASSES"):
                                if PLOT_Layer.showLabels:
                                        for lblClass in PLOT_Layer.labelClasses:
                                                lblClass.expression = lblClass.expression = "\"<FNT size = '"+ str(SURRONDED_PLOTS_FONTSIZE) +"'>\" & [PLOTNUMBER] & \"</FNT>\""
                        if PLOTINTERNALID == "":
                                if Plot == "Bulk":
                                        if BulkPlotids <> "":
                                                PLOT_Layer.definitionQuery = "not (PLOTINTERNALID IN (" + BulkPlotids + "))"
                                if Plot <> "Bulk":
                                        PLOT_Layer.definitionQuery = "not(SECTORTPSSNUMBER = '" + Sector + "' AND PLOTNUMBER = '" + Plot + "')"
                        if PLOTINTERNALID <> "":
                                if Plot == "Bulk":
                                        if BulkPlotids <> "":
                                                PLOT_Layer.definitionQuery = "not ((PLOTINTERNALID IN (" + BulkPlotids + ")) or (PLOTINTERNALID IN (" + PLOTINTERNALID + ")))"
                                        if BulkPlotids == "":
                                                PLOT_Layer.definitionQuery = "not ((PLOTINTERNALID IN (" + PLOTINTERNALID + ")))"
                                if Plot <> "Bulk":
                                         PLOT_Layer.definitionQuery = "not ((SECTORTPSSNUMBER = '" + Sector + "' AND PLOTNUMBER = '" + Plot + "') or (PLOTINTERNALID IN (" + PLOTINTERNALID + ")))"
                        arcpy.mapping.AddLayer(pDataFrame, PLOT_Layer, "BOTTOM")
                else:
                        PLOT_Layer = arcpy.mapping.Layer(LayersPath + "\\" + "PLOT.lyr")
                        if PLOT_Layer.supports("LABELCLASSES"):
                                if PLOT_Layer.showLabels:
                                        for lblClass in PLOT_Layer.labelClasses:
                                                lblClass.expression = lblClass.expression = "\"<FNT size = '"+ str(SURRONDED_PLOTS_FONTSIZE) +"'>\" & [PLOTNUMBER] & \"</FNT>\""
                        if PLOTINTERNALID == "":
                                if Plot == "Bulk":
                                        if BulkPlotids <> "":
                                                PLOT_Layer.definitionQuery = "not (PLOTINTERNALID IN (" + BulkPlotids + "))"
                                if Plot <> "Bulk":
                                        PLOT_Layer.definitionQuery = "not(SECTORTPSSNUMBER = '" + Sector + "' AND PLOTNUMBER = '" + Plot + "')"
                        if PLOTINTERNALID <> "":
                                if Plot == "Bulk":
                                        if BulkPlotids <> "":
                                                PLOT_Layer.definitionQuery = "not ((PLOTINTERNALID IN (" + BulkPlotids + ")) or (PLOTINTERNALID IN (" + PLOTINTERNALID + ")))"
                                        if BulkPlotids == "":
                                                PLOT_Layer.definitionQuery = "not ((PLOTINTERNALID IN (" + PLOTINTERNALID + ")))"
                                if Plot <> "Bulk":
                                         PLOT_Layer.definitionQuery = "not ((SECTORTPSSNUMBER = '" + Sector + "' AND PLOTNUMBER = '" + Plot + "') or (PLOTINTERNALID IN (" + PLOTINTERNALID + ")))"
                        arcpy.mapping.AddLayer(pDataFrame, PLOT_Layer, "BOTTOM")
        #===========================================================================================================
        lyname = "NewPlots"
        if lyname in lylist:
                if PLOTBNDRY_DASHED <> "1":
                        NewPlotsType_Layer = arcpy.mapping.Layer(LayersPath + "\\" + "NewPlotsType.lyr")
                        NewPlotsType_Layer.definitionQuery = "REQNUM = '" + REQNUM + "' AND EXTTYPE = 'P'"
                        arcpy.mapping.AddLayer(pDataFrame, NewPlotsType_Layer, "BOTTOM")

        lyname = "PROPOSED_PLOT"
        if lyname in lylist:
                PROPOSED_PLOT_Layer = arcpy.mapping.Layer(LayersPath + "\\" + lyname + ".lyr")
                if PlotProposedInternalID <> "":
                        PROPOSED_PLOT_Layer.definitionQuery = "(PLOTINTERNALID IN (" + PlotProposedInternalID + "))"
                arcpy.mapping.AddLayer(pDataFrame, PROPOSED_PLOT_Layer, "BOTTOM")

        lyname = "PlotDesign"
        if lyname in lylist:
                PLOTDESIGN_Layer = arcpy.mapping.Layer(LayersPath + "\\" + lyname + ".lyr")
                if PlotDesignInternalID <> "":
                        PLOTDESIGN_Layer.definitionQuery = "(PLOTDESIGNID IN (" + PlotDesignInternalID + "))"
                arcpy.mapping.AddLayer(pDataFrame, PLOTDESIGN_Layer, "BOTTOM")
                         
        lyname = "NewPlots"
        if lyname in lylist:
                if PLOTBNDRY_DASHED <> "1":
                        NewPlots_Layer = arcpy.mapping.Layer(LayersPath + "\\" + "NewPlots.lyr")
                        if NewPlots_Layer.supports("LABELCLASSES"):
                                if NewPlots_Layer.showLabels:
                                        for lblClass in NewPlots_Layer.labelClasses:
                                                if lblClass.className == "Default":
                                                        #if lblClass.showClassLabels:
                                                        lblClass.expression = lblClass.expression = "\"<FNT size = '"+ str(PLOT_LABEL) +"'>\" & [PLOTNUMBER] & \"</FNT>\""
                                                if lblClass.className == "Extension":
                                                        #if lblClass.showClassLabels:
                                                        extar = u"Ã“¡ „÷«›"
                                                        lblClass.expression = lblClass.expression = "\"<FNT size = '"+ str(PLOTEXTN_FONTSIZE) +"'>"+ extar +"</FNT>\""
                                                if lblClass.className == "GAP":
                                                        #if lblClass.showClassLabels:
                                                        gapar = u"÷„ ›«’·"
                                                        lblClass.expression = lblClass.expression = "\"<FNT size = '"+ str(PLOTEXTN_FONTSIZE) +"'>"+ gapar + "</FNT>\""
                                                if lblClass.className == "CUT":
                                                        #if lblClass.showClassLabels:
                                                        cutar = u"Ã“¡ „” ﬁÿ⁄"
                                                        lblClass.expression = lblClass.expression = "\"<FNT size = '"+ str(PLOTEXTN_FONTSIZE) +"'>"+ cutar +"</FNT>\""
                        NewPlots_Layer.definitionQuery = "REQNUM = '" + REQNUM + "'"
                        if satelliteexist is not None:
                                NewPlots_Layer.transparency = 40
                        arcpy.mapping.AddLayer(pDataFrame, NewPlots_Layer, "BOTTOM")
                else:
                        NewPlots_Layer = arcpy.mapping.Layer(LayersPath + "\\" + "NewPlots_Dashed.lyr")
                        if NewPlots_Layer.supports("LABELCLASSES"):
                                if NewPlots_Layer.showLabels:
                                        for lblClass in NewPlots_Layer.labelClasses:
                                                if lblClass.className == "Default":
                                                        #if lblClass.showClassLabels:
                                                        lblClass.expression = lblClass.expression = "\"<FNT size = '"+ str(PLOT_LABEL) +"'>\" & [PLOTNUMBER] & \"</FNT>\""
                                                if lblClass.className == "Extension":
                                                        #if lblClass.showClassLabels:
                                                        extar = u"Ã“¡ „÷«›"
                                                        lblClass.expression = lblClass.expression = "\"<FNT size = '"+ str(PLOTEXTN_FONTSIZE) +"'>"+ extar +"</FNT>\""
                                                if lblClass.className == "GAP":
                                                        #if lblClass.showClassLabels:
                                                        gapar = u"÷„ ›«’·"
                                                        lblClass.expression = lblClass.expression = "\"<FNT size = '"+ str(PLOTEXTN_FONTSIZE) +"'>"+ gapar + "</FNT>\""
                                                if lblClass.className == "CUT":
                                                        #if lblClass.showClassLabels:
                                                        cutar = u"Ã“¡ „” ﬁÿ⁄"
                                                        lblClass.expression = lblClass.expression = "\"<FNT size = '"+ str(PLOTEXTN_FONTSIZE) +"'>"+ cutar +"</FNT>\""
                        NewPlots_Layer.definitionQuery = "REQNUM = '" + REQNUM + "'"
                        arcpy.mapping.AddLayer(pDataFrame, NewPlots_Layer, "BOTTOM")

        #===========================================================================================================
        arcpy.AddMessage("Setting map extent and scale")
        #-----------------------------------------------------------------------------------------------------------
        pExtent= arcpy.Extent(MAP_XMIN,MAP_YMIN,MAP_XMAX,MAP_YMAX)
        pExtent.spatialReference = pSpatialReference
        pDataFrame.extent = pExtent
        if Scale <> "":
                pDataFrame.scale = Scale
        #===========================================================================================================
        lyname = "PLANNED_ROAD_EDGE_SML"
        if lyname in lylist:
                PLANNED_ROAD_EDGE_SML_Layer = arcpy.mapping.Layer(LayersPath + "\\" + lyname + ".lyr")
                arcpy.mapping.AddLayer(pDataFrame, PLANNED_ROAD_EDGE_SML_Layer, "BOTTOM")

        lyname = "PLANNED_ROAD_CENTRE_LINE_SML"
        if lyname in lylist:
                PLANNED_ROAD_CENTRE_LINE_SML_Layer = arcpy.mapping.Layer(LayersPath + "\\" + lyname + ".lyr")
                arcpy.mapping.AddLayer(pDataFrame, PLANNED_ROAD_CENTRE_LINE_SML_Layer, "BOTTOM")
                
        lyname = "PLANNED_ROAD_EDGE"
        if lyname in lylist:
                PLANNED_ROAD_EDGE_Layer = arcpy.mapping.Layer(LayersPath + "\\" + lyname + ".lyr")
                arcpy.mapping.AddLayer(pDataFrame, PLANNED_ROAD_EDGE_Layer, "BOTTOM")

        lyname = "PLANNED_ROAD_CENTRE_LINE"
        if lyname in lylist:
                PLANNED_ROAD_CENTRE_LINE_Layer = arcpy.mapping.Layer(LayersPath + "\\" + lyname + ".lyr")
                arcpy.mapping.AddLayer(pDataFrame, PLANNED_ROAD_CENTRE_LINE_Layer, "BOTTOM")

        lyname = "ROAD_EDGE"
        if lyname in lylist:
                ROAD_EDGE_Layer = arcpy.mapping.Layer(LayersPath + "\\" + lyname + ".lyr")
                arcpy.mapping.AddLayer(pDataFrame, ROAD_EDGE_Layer, "BOTTOM")

        lyname = "ROAD_CENTRE_LINE"
        if lyname in lylist:
                ROAD_CENTRE_LINE_Layer = arcpy.mapping.Layer(LayersPath + "\\" + lyname + ".lyr")
                arcpy.mapping.AddLayer(pDataFrame, ROAD_CENTRE_LINE_Layer, "BOTTOM")

        lyname = "ZONE"
        if lyname in lylist:
                ZONE_Layer = arcpy.mapping.Layer(LayersPath + "\\" + lyname + ".lyr")
                arcpy.mapping.AddLayer(pDataFrame, ZONE_Layer, "BOTTOM")

        lyname = "SECTOR"
        if lyname in lylist:
                SECTOR_Layer = arcpy.mapping.Layer(LayersPath + "\\" + lyname + ".lyr")
                arcpy.mapping.AddLayer(pDataFrame, SECTOR_Layer, "BOTTOM")

        lyname = "SERVICE_CORRIDOR"
        if lyname in lylist:
                SERVICE_CORRIDOR_Layer = arcpy.mapping.Layer(LayersPath + "\\" + lyname + ".lyr")
                arcpy.mapping.AddLayer(pDataFrame, SERVICE_CORRIDOR_Layer, "BOTTOM")

        lyname = "SATELLITE2018"
        if lyname in lylist:
                ORTHO_Layer = arcpy.mapping.Layer(LayersPath + "\\" + lyname + ".lyr")
                arcpy.mapping.AddLayer(pDataFrame, ORTHO_Layer, "BOTTOM")

        lyname = "SATELLITE2017"
        if lyname in lylist:
                ORTHO_Layer = arcpy.mapping.Layer(LayersPath + "\\" + lyname + ".lyr")
                arcpy.mapping.AddLayer(pDataFrame, ORTHO_Layer, "BOTTOM")

        lyname = "SATELLITE2019Nov"
        if lyname in lylist:
                ORTHO_Layer = arcpy.mapping.Layer(LayersPath + "\\" + lyname + ".lyr")
                arcpy.mapping.AddLayer(pDataFrame, ORTHO_Layer, "BOTTOM")

        lyname = "SATELLITE2021"
        if lyname in lylist:
                ORTHO_Layer = arcpy.mapping.Layer(LayersPath + "\\" + lyname + ".lyr")
                arcpy.mapping.AddLayer(pDataFrame, ORTHO_Layer, "BOTTOM")

        lyname = "SATELLITE2019Ortho"
        if lyname in lylist:
                ORTHO_Layer = arcpy.mapping.Layer(LayersPath + "\\" + lyname + ".lyr")
                arcpy.mapping.AddLayer(pDataFrame, ORTHO_Layer, "BOTTOM")

        #===========================================================================================================
        #arcpy.AddMessage("Setting map extent and scale")
        #-----------------------------------------------------------------------------------------------------------
        #pExtent= arcpy.Extent(MAP_XMIN,MAP_YMIN,MAP_XMAX,MAP_YMAX)
        #pExtent.spatialReference = pSpatialReference
        #pDataFrame.extent = pExtent
        #pDataFrame.scale = (round(pDataFrame.scale / 25) + 1) * 25
        #if Scale <> "":
                #mxd.activeView=pDataFrame.name
                #pDataFrame.scale = Scale
                #mxd.activeView='PAGE_LAYOUT'
                #arcpy.RefreshActiveView()
##                m = float(Scale)
##                n = 250.0
##                r = m % n
##                if r > n / 2.0:
##                    m = n * math.ceil(m / n)
##                elif r <= n / 2.0:
##                    m = n * math.floor(m / n)
##                pDataFrame.scale = m
                #pDataFrame.scale = (round(pDataFrame.scale / 25) + 1) * 25
        #===========================================================================================================
        #arcpy.AddMessage("Getting plots Info")
        order = None
        arcpy.AddMessage(ORDERBY)
        if ORDERBY == 'PlotNumber':
                order = 'ORDER BY PLOTNUMBER'
                arcpy.AddMessage(order)
        elif ORDERBY == 'UsageType':
                order = 'ORDER BY CATEGORYUSAGE'
        V_BULKREQUEST_PLOTUSAGE = arcpy.mapping.ListTableViews(mxd,"V_BULKREQUEST_PLOTUSAGE")[0]
        pCursor = arcpy.da.SearchCursor(V_BULKREQUEST_PLOTUSAGE, "*", "REQNUM = '" + REQNUM + "'", sql_clause=(None, order))        
        SM_PLOTS = []
        for row in pCursor:
                SM_PLOTS.insert(len(SM_PLOTS),row)
        #===========================================================================================================
        arcpy.AddMessage("Setting Sitemap Info")
        #-----------------------------------------------------------------------------------------------------------
        appvdPlotsTxtElems = arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT", "LANDUSE_PLOT_*")
        areaPlotElems = arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT", "AREA_PLOT_*")
        extPlotElems = arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT", "EXT_PLOT_*")
        gapPlotElems = arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT", "GAP_PLOT_*")
        cutPlotElems = arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT", "CUT_PLOT_*")
        plotTxtElems = arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT", "PLOT_*")

        if len(SM_PLOTS) < 37:
                Excel_PIC = arcpy.mapping.ListLayoutElements(mxd, "PICTURE_ELEMENT", "Excel")[0]
                Excel_PIC.elementWidth = 0
                Excel_PIC.elementHeight = 0

        
        if len(SM_PLOTS) > 37: 
##             for i in range(0, len(appvdPlotsTxtElems)) :                
##                 appvdPlotsTxtElems[i].delete
##                 areaPlotElems[i].delete
##                 plotTxtElems[i].delete
                for elm in appvdPlotsTxtElems + areaPlotElems + extPlotElems + gapPlotElems + cutPlotElems + plotTxtElems:
                        elm.delete()

        else:
                for elm in appvdPlotsTxtElems + areaPlotElems + extPlotElems + gapPlotElems + cutPlotElems + plotTxtElems:
                        elm.fontSize = APPPLT_TBLFONT                                                              
                
        for elm in arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT"):
                if elm.name == "REQNUM":
                        if REQNUM is None: elm.text = " "
                        else: elm.text = REQNUM
                if elm.name == "TheDate":
                        if TheDate is None: elm.text = " "
                        else:elm.text = TheDate
                if elm.name == "Request":
                        if REQUEST_BY is None: elm.text = " "
                        else: elm.text = REQUEST_BY
                if elm.name == "RequestNo":
                        if LETTER_NO is None: elm.text = " "
                        else: elm.text = LETTER_NO
                if elm.name == "RequestDate":
                        if LETTER_DATE is None: elm.text = " "
                        else: elm.text = LETTER_DATE
                if elm.name == "UPDAPPROVALLETTERNO":	
                        if UPD_LETTER_NO is None: elm.text = " "
                        else: elm.text = UPD_LETTER_NO
                if elm.name == "UPDAPPDATE":	
                        if UPD_LETTER_DATE is None: elm.text = " "
                        else: elm.text = UPD_LETTER_DATE
                if elm.name == "Notes":
                        if REMARKS is None: elm.text = " "
                        else:
                                REMARKS = REMARKS.replace("\n","\\n")
                                REMARKS = REMARKS.replace(":(",":-(")
                                if(REMARKS_FONTSIZE is not None and REMARKS_FONTSIZE <> 0):
                                        elm.fontSize = REMARKS_FONTSIZE
                                elm.text = REMARKS
                                #REMARKS = REMARKS.replace("(","")
                                #REMARKS = REMARKS.replace(")","")
                                #elm.fontSize = int(REMARKS_FONTSIZE)
                                #elm.text = REMARKS
                                #lengthofremarks = len(REMARKS)
                                #arcpy.AddMessage(lengthofremarks)
##                                if (lengthofremarks > 265 and lengthofremarks < 356):
##                                        elm.fontSize = 8
##                                        elm.text = REMARKS
##                                elif lengthofremarks > 356:
##                                        elm.fontSize = 6
##                                        elm.text = REMARKS
##                                else:
##                                        elm.text = REMARKS
                #+++++++++++++++++++++++++++++++++++++++++++
                #+++++++++++++++++++++++++++++++++++++++++++
                if elm.name == "Zone":
                        if Zone is None: elm.text = " "
                        else:
##                                if(ZONE_FONTSIZE is not None and ZONE_FONTSIZE <> 0):
##                                        elm.fontSize = ZONE_FONTSIZE
                                if Zone == "Bulk":
                                        elm.text = "Õ”» «·ÃœÊ· «·„—›ﬁ"
                                else:
                                        elm.text = Zone
                if elm.name == "Sector":
                        if Sector is None: elm.text = " "
                        else:
##                                if(SECTOR_FONTSIZE is not None and SECTOR_FONTSIZE <> 0):
##                                        elm.fontSize = SECTOR_FONTSIZE
                                if Sector == "Bulk":
                                        elm.text = "Õ”» «·ÃœÊ· «·„—›ﬁ"
                                else:
                                        elm.text = Sector
                if elm.name == "Plot":
                        if Plot is None: elm.text = " "
                        else: elm.text = Plot
                if elm.name == "Action":
                        if purpose is None: elm.text = " "
                        else:
                                elm.fontSize = ACTION_FONTSIZE
                                elm.text = purpose
                if elm.name == "UserName":
                        if UserName is None: elm.text = " "
                        else: elm.text = UserName
                if elm.name == "TemplateName":
                        if TempalteName is None: elm.text = " "
                        else: elm.text = TempalteName
                if elm.name == "CopyTo":
                        if CopyTo == "0": elm.text = " "
##                        else: elm.text = CopyTo
                if elm.name == "Payee":
                        if IS_SRVC_PAYABLE == "0": elm.text = " "
##                        else: elm.text = CopyTo
                #+++++++++++++++++++++++++++++++++++++++++++
                #+++++++++++++++++++++++++++++++++++++++++++
                if elm.name == "PLOT":
                        if len(SM_PLOTS) > 37: elm.delete()
                if elm.name == "PLOT_1":
                        if len(SM_PLOTS) < 1: elm.delete()
                        else: elm.text = SM_PLOTS[0][2]	   
                if elm.name == "PLOT_2":
                        if len(SM_PLOTS) < 2: elm.delete()
                        else: elm.text = SM_PLOTS[1][2]	   
                if elm.name == "PLOT_3":
                        if len(SM_PLOTS) < 3: elm.delete()
                        else: elm.text = SM_PLOTS[2][2]	   
                if elm.name == "PLOT_4":
                        if len(SM_PLOTS) < 4: elm.delete()
                        else: elm.text = SM_PLOTS[3][2]	   
                if elm.name == "PLOT_5":
                        if len(SM_PLOTS) < 5: elm.delete()
                        else: elm.text = SM_PLOTS[4][2]	   
                if elm.name == "PLOT_6":
                        if len(SM_PLOTS) < 6: elm.delete()
                        else: elm.text = SM_PLOTS[5][2]	   
                if elm.name == "PLOT_7":
                        if len(SM_PLOTS) < 7: elm.delete()
                        else: elm.text = SM_PLOTS[6][2]	   
                if elm.name == "PLOT_8":
                        if len(SM_PLOTS) < 8: elm.delete()
                        else: elm.text = SM_PLOTS[7][2]	   
                if elm.name == "PLOT_9":
                        if len(SM_PLOTS) < 9: elm.delete()
                        else: elm.text = SM_PLOTS[8][2]	   
                if elm.name == "PLOT_10":
                        if len(SM_PLOTS) < 10: elm.delete()
                        else: elm.text = SM_PLOTS[9][2]	   
                if elm.name == "PLOT_11":
                        if len(SM_PLOTS) < 11: elm.delete()
                        else: elm.text = SM_PLOTS[10][2]	   
                if elm.name == "PLOT_12":
                        if len(SM_PLOTS) < 12: elm.delete()
                        else: elm.text = SM_PLOTS[11][2]	   
                if elm.name == "PLOT_13":
                        if len(SM_PLOTS) < 13: elm.delete()
                        else: elm.text = SM_PLOTS[12][2]	   
                if elm.name == "PLOT_14":
                        if len(SM_PLOTS) < 14: elm.delete()
                        else: elm.text = SM_PLOTS[13][2]	   
                if elm.name == "PLOT_15":
                        if len(SM_PLOTS) < 15: elm.delete()
                        else: elm.text = SM_PLOTS[14][2]	   
                if elm.name == "PLOT_16":
                        if len(SM_PLOTS) < 16: elm.delete()
                        else: elm.text = SM_PLOTS[15][2]	   
                if elm.name == "PLOT_17":
                        if len(SM_PLOTS) < 17: elm.delete()
                        else: elm.text = SM_PLOTS[16][2]	   
                if elm.name == "PLOT_18":
                        if len(SM_PLOTS) < 18: elm.delete()
                        else: elm.text = SM_PLOTS[17][2]	   
                if elm.name == "PLOT_19":
                        if len(SM_PLOTS) < 19: elm.delete()
                        else: elm.text = SM_PLOTS[18][2]	   
                if elm.name == "PLOT_20":
                        if len(SM_PLOTS) < 20: elm.delete()
                        else: elm.text = SM_PLOTS[19][2]	   
                if elm.name == "PLOT_21":
                        if len(SM_PLOTS) < 21: elm.delete()
                        else: elm.text = SM_PLOTS[20][2]	   
                if elm.name == "PLOT_22":
                        if len(SM_PLOTS) < 22: elm.delete()
                        else: elm.text = SM_PLOTS[21][2]	   
                if elm.name == "PLOT_23":
                        if len(SM_PLOTS) < 23: elm.delete()
                        else: elm.text = SM_PLOTS[22][2]	   
                if elm.name == "PLOT_24":
                        if len(SM_PLOTS) < 24: elm.delete()
                        else: elm.text = SM_PLOTS[23][2]	   
                if elm.name == "PLOT_25":
                        if len(SM_PLOTS) < 25: elm.delete()
                        else: elm.text = SM_PLOTS[24][2]	   
                if elm.name == "PLOT_26":
                        if len(SM_PLOTS) < 26: elm.delete()
                        else: elm.text = SM_PLOTS[25][2]	   
                if elm.name == "PLOT_27":
                        if len(SM_PLOTS) < 27: elm.delete()
                        else: elm.text = SM_PLOTS[26][2]	   
                if elm.name == "PLOT_28":
                        if len(SM_PLOTS) < 28: elm.delete()
                        else: elm.text = SM_PLOTS[27][2]	   
                if elm.name == "PLOT_29":
                        if len(SM_PLOTS) < 29: elm.delete()
                        else: elm.text = SM_PLOTS[28][2]	   
                if elm.name == "PLOT_30":
                        if len(SM_PLOTS) < 30: elm.delete()
                        else: elm.text = SM_PLOTS[29][2]	   
                if elm.name == "PLOT_31":
                        if len(SM_PLOTS) < 31: elm.delete()
                        else: elm.text = SM_PLOTS[30][2]	   
                if elm.name == "PLOT_32":
                        if len(SM_PLOTS) < 32: elm.delete()
                        else: elm.text = SM_PLOTS[31][2]	   
                if elm.name == "PLOT_33":
                        if len(SM_PLOTS) < 33: elm.delete()
                        else: elm.text = SM_PLOTS[32][2]	   
                if elm.name == "PLOT_34":
                        if len(SM_PLOTS) < 34: elm.delete()
                        else: elm.text = SM_PLOTS[33][2]	   
                if elm.name == "PLOT_35":
                        if len(SM_PLOTS) < 35: elm.delete()
                        else: elm.text = SM_PLOTS[34][2]
                if elm.name == "PLOT_36":
                        if len(SM_PLOTS) < 36: elm.delete()
                        else: elm.text = SM_PLOTS[35][2]
                if elm.name == "PLOT_37":
                        if len(SM_PLOTS) < 37: elm.delete()
                        else: elm.text = SM_PLOTS[36][2]
                #+++++++++++++++++++++++++++++++++++++++++++
                #+++++++++++++++++++++++++++++++++++++++++++
                if elm.name == "AREA_PLOT":
                        if len(SM_PLOTS) > 37: elm.delete()
                if elm.name == "AREA_PLOT_1":
                        if len(SM_PLOTS) < 1: elm.delete()
                        else: elm.text = SM_PLOTS[0][4]	   
                if elm.name == "AREA_PLOT_2":
                        if len(SM_PLOTS) < 2: elm.delete()
                        else: elm.text = SM_PLOTS[1][4]	   
                if elm.name == "AREA_PLOT_3":
                        if len(SM_PLOTS) < 3: elm.delete()
                        else: elm.text = SM_PLOTS[2][4]	   
                if elm.name == "AREA_PLOT_4":
                        if len(SM_PLOTS) < 4: elm.delete()
                        else: elm.text = SM_PLOTS[3][4]	   
                if elm.name == "AREA_PLOT_5":
                        if len(SM_PLOTS) < 5: elm.delete()
                        else: elm.text = SM_PLOTS[4][4]	   
                if elm.name == "AREA_PLOT_6":
                        if len(SM_PLOTS) < 6: elm.delete()
                        else: elm.text = SM_PLOTS[5][4]	   
                if elm.name == "AREA_PLOT_7":
                        if len(SM_PLOTS) < 7: elm.delete()
                        else: elm.text = SM_PLOTS[6][4]	   
                if elm.name == "AREA_PLOT_8":
                        if len(SM_PLOTS) < 8: elm.delete()
                        else: elm.text = SM_PLOTS[7][4]	   
                if elm.name == "AREA_PLOT_9":
                        if len(SM_PLOTS) < 9:elm.delete()
                        else: elm.text = SM_PLOTS[8][4]	   
                if elm.name == "AREA_PLOT_10":
                        if len(SM_PLOTS) < 10: elm.delete()
                        else: elm.text = SM_PLOTS[9][4]	   
                if elm.name == "AREA_PLOT_11":
                        if len(SM_PLOTS) < 11: elm.delete()
                        else: elm.text = SM_PLOTS[10][4]	   
                if elm.name == "AREA_PLOT_12":
                        if len(SM_PLOTS) < 12: elm.delete()
                        else: elm.text = SM_PLOTS[11][4]	   
                if elm.name == "AREA_PLOT_13":
                        if len(SM_PLOTS) < 13: elm.delete()
                        else: elm.text = SM_PLOTS[12][4]	   
                if elm.name == "AREA_PLOT_14":
                        if len(SM_PLOTS) < 14: elm.delete()
                        else: elm.text = SM_PLOTS[13][4]	   
                if elm.name == "AREA_PLOT_15":
                        if len(SM_PLOTS) < 15: elm.delete()
                        else: elm.text = SM_PLOTS[14][4]	   
                if elm.name == "AREA_PLOT_16":
                        if len(SM_PLOTS) < 16: elm.delete()
                        else: elm.text = SM_PLOTS[15][4]	   
                if elm.name == "AREA_PLOT_17":
                        if len(SM_PLOTS) < 17: elm.delete()
                        else: elm.text = SM_PLOTS[16][4]	   
                if elm.name == "AREA_PLOT_18":
                        if len(SM_PLOTS) < 18: elm.delete()
                        else: elm.text = SM_PLOTS[17][4]	   
                if elm.name == "AREA_PLOT_19":
                        if len(SM_PLOTS) < 19: elm.delete()
                        else: elm.text = SM_PLOTS[18][4]	   
                if elm.name == "AREA_PLOT_20":
                        if len(SM_PLOTS) < 20: elm.delete()
                        else: elm.text = SM_PLOTS[19][4]	   
                if elm.name == "AREA_PLOT_21":
                        if len(SM_PLOTS) < 21: elm.delete()
                        else: elm.text = SM_PLOTS[20][4]	   
                if elm.name == "AREA_PLOT_22":
                        if len(SM_PLOTS) < 22: elm.delete()
                        else: elm.text = SM_PLOTS[21][4]	   
                if elm.name == "AREA_PLOT_23":
                        if len(SM_PLOTS) < 23: elm.delete()
                        else: elm.text = SM_PLOTS[22][4]	   
                if elm.name == "AREA_PLOT_24":
                        if len(SM_PLOTS) < 24: elm.delete()
                        else: elm.text = SM_PLOTS[23][4]	   
                if elm.name == "AREA_PLOT_25":
                        if len(SM_PLOTS) < 25: elm.delete()
                        else: elm.text = SM_PLOTS[24][4]	   
                if elm.name == "AREA_PLOT_26":
                        if len(SM_PLOTS) < 26: elm.delete()
                        else: elm.text = SM_PLOTS[25][4]	   
                if elm.name == "AREA_PLOT_27":
                        if len(SM_PLOTS) < 27: elm.delete()
                        else: elm.text = SM_PLOTS[26][4]	   
                if elm.name == "AREA_PLOT_28":
                        if len(SM_PLOTS) < 28: elm.delete()
                        else: elm.text = SM_PLOTS[27][4]	   
                if elm.name == "AREA_PLOT_29":
                        if len(SM_PLOTS) < 29: elm.delete()
                        else: elm.text = SM_PLOTS[28][4]	   
                if elm.name == "AREA_PLOT_30":
                        if len(SM_PLOTS) < 30: elm.delete()
                        else: elm.text = SM_PLOTS[29][4]	   
                if elm.name == "AREA_PLOT_31":
                        if len(SM_PLOTS) < 31: elm.delete()
                        else: elm.text = SM_PLOTS[30][4]	   
                if elm.name == "AREA_PLOT_32":
                        if len(SM_PLOTS) < 32: elm.delete()
                        else: elm.text = SM_PLOTS[31][4]	   
                if elm.name == "AREA_PLOT_33":
                        if len(SM_PLOTS) < 33: elm.delete()
                        else: elm.text = SM_PLOTS[32][4]	   
                if elm.name == "AREA_PLOT_34":
                        if len(SM_PLOTS) < 34: elm.delete()
                        else: elm.text = SM_PLOTS[33][4]	   
                if elm.name == "AREA_PLOT_35":
                        if len(SM_PLOTS) < 35: elm.delete()
                        else: elm.text = SM_PLOTS[34][4]
                if elm.name == "AREA_PLOT_36":
                        if len(SM_PLOTS) < 36: elm.delete()
                        else: elm.text = SM_PLOTS[35][4]
                if elm.name == "AREA_PLOT_37":
                        if len(SM_PLOTS) < 37: elm.delete()
                        else: elm.text = SM_PLOTS[36][4]
                #+++++++++++++++++++++++++++++++++++++++++++
                #+++++++++++++++++++++++++++++++++++++++++++
                if elm.name == "EXT_PLOT":
                        if len(SM_PLOTS) > 37: elm.delete()
                if elm.name == "EXT_PLOT_1":
                        if len(SM_PLOTS) < 1: elm.delete()
                        else: elm.text = SM_PLOTS[0][11]	   
                if elm.name == "EXT_PLOT_2":
                        if len(SM_PLOTS) < 2: elm.delete()
                        else: elm.text = SM_PLOTS[1][11]	   
                if elm.name == "EXT_PLOT_3":
                        if len(SM_PLOTS) < 3: elm.delete()
                        else: elm.text = SM_PLOTS[2][11]	   
                if elm.name == "EXT_PLOT_4":
                        if len(SM_PLOTS) < 4: elm.delete()
                        else: elm.text = SM_PLOTS[3][11]	   
                if elm.name == "EXT_PLOT_5":
                        if len(SM_PLOTS) < 5: elm.delete()
                        else: elm.text = SM_PLOTS[4][11]	   
                if elm.name == "EXT_PLOT_6":
                        if len(SM_PLOTS) < 6: elm.delete()
                        else: elm.text = SM_PLOTS[5][11]	   
                if elm.name == "EXT_PLOT_7":
                        if len(SM_PLOTS) < 7: elm.delete()
                        else: elm.text = SM_PLOTS[6][11]	   
                if elm.name == "EXT_PLOT_8":
                        if len(SM_PLOTS) < 8: elm.delete()
                        else: elm.text = SM_PLOTS[7][11]	   
                if elm.name == "EXT_PLOT_9":
                        if len(SM_PLOTS) < 9:elm.delete()
                        else: elm.text = SM_PLOTS[8][11]	   
                if elm.name == "EXT_PLOT_10":
                        if len(SM_PLOTS) < 10: elm.delete()
                        else: elm.text = SM_PLOTS[9][11]	   
                if elm.name == "EXT_PLOT_11":
                        if len(SM_PLOTS) < 11: elm.delete()
                        else: elm.text = SM_PLOTS[10][11]	   
                if elm.name == "EXT_PLOT_12":
                        if len(SM_PLOTS) < 12: elm.delete()
                        else: elm.text = SM_PLOTS[11][11]	   
                if elm.name == "EXT_PLOT_13":
                        if len(SM_PLOTS) < 13: elm.delete()
                        else: elm.text = SM_PLOTS[12][11]	   
                if elm.name == "EXT_PLOT_14":
                        if len(SM_PLOTS) < 14: elm.delete()
                        else: elm.text = SM_PLOTS[13][11]	   
                if elm.name == "EXT_PLOT_15":
                        if len(SM_PLOTS) < 15: elm.delete()
                        else: elm.text = SM_PLOTS[14][11]	   
                if elm.name == "EXT_PLOT_16":
                        if len(SM_PLOTS) < 16: elm.delete()
                        else: elm.text = SM_PLOTS[15][11]	   
                if elm.name == "EXT_PLOT_17":
                        if len(SM_PLOTS) < 17: elm.delete()
                        else: elm.text = SM_PLOTS[16][11]	   
                if elm.name == "EXT_PLOT_18":
                        if len(SM_PLOTS) < 18: elm.delete()
                        else: elm.text = SM_PLOTS[17][11]	   
                if elm.name == "EXT_PLOT_19":
                        if len(SM_PLOTS) < 19: elm.delete()
                        else: elm.text = SM_PLOTS[18][11]	   
                if elm.name == "EXT_PLOT_20":
                        if len(SM_PLOTS) < 20: elm.delete()
                        else: elm.text = SM_PLOTS[19][11]	   
                if elm.name == "EXT_PLOT_21":
                        if len(SM_PLOTS) < 21: elm.delete()
                        else: elm.text = SM_PLOTS[20][11]	   
                if elm.name == "EXT_PLOT_22":
                        if len(SM_PLOTS) < 22: elm.delete()
                        else: elm.text = SM_PLOTS[21][11]	   
                if elm.name == "EXT_PLOT_23":
                        if len(SM_PLOTS) < 23: elm.delete()
                        else: elm.text = SM_PLOTS[22][11]	   
                if elm.name == "EXT_PLOT_24":
                        if len(SM_PLOTS) < 24: elm.delete()
                        else: elm.text = SM_PLOTS[23][11]	   
                if elm.name == "EXT_PLOT_25":
                        if len(SM_PLOTS) < 25: elm.delete()
                        else: elm.text = SM_PLOTS[24][11]	   
                if elm.name == "EXT_PLOT_26":
                        if len(SM_PLOTS) < 26: elm.delete()
                        else: elm.text = SM_PLOTS[25][11]	   
                if elm.name == "EXT_PLOT_27":
                        if len(SM_PLOTS) < 27: elm.delete()
                        else: elm.text = SM_PLOTS[26][11]	   
                if elm.name == "EXT_PLOT_28":
                        if len(SM_PLOTS) < 28: elm.delete()
                        else: elm.text = SM_PLOTS[27][11]	   
                if elm.name == "EXT_PLOT_29":
                        if len(SM_PLOTS) < 29: elm.delete()
                        else: elm.text = SM_PLOTS[28][11]	   
                if elm.name == "EXT_PLOT_30":
                        if len(SM_PLOTS) < 30: elm.delete()
                        else: elm.text = SM_PLOTS[29][11]	   
                if elm.name == "EXT_PLOT_31":
                        if len(SM_PLOTS) < 31: elm.delete()
                        else: elm.text = SM_PLOTS[30][11]	   
                if elm.name == "EXT_PLOT_32":
                        if len(SM_PLOTS) < 32: elm.delete()
                        else: elm.text = SM_PLOTS[31][11]	   
                if elm.name == "EXT_PLOT_33":
                        if len(SM_PLOTS) < 33: elm.delete()
                        else: elm.text = SM_PLOTS[32][11]	   
                if elm.name == "EXT_PLOT_34":
                        if len(SM_PLOTS) < 34: elm.delete()
                        else: elm.text = SM_PLOTS[33][11]	   
                if elm.name == "EXT_PLOT_35":
                        if len(SM_PLOTS) < 35: elm.delete()
                        else: elm.text = SM_PLOTS[34][11]
                if elm.name == "EXT_PLOT_36":
                        if len(SM_PLOTS) < 36: elm.delete()
                        else: elm.text = SM_PLOTS[35][11]
                if elm.name == "EXT_PLOT_37":
                        if len(SM_PLOTS) < 37: elm.delete()
                        else: elm.text = SM_PLOTS[36][11]
                #+++++++++++++++++++++++++++++++++++++++++++
                #+++++++++++++++++++++++++++++++++++++++++++
                if elm.name == "GAP_PLOT":
                        if len(SM_PLOTS) > 37: elm.delete()
                if elm.name == "GAP_PLOT_1":
                        if len(SM_PLOTS) < 1: elm.delete()
                        else: elm.text = SM_PLOTS[0][10]	   
                if elm.name == "GAP_PLOT_2":
                        if len(SM_PLOTS) < 2: elm.delete()
                        else: elm.text = SM_PLOTS[1][10]	   
                if elm.name == "GAP_PLOT_3":
                        if len(SM_PLOTS) < 3: elm.delete()
                        else: elm.text = SM_PLOTS[2][10]	   
                if elm.name == "GAP_PLOT_4":
                        if len(SM_PLOTS) < 4: elm.delete()
                        else: elm.text = SM_PLOTS[3][10]	   
                if elm.name == "GAP_PLOT_5":
                        if len(SM_PLOTS) < 5: elm.delete()
                        else: elm.text = SM_PLOTS[4][10]	   
                if elm.name == "GAP_PLOT_6":
                        if len(SM_PLOTS) < 6: elm.delete()
                        else: elm.text = SM_PLOTS[5][10]	   
                if elm.name == "GAP_PLOT_7":
                        if len(SM_PLOTS) < 7: elm.delete()
                        else: elm.text = SM_PLOTS[6][10]	   
                if elm.name == "GAP_PLOT_8":
                        if len(SM_PLOTS) < 8: elm.delete()
                        else: elm.text = SM_PLOTS[7][10]	   
                if elm.name == "GAP_PLOT_9":
                        if len(SM_PLOTS) < 9:elm.delete()
                        else: elm.text = SM_PLOTS[8][10]	   
                if elm.name == "GAP_PLOT_10":
                        if len(SM_PLOTS) < 10: elm.delete()
                        else: elm.text = SM_PLOTS[9][10]	   
                if elm.name == "GAP_PLOT_11":
                        if len(SM_PLOTS) < 11: elm.delete()
                        else: elm.text = SM_PLOTS[10][10]	   
                if elm.name == "GAP_PLOT_12":
                        if len(SM_PLOTS) < 12: elm.delete()
                        else: elm.text = SM_PLOTS[11][10]	   
                if elm.name == "GAP_PLOT_13":
                        if len(SM_PLOTS) < 13: elm.delete()
                        else: elm.text = SM_PLOTS[12][10]	   
                if elm.name == "GAP_PLOT_14":
                        if len(SM_PLOTS) < 14: elm.delete()
                        else: elm.text = SM_PLOTS[13][10]	   
                if elm.name == "GAP_PLOT_15":
                        if len(SM_PLOTS) < 15: elm.delete()
                        else: elm.text = SM_PLOTS[14][10]	   
                if elm.name == "GAP_PLOT_16":
                        if len(SM_PLOTS) < 16: elm.delete()
                        else: elm.text = SM_PLOTS[15][10]	   
                if elm.name == "GAP_PLOT_17":
                        if len(SM_PLOTS) < 17: elm.delete()
                        else: elm.text = SM_PLOTS[16][10]	   
                if elm.name == "GAP_PLOT_18":
                        if len(SM_PLOTS) < 18: elm.delete()
                        else: elm.text = SM_PLOTS[17][10]	   
                if elm.name == "GAP_PLOT_19":
                        if len(SM_PLOTS) < 19: elm.delete()
                        else: elm.text = SM_PLOTS[18][10]	   
                if elm.name == "GAP_PLOT_20":
                        if len(SM_PLOTS) < 20: elm.delete()
                        else: elm.text = SM_PLOTS[19][10]	   
                if elm.name == "GAP_PLOT_21":
                        if len(SM_PLOTS) < 21: elm.delete()
                        else: elm.text = SM_PLOTS[20][10]	   
                if elm.name == "GAP_PLOT_22":
                        if len(SM_PLOTS) < 22: elm.delete()
                        else: elm.text = SM_PLOTS[21][10]	   
                if elm.name == "GAP_PLOT_23":
                        if len(SM_PLOTS) < 23: elm.delete()
                        else: elm.text = SM_PLOTS[22][10]	   
                if elm.name == "GAP_PLOT_24":
                        if len(SM_PLOTS) < 24: elm.delete()
                        else: elm.text = SM_PLOTS[23][10]	   
                if elm.name == "GAP_PLOT_25":
                        if len(SM_PLOTS) < 25: elm.delete()
                        else: elm.text = SM_PLOTS[24][10]	   
                if elm.name == "GAP_PLOT_26":
                        if len(SM_PLOTS) < 26: elm.delete()
                        else: elm.text = SM_PLOTS[25][10]	   
                if elm.name == "GAP_PLOT_27":
                        if len(SM_PLOTS) < 27: elm.delete()
                        else: elm.text = SM_PLOTS[26][10]	   
                if elm.name == "GAP_PLOT_28":
                        if len(SM_PLOTS) < 28: elm.delete()
                        else: elm.text = SM_PLOTS[27][10]	   
                if elm.name == "GAP_PLOT_29":
                        if len(SM_PLOTS) < 29: elm.delete()
                        else: elm.text = SM_PLOTS[28][10]	   
                if elm.name == "GAP_PLOT_30":
                        if len(SM_PLOTS) < 30: elm.delete()
                        else: elm.text = SM_PLOTS[29][10]	   
                if elm.name == "GAP_PLOT_31":
                        if len(SM_PLOTS) < 31: elm.delete()
                        else: elm.text = SM_PLOTS[30][10]	   
                if elm.name == "GAP_PLOT_32":
                        if len(SM_PLOTS) < 32: elm.delete()
                        else: elm.text = SM_PLOTS[31][10]	   
                if elm.name == "GAP_PLOT_33":
                        if len(SM_PLOTS) < 33: elm.delete()
                        else: elm.text = SM_PLOTS[32][10]	   
                if elm.name == "GAP_PLOT_34":
                        if len(SM_PLOTS) < 34: elm.delete()
                        else: elm.text = SM_PLOTS[33][10]	   
                if elm.name == "GAP_PLOT_35":
                        if len(SM_PLOTS) < 35: elm.delete()
                        else: elm.text = SM_PLOTS[34][10]
                if elm.name == "GAP_PLOT_36":
                        if len(SM_PLOTS) < 36: elm.delete()
                        else: elm.text = SM_PLOTS[35][10]
                if elm.name == "GAP_PLOT_37":
                        if len(SM_PLOTS) < 37: elm.delete()
                        else: elm.text = SM_PLOTS[36][10]
                #+++++++++++++++++++++++++++++++++++++++++++
                #+++++++++++++++++++++++++++++++++++++++++++
                if elm.name == "CUT_PLOT":
                        if len(SM_PLOTS) > 37: elm.delete()
                if elm.name == "CUT_PLOT_1":
                        if len(SM_PLOTS) < 1: elm.delete()
                        else: elm.text = SM_PLOTS[0][9]	   
                if elm.name == "CUT_PLOT_2":
                        if len(SM_PLOTS) < 2: elm.delete()
                        else: elm.text = SM_PLOTS[1][9]	   
                if elm.name == "CUT_PLOT_3":
                        if len(SM_PLOTS) < 3: elm.delete()
                        else: elm.text = SM_PLOTS[2][9]	   
                if elm.name == "CUT_PLOT_4":
                        if len(SM_PLOTS) < 4: elm.delete()
                        else: elm.text = SM_PLOTS[3][9]	   
                if elm.name == "CUT_PLOT_5":
                        if len(SM_PLOTS) < 5: elm.delete()
                        else: elm.text = SM_PLOTS[4][9]	   
                if elm.name == "CUT_PLOT_6":
                        if len(SM_PLOTS) < 6: elm.delete()
                        else: elm.text = SM_PLOTS[5][9]	   
                if elm.name == "CUT_PLOT_7":
                        if len(SM_PLOTS) < 7: elm.delete()
                        else: elm.text = SM_PLOTS[6][9]	   
                if elm.name == "CUT_PLOT_8":
                        if len(SM_PLOTS) < 8: elm.delete()
                        else: elm.text = SM_PLOTS[7][9]	   
                if elm.name == "CUT_PLOT_9":
                        if len(SM_PLOTS) < 9:elm.delete()
                        else: elm.text = SM_PLOTS[8][9]	   
                if elm.name == "CUT_PLOT_10":
                        if len(SM_PLOTS) < 10: elm.delete()
                        else: elm.text = SM_PLOTS[9][9]	   
                if elm.name == "CUT_PLOT_11":
                        if len(SM_PLOTS) < 11: elm.delete()
                        else: elm.text = SM_PLOTS[10][9]	   
                if elm.name == "CUT_PLOT_12":
                        if len(SM_PLOTS) < 12: elm.delete()
                        else: elm.text = SM_PLOTS[11][9]	   
                if elm.name == "CUT_PLOT_13":
                        if len(SM_PLOTS) < 13: elm.delete()
                        else: elm.text = SM_PLOTS[12][9]	   
                if elm.name == "CUT_PLOT_14":
                        if len(SM_PLOTS) < 14: elm.delete()
                        else: elm.text = SM_PLOTS[13][9]	   
                if elm.name == "CUT_PLOT_15":
                        if len(SM_PLOTS) < 15: elm.delete()
                        else: elm.text = SM_PLOTS[14][9]	   
                if elm.name == "CUT_PLOT_16":
                        if len(SM_PLOTS) < 16: elm.delete()
                        else: elm.text = SM_PLOTS[15][9]	   
                if elm.name == "CUT_PLOT_17":
                        if len(SM_PLOTS) < 17: elm.delete()
                        else: elm.text = SM_PLOTS[16][9]	   
                if elm.name == "CUT_PLOT_18":
                        if len(SM_PLOTS) < 18: elm.delete()
                        else: elm.text = SM_PLOTS[17][9]	   
                if elm.name == "CUT_PLOT_19":
                        if len(SM_PLOTS) < 19: elm.delete()
                        else: elm.text = SM_PLOTS[18][9]	   
                if elm.name == "CUT_PLOT_20":
                        if len(SM_PLOTS) < 20: elm.delete()
                        else: elm.text = SM_PLOTS[19][9]	   
                if elm.name == "CUT_PLOT_21":
                        if len(SM_PLOTS) < 21: elm.delete()
                        else: elm.text = SM_PLOTS[20][9]	   
                if elm.name == "CUT_PLOT_22":
                        if len(SM_PLOTS) < 22: elm.delete()
                        else: elm.text = SM_PLOTS[21][11]	   
                if elm.name == "CUT_PLOT_23":
                        if len(SM_PLOTS) < 23: elm.delete()
                        else: elm.text = SM_PLOTS[22][9]	   
                if elm.name == "CUT_PLOT_24":
                        if len(SM_PLOTS) < 24: elm.delete()
                        else: elm.text = SM_PLOTS[23][9]	   
                if elm.name == "CUT_PLOT_25":
                        if len(SM_PLOTS) < 25: elm.delete()
                        else: elm.text = SM_PLOTS[24][9]	   
                if elm.name == "CUT_PLOT_26":
                        if len(SM_PLOTS) < 26: elm.delete()
                        else: elm.text = SM_PLOTS[25][9]	   
                if elm.name == "CUT_PLOT_27":
                        if len(SM_PLOTS) < 27: elm.delete()
                        else: elm.text = SM_PLOTS[26][9]	   
                if elm.name == "CUT_PLOT_28":
                        if len(SM_PLOTS) < 28: elm.delete()
                        else: elm.text = SM_PLOTS[27][9]	   
                if elm.name == "CUT_PLOT_29":
                        if len(SM_PLOTS) < 29: elm.delete()
                        else: elm.text = SM_PLOTS[28][9]	   
                if elm.name == "CUT_PLOT_30":
                        if len(SM_PLOTS) < 30: elm.delete()
                        else: elm.text = SM_PLOTS[29][9]	   
                if elm.name == "CUT_PLOT_31":
                        if len(SM_PLOTS) < 31: elm.delete()
                        else: elm.text = SM_PLOTS[30][9]	   
                if elm.name == "CUT_PLOT_32":
                        if len(SM_PLOTS) < 32: elm.delete()
                        else: elm.text = SM_PLOTS[31][9]	   
                if elm.name == "CUT_PLOT_33":
                        if len(SM_PLOTS) < 33: elm.delete()
                        else: elm.text = SM_PLOTS[32][9]	   
                if elm.name == "CUT_PLOT_34":
                        if len(SM_PLOTS) < 34: elm.delete()
                        else: elm.text = SM_PLOTS[33][9]	   
                if elm.name == "CUT_PLOT_35":
                        if len(SM_PLOTS) < 35: elm.delete()
                        else: elm.text = SM_PLOTS[34][9]
                if elm.name == "CUT_PLOT_36":
                        if len(SM_PLOTS) < 36: elm.delete()
                        else: elm.text = SM_PLOTS[35][9]
                if elm.name == "CUT_PLOT_37":
                        if len(SM_PLOTS) < 37: elm.delete()
                        else: elm.text = SM_PLOTS[36][9]
                 #+++++++++++++++++++++++++++++++++++++++++++
                #+++++++++++++++++++++++++++++++++++++++++++
                if elm.name == "ALCTYE_CODE_DESC_ARA":
                        if len(SM_PLOTS) > 37: elm.delete()
                if elm.name == "ALCTYE_CODE_DESC_ARA_1":
                        if len(SM_PLOTS) < 1: elm.delete()
                        else: elm.text = SM_PLOTS[0][12]	   
                if elm.name == "ALCTYE_CODE_DESC_ARA_2":
                        if len(SM_PLOTS) < 2: elm.delete()
                        else: elm.text = SM_PLOTS[1][12]	   
                if elm.name == "ALCTYE_CODE_DESC_ARA_3":
                        if len(SM_PLOTS) < 3: elm.delete()
                        else: elm.text = SM_PLOTS[2][12]	   
                if elm.name == "ALCTYE_CODE_DESC_ARA_4":
                        if len(SM_PLOTS) < 4: elm.delete()
                        else: elm.text = SM_PLOTS[3][12]	   
                if elm.name == "ALCTYE_CODE_DESC_ARA_5":
                        if len(SM_PLOTS) < 5: elm.delete()
                        else: elm.text = SM_PLOTS[4][12]	   
                if elm.name == "ALCTYE_CODE_DESC_ARA_6":
                        if len(SM_PLOTS) < 6: elm.delete()
                        else: elm.text = SM_PLOTS[5][12]	   
                if elm.name == "ALCTYE_CODE_DESC_ARA_7":
                        if len(SM_PLOTS) < 7: elm.delete()
                        else: elm.text = SM_PLOTS[6][12]	   
                if elm.name == "ALCTYE_CODE_DESC_ARA_8":
                        if len(SM_PLOTS) < 8: elm.delete()
                        else: elm.text = SM_PLOTS[7][12]	   
                if elm.name == "ALCTYE_CODE_DESC_ARA_9":
                        if len(SM_PLOTS) < 9: elm.delete()
                        else: elm.text = SM_PLOTS[8][12]	   
                if elm.name == "ALCTYE_CODE_DESC_ARA_10":
                        if len(SM_PLOTS) < 10: elm.delete()
                        else: elm.text = SM_PLOTS[9][12]	   
                if elm.name == "ALCTYE_CODE_DESC_ARA_11":
                        if len(SM_PLOTS) < 11: elm.delete()
                        else: elm.text = SM_PLOTS[10][12]	   
                if elm.name == "ALCTYE_CODE_DESC_ARA_12":
                        if len(SM_PLOTS) < 12: elm.delete()
                        else: elm.text = SM_PLOTS[11][12]	   
                if elm.name == "ALCTYE_CODE_DESC_ARA_13":
                        if len(SM_PLOTS) < 13: elm.delete()
                        else: elm.text = SM_PLOTS[12][12]	   
                if elm.name == "ALCTYE_CODE_DESC_ARA_14":
                        if len(SM_PLOTS) < 14: elm.delete()
                        else: elm.text = SM_PLOTS[13][12]	   
                if elm.name == "ALCTYE_CODE_DESC_ARA_15":
                        if len(SM_PLOTS) < 15: elm.delete()
                        else: elm.text = SM_PLOTS[14][12]	   
                if elm.name == "ALCTYE_CODE_DESC_ARA_16":
                        if len(SM_PLOTS) < 16: elm.delete()
                        else: elm.text = SM_PLOTS[15][12]	   
                if elm.name == "ALCTYE_CODE_DESC_ARA_17":
                        if len(SM_PLOTS) < 17: elm.delete()
                        else: elm.text = SM_PLOTS[16][12]	   
                if elm.name == "ALCTYE_CODE_DESC_ARA_18":
                        if len(SM_PLOTS) < 18: elm.delete()
                        else: elm.text = SM_PLOTS[17][12]	   
                if elm.name == "ALCTYE_CODE_DESC_ARA_19":
                        if len(SM_PLOTS) < 19: elm.delete()
                        else: elm.text = SM_PLOTS[18][12]	   
                if elm.name == "ALCTYE_CODE_DESC_ARA_20":
                        if len(SM_PLOTS) < 20: elm.delete()
                        else: elm.text = SM_PLOTS[19][12]	   
                if elm.name == "ALCTYE_CODE_DESC_ARA_21":
                        if len(SM_PLOTS) < 21: elm.delete()
                        else: elm.text = SM_PLOTS[20][12]	   
                if elm.name == "ALCTYE_CODE_DESC_ARA_22":
                        if len(SM_PLOTS) < 22: elm.delete()
                        else: elm.text = SM_PLOTS[21][12]	   
                if elm.name == "ALCTYE_CODE_DESC_ARA_23":
                        if len(SM_PLOTS) < 23: elm.delete()
                        else: elm.text = SM_PLOTS[22][12]	   
                if elm.name == "ALCTYE_CODE_DESC_ARA_24":
                        if len(SM_PLOTS) < 24: elm.delete()
                        else: elm.text = SM_PLOTS[23][12]	   
                if elm.name == "ALCTYE_CODE_DESC_ARA_25":
                        if len(SM_PLOTS) < 25: elm.delete()
                        else: elm.text = SM_PLOTS[24][12]	   
                if elm.name == "ALCTYE_CODE_DESC_ARA_26":
                        if len(SM_PLOTS) < 26: elm.delete()
                        else: elm.text = SM_PLOTS[25][12]	   
                if elm.name == "ALCTYE_CODE_DESC_ARA_27":
                        if len(SM_PLOTS) < 27: elm.delete()
                        else: elm.text = SM_PLOTS[26][12]	   
                if elm.name == "ALCTYE_CODE_DESC_ARA_28":
                        if len(SM_PLOTS) < 28: elm.delete()
                        else: elm.text = SM_PLOTS[27][12]	   
                if elm.name == "ALCTYE_CODE_DESC_ARA_29":
                        if len(SM_PLOTS) < 29: elm.delete()
                        else: elm.text = SM_PLOTS[28][12]	   
                if elm.name == "ALCTYE_CODE_DESC_ARA_30":
                        if len(SM_PLOTS) < 30: elm.delete()
                        else: elm.text = SM_PLOTS[29][12]	   
                if elm.name == "ALCTYE_CODE_DESC_ARA_31":
                        if len(SM_PLOTS) < 31: elm.delete()
                        else: elm.text = SM_PLOTS[30][12]	   
                if elm.name == "ALCTYE_CODE_DESC_ARA_32":
                        if len(SM_PLOTS) < 32: elm.delete()
                        else: elm.text = SM_PLOTS[31][12]	   
                if elm.name == "ALCTYE_CODE_DESC_ARA_33":
                        if len(SM_PLOTS) < 33: elm.delete()
                        else: elm.text = SM_PLOTS[32][12]	   
                if elm.name == "ALCTYE_CODE_DESC_ARA_34":
                        if len(SM_PLOTS) < 34: elm.delete()
                        else: elm.text = SM_PLOTS[33][12]	   
                if elm.name == "ALCTYE_CODE_DESC_ARA_35":
                        if len(SM_PLOTS) < 35: elm.delete()
                        else: elm.text = SM_PLOTS[34][12]
                if elm.name == "ALCTYE_CODE_DESC_ARA_36":
                        if len(SM_PLOTS) < 36: elm.delete()
                        else: elm.text = SM_PLOTS[35][12]
                if elm.name == "ALCTYE_CODE_DESC_ARA_37":
                        if len(SM_PLOTS) < 37: elm.delete()
                        else: elm.text = SM_PLOTS[36][12]
                #+++++++++++++++++++++++++++++++++++++++++++
                #+++++++++++++++++++++++++++++++++++++++++++
                if DMT_PRINT_USAGE == "1":
##                        if(ALLOCATIONTYPE_DM is not None and PURPOSEOFUSE is not None):
##                                elm.text = str(ALLOCATIONTYPE_DM) + "/" + str(PURPOSEOFUSE)
##                        else: elm.text = " "
                        if elm.name == "LANDUSE_PLOT":
                                if len(SM_PLOTS) > 37: elm.delete()
                        if elm.name == "LANDUSE_PLOT_1":
                                if len(SM_PLOTS) < 1: elm.delete()
                                else:
                                        if(SM_PLOTS[0][7] is not None and SM_PLOTS[0][8] is not None):
                                                elm.text = SM_PLOTS[0][7] + "/" + SM_PLOTS[0][8]
                                        else: elm.text = " "
                        if elm.name == "LANDUSE_PLOT_2":
                                if len(SM_PLOTS) < 2: elm.delete()
                                else:
                                        if(SM_PLOTS[1][7] is not None and SM_PLOTS[1][8] is not None):
                                                elm.text = SM_PLOTS[1][7] + "/" + SM_PLOTS[1][8]
                                        else: elm.text = " "
                        if elm.name == "LANDUSE_PLOT_3":
                                if len(SM_PLOTS) < 3: elm.delete()
                                else:
                                        if(SM_PLOTS[2][7] is not None and SM_PLOTS[2][7] is not None):
                                                elm.text = SM_PLOTS[2][7] + "/" + SM_PLOTS[2][8]
                                        else: elm.text = " "
                        if elm.name == "LANDUSE_PLOT_4":
                                if len(SM_PLOTS) < 4: elm.delete()
                                else:
                                        if(SM_PLOTS[3][7] is not None and SM_PLOTS[3][8] is not None):
                                                elm.text = SM_PLOTS[3][7] + "/" + SM_PLOTS[3][8]
                                        else: elm.text = " "
                        if elm.name == "LANDUSE_PLOT_5":
                                if len(SM_PLOTS) < 5: elm.delete()
                                else:
                                        if(SM_PLOTS[4][7] is not None and SM_PLOTS[4][8] is not None):
                                                elm.text = SM_PLOTS[4][7] + "/" + SM_PLOTS[4][8]
                                        else: elm.text = " "
                        if elm.name == "LANDUSE_PLOT_6":
                                if len(SM_PLOTS) < 6: elm.delete()
                                else:
                                        if(SM_PLOTS[5][7] is not None and SM_PLOTS[5][8] is not None):
                                                elm.text = SM_PLOTS[5][7] + "/" + SM_PLOTS[5][8]
                                        else: elm.text = " "
                                       	   
                        if elm.name == "LANDUSE_PLOT_7":
                                if len(SM_PLOTS) < 7: elm.delete()
                                else:
                                        if(SM_PLOTS[6][7] is not None and SM_PLOTS[6][8] is not None):
                                                elm.text = SM_PLOTS[6][7] + "/" + SM_PLOTS[6][8]
                                        else: elm.text = " "
                                        
                        if elm.name == "LANDUSE_PLOT_8":
                                if len(SM_PLOTS) < 8: elm.delete()
                                else:
                                        if(SM_PLOTS[7][7] is not None and SM_PLOTS[7][8] is not None):
                                                elm.text = SM_PLOTS[7][7] + "/" + SM_PLOTS[7][8]
                                        else: elm.text = " "
                                        	   
                        if elm.name == "LANDUSE_PLOT_9":
                                if len(SM_PLOTS) < 9: elm.delete()
                                else:
                                        if(SM_PLOTS[8][7] is not None and SM_PLOTS[8][8] is not None):
                                                elm.text = SM_PLOTS[8][7] + "/" + SM_PLOTS[8][8]
                                        else: elm.text = " "
                                        
                        if elm.name == "LANDUSE_PLOT_10":
                                if len(SM_PLOTS) < 10: elm.delete()
                                else:
                                        if(SM_PLOTS[9][7] is not None and SM_PLOTS[9][8] is not None):
                                                elm.text = SM_PLOTS[9][7] + "/" + SM_PLOTS[9][8]
                                        else: elm.text = " "
                                        	   
                        if elm.name == "LANDUSE_PLOT_11":
                                if len(SM_PLOTS) < 11: elm.delete()
                                else:
                                        if(SM_PLOTS[10][7] is not None and SM_PLOTS[10][8] is not None):
                                                elm.text = SM_PLOTS[10][7] + "/" + SM_PLOTS[10][8]
                                        else: elm.text = " "
                                        	   
                        if elm.name == "LANDUSE_PLOT_12":
                                if len(SM_PLOTS) < 12: elm.delete()
                                else:
                                        if(SM_PLOTS[11][7] is not None and SM_PLOTS[11][8] is not None):
                                               elm.text = SM_PLOTS[11][7] + "/" + SM_PLOTS[11][8]
                                        else: elm.text = " "
                                        	   
                        if elm.name == "LANDUSE_PLOT_13":
                                if len(SM_PLOTS) < 13: elm.delete()
                                else:
                                        if(SM_PLOTS[12][7] is not None and SM_PLOTS[12][8] is not None):
                                                elm.text = SM_PLOTS[12][7] + "/" + SM_PLOTS[12][8]
                                        else: elm.text = " "
                                        	   
                        if elm.name == "LANDUSE_PLOT_14":
                                if len(SM_PLOTS) < 14: elm.delete()
                                else:
                                        if(SM_PLOTS[13][7] is not None and SM_PLOTS[13][8] is not None):
                                               elm.text = SM_PLOTS[13][7] + "/" + SM_PLOTS[13][8]
                                        else: elm.text = " "
                                        	   
                        if elm.name == "LANDUSE_PLOT_15":
                                if len(SM_PLOTS) < 15: elm.delete()
                                else:
                                        if(SM_PLOTS[14][7] is not None and SM_PLOTS[14][8] is not None):
                                                elm.text = SM_PLOTS[14][7] + "/" + SM_PLOTS[14][8]
                                        else: elm.text = " "
                                        	   
                        if elm.name == "LANDUSE_PLOT_16":
                                if len(SM_PLOTS) < 16: elm.delete()
                                else:
                                        if(SM_PLOTS[15][7] is not None and SM_PLOTS[15][8] is not None):
                                                elm.text = SM_PLOTS[15][7] + "/" + SM_PLOTS[15][8]
                                        else: elm.text = " "
                                        	   
                        if elm.name == "LANDUSE_PLOT_17":
                                if len(SM_PLOTS) < 17: elm.delete()
                                else:
                                        if(SM_PLOTS[16][7] is not None and SM_PLOTS[16][8] is not None):
                                               elm.text = SM_PLOTS[16][7] + "/" + SM_PLOTS[16][8]
                                        else: elm.text = " "
                                        	   
                        if elm.name == "LANDUSE_PLOT_18":
                                if len(SM_PLOTS) < 18: elm.delete()
                                else:
                                        if(SM_PLOTS[17][7] is not None and SM_PLOTS[17][8] is not None):
                                               elm.text = SM_PLOTS[17][7] + "/" + SM_PLOTS[17][8]
                                        else: elm.text = " "
                                        	   
                        if elm.name == "LANDUSE_PLOT_19":
                                if len(SM_PLOTS) < 19: elm.delete()
                                else:
                                        if(SM_PLOTS[18][7] is not None and SM_PLOTS[18][8] is not None):
                                                elm.text = SM_PLOTS[18][7] + "/" + SM_PLOTS[18][8]
                                        else: elm.text = " "
                                        	   
                        if elm.name == "LANDUSE_PLOT_20":
                                if len(SM_PLOTS) < 20: elm.delete()
                                else:
                                        if(SM_PLOTS[19][7] is not None and SM_PLOTS[19][8] is not None):
                                               elm.text = SM_PLOTS[19][7] + "/" + SM_PLOTS[19][8]
                                        else: elm.text = " "
                                        	   
                        if elm.name == "LANDUSE_PLOT_21":
                                if len(SM_PLOTS) < 21: elm.delete()
                                else:
                                        if(SM_PLOTS[20][7] is not None and SM_PLOTS[20][8] is not None):
                                                elm.text =SM_PLOTS[20][7] + "/" + SM_PLOTS[20][8]
                                        else: elm.text = " "
                                        	   
                        if elm.name == "LANDUSE_PLOT_22":
                                if len(SM_PLOTS) < 22: elm.delete()
                                else:
                                        if(SM_PLOTS[21][7] is not None and SM_PLOTS[21][8] is not None):
                                                elm.text = SM_PLOTS[21][7] + "/" + SM_PLOTS[21][8]
                                        else: elm.text = " "
                                        	   
                        if elm.name == "LANDUSE_PLOT_23":
                                if len(SM_PLOTS) < 23: elm.delete()
                                else:
                                        if(SM_PLOTS[22][7] is not None and SM_PLOTS[22][8] is not None):
                                                elm.text = SM_PLOTS[22][7] + "/" + SM_PLOTS[22][8]
                                        else: elm.text = " "
                                        	   
                        if elm.name == "LANDUSE_PLOT_24":
                                if len(SM_PLOTS) < 24: elm.delete()
                                else:
                                        if(SM_PLOTS[23][7] is not None and SM_PLOTS[23][8] is not None):
                                                elm.text = SM_PLOTS[23][7] + "/" + SM_PLOTS[23][8]
                                        else: elm.text = " "
                                        	   
                        if elm.name == "LANDUSE_PLOT_25":
                                if len(SM_PLOTS) < 25: elm.delete()
                                else:
                                        if(SM_PLOTS[24][7] is not None and SM_PLOTS[24][8] is not None):
                                                elm.text = SM_PLOTS[24][7] + "/" + SM_PLOTS[24][8]
                                        else: elm.text = " "
                                        	   
                        if elm.name == "LANDUSE_PLOT_26":
                                if len(SM_PLOTS) < 26: elm.delete()
                                else:
                                        if(SM_PLOTS[25][7] is not None and SM_PLOTS[25][8] is not None):
                                                elm.text = SM_PLOTS[25][7] + "/" + SM_PLOTS[25][8]
                                        else: elm.text = " "
                                        	   
                        if elm.name == "LANDUSE_PLOT_27":
                                if len(SM_PLOTS) < 27: elm.delete()
                                else:
                                        if(SM_PLOTS[26][7] is not None and SM_PLOTS[26][8] is not None):
                                                elm.text = SM_PLOTS[26][7] + "/" + SM_PLOTS[26][8]
                                        else: elm.text = " "
                                        	   
                        if elm.name == "LANDUSE_PLOT_28":
                                if len(SM_PLOTS) < 28: elm.delete()
                                else:
                                        if(SM_PLOTS[27][7] is not None and SM_PLOTS[27][8] is not None):
                                                elm.text = SM_PLOTS[27][7] + "/" + SM_PLOTS[27][8]
                                        else: elm.text = " "
                                        	   
                        if elm.name == "LANDUSE_PLOT_29":
                                if len(SM_PLOTS) < 29: elm.delete()
                                else:
                                        if(SM_PLOTS[28][7] is not None and SM_PLOTS[28][8] is not None):
                                               elm.text = SM_PLOTS[28][7] + "/" + SM_PLOTS[28][8]
                                        else: elm.text = " "
                                        	   
                        if elm.name == "LANDUSE_PLOT_30":
                                if len(SM_PLOTS) < 30: elm.delete()
                                else:
                                        if(SM_PLOTS[29][7] is not None and SM_PLOTS[29][8] is not None):
                                                elm.text = SM_PLOTS[29][7] + "/" + SM_PLOTS[29][8]
                                        else: elm.text = " "
                                        	   
                        if elm.name == "LANDUSE_PLOT_31":
                                if len(SM_PLOTS) < 31: elm.delete()
                                else:
                                        if(SM_PLOTS[30][7] is not None and SM_PLOTS[30][8] is not None):
                                                elm.text = SM_PLOTS[30][7] + "/" + SM_PLOTS[30][8]
                                        else: elm.text = " "
                                        	   
                        if elm.name == "LANDUSE_PLOT_32":
                                if len(SM_PLOTS) < 32: elm.delete()
                                else:
                                        if(SM_PLOTS[31][7] is not None and SM_PLOTS[31][8] is not None):
                                                elm.text = SM_PLOTS[31][7] + "/" + SM_PLOTS[31][8]
                                        else: elm.text = " "
                                        	   
                        if elm.name == "LANDUSE_PLOT_33":
                                if len(SM_PLOTS) < 33: elm.delete()
                                else:
                                        if(SM_PLOTS[32][7] is not None and SM_PLOTS[32][8] is not None):
                                                elm.text = SM_PLOTS[32][7] + "/" + SM_PLOTS[32][8]
                                        else: elm.text = " "
                                        	   
                        if elm.name == "LANDUSE_PLOT_34":
                                if len(SM_PLOTS) < 34: elm.delete()
                                else:
                                        if(SM_PLOTS[33][7] is not None and SM_PLOTS[33][8] is not None):
                                                elm.text = SM_PLOTS[33][7] + "/" + SM_PLOTS[33][8]
                                        else: elm.text = " "
                                        	   
                        if elm.name == "LANDUSE_PLOT_35":
                                if len(SM_PLOTS) < 35: elm.delete()
                                else:
                                        if(SM_PLOTS[34][7] is not None and SM_PLOTS[34][8] is not None):
                                                 elm.text = SM_PLOTS[34][7] + "/" + SM_PLOTS[34][8]
                                        else: elm.text = " "
                                       
                        if elm.name == "LANDUSE_PLOT_36":
                                if len(SM_PLOTS) < 36: elm.delete()
                                else:
                                        if(SM_PLOTS[35][7] is not None and SM_PLOTS[35][8] is not None):
                                                elm.text = SM_PLOTS[35][7] + "/" + SM_PLOTS[35][8]
                                        else: elm.text = " "
                                        
                        if elm.name == "LANDUSE_PLOT_37":
                                if len(SM_PLOTS) < 37: elm.delete()
                                else:
                                        if(SM_PLOTS[36][7] is not None and SM_PLOTS[36][8] is not None):
                                                elm.text = SM_PLOTS[36][7] + "/" + SM_PLOTS[36][8]
                                        else: elm.text = " "
                else:
                        if elm.name == "LANDUSE_PLOT":
                                if len(SM_PLOTS) > 37: elm.delete()
                        if elm.name == "LANDUSE_PLOT_1":
                                if len(SM_PLOTS) < 1: elm.delete()
                                else: elm.text = SM_PLOTS[0][13]
                        if elm.name == "LANDUSE_PLOT_2":
                                if len(SM_PLOTS) < 2: elm.delete()
                                else: elm.text = SM_PLOTS[1][13]
                        if elm.name == "LANDUSE_PLOT_3":
                                if len(SM_PLOTS) < 3: elm.delete()
                                else: elm.text = SM_PLOTS[2][13]	   
                        if elm.name == "LANDUSE_PLOT_4":
                                if len(SM_PLOTS) < 4: elm.delete()
                                else: elm.text = SM_PLOTS[3][13]
                        if elm.name == "LANDUSE_PLOT_5":
                                if len(SM_PLOTS) < 5: elm.delete()
                                else: elm.text = SM_PLOTS[4][13]	   
                        if elm.name == "LANDUSE_PLOT_6":
                                if len(SM_PLOTS) < 6: elm.delete()
                                else: elm.text = SM_PLOTS[5][13]	   
                        if elm.name == "LANDUSE_PLOT_7":
                                if len(SM_PLOTS) < 7: elm.delete()
                                else: elm.text = SM_PLOTS[6][13]
                        if elm.name == "LANDUSE_PLOT_8":
                                if len(SM_PLOTS) < 8: elm.delete()
                                else: elm.text = SM_PLOTS[7][13]	   
                        if elm.name == "LANDUSE_PLOT_9":
                                if len(SM_PLOTS) < 9: elm.delete()
                                else: elm.text = SM_PLOTS[8][13]
                        if elm.name == "LANDUSE_PLOT_10":
                                if len(SM_PLOTS) < 10: elm.delete()
                                else: elm.text = SM_PLOTS[9][13]	   
                        if elm.name == "LANDUSE_PLOT_11":
                                if len(SM_PLOTS) < 11: elm.delete()
                                else: elm.text = SM_PLOTS[10][13]	   
                        if elm.name == "LANDUSE_PLOT_12":
                                if len(SM_PLOTS) < 12: elm.delete()
                                else: elm.text = SM_PLOTS[11][13]	   
                        if elm.name == "LANDUSE_PLOT_13":
                                if len(SM_PLOTS) < 13: elm.delete()
                                else: elm.text = SM_PLOTS[12][13]	   
                        if elm.name == "LANDUSE_PLOT_14":
                                if len(SM_PLOTS) < 14: elm.delete()
                                else: elm.text = SM_PLOTS[13][13]	   
                        if elm.name == "LANDUSE_PLOT_15":
                                if len(SM_PLOTS) < 15: elm.delete()
                                else: elm.text = SM_PLOTS[14][13]	   
                        if elm.name == "LANDUSE_PLOT_16":
                                if len(SM_PLOTS) < 16: elm.delete()
                                else: elm.text = SM_PLOTS[15][13]	   
                        if elm.name == "LANDUSE_PLOT_17":
                                if len(SM_PLOTS) < 17: elm.delete()
                                else: elm.text = SM_PLOTS[16][13]	   
                        if elm.name == "LANDUSE_PLOT_18":
                                if len(SM_PLOTS) < 18: elm.delete()
                                else: elm.text = SM_PLOTS[17][13]	   
                        if elm.name == "LANDUSE_PLOT_19":
                                if len(SM_PLOTS) < 19: elm.delete()
                                else: elm.text = SM_PLOTS[18][13]	   
                        if elm.name == "LANDUSE_PLOT_20":
                                if len(SM_PLOTS) < 20: elm.delete()
                                else: elm.text = SM_PLOTS[19][13]	   
                        if elm.name == "LANDUSE_PLOT_21":
                                if len(SM_PLOTS) < 21: elm.delete()
                                else: elm.text = SM_PLOTS[20][13]	   
                        if elm.name == "LANDUSE_PLOT_22":
                                if len(SM_PLOTS) < 22: elm.delete()
                                else: elm.text = SM_PLOTS[21][3]	   
                        if elm.name == "LANDUSE_PLOT_23":
                                if len(SM_PLOTS) < 23: elm.delete()
                                else: elm.text = SM_PLOTS[22][13]	   
                        if elm.name == "LANDUSE_PLOT_24":
                                if len(SM_PLOTS) < 24: elm.delete()
                                else:elm.text = SM_PLOTS[23][13]	   
                        if elm.name == "LANDUSE_PLOT_25":
                                if len(SM_PLOTS) < 25: elm.delete()
                                else: elm.text = SM_PLOTS[24][13]	   
                        if elm.name == "LANDUSE_PLOT_26":
                                if len(SM_PLOTS) < 26: elm.delete()
                                else: elm.text = SM_PLOTS[25][13]	   
                        if elm.name == "LANDUSE_PLOT_27":
                                if len(SM_PLOTS) < 27: elm.delete()
                                else: elm.text = SM_PLOTS[26][13]	   
                        if elm.name == "LANDUSE_PLOT_28":
                                if len(SM_PLOTS) < 28: elm.delete()
                                else: elm.text = SM_PLOTS[27][13]	   
                        if elm.name == "LANDUSE_PLOT_29":
                                if len(SM_PLOTS) < 29: elm.delete()
                                else: elm.text = SM_PLOTS[28][13]	   
                        if elm.name == "LANDUSE_PLOT_30":
                                if len(SM_PLOTS) < 30: elm.delete()
                                else: elm.text = SM_PLOTS[29][13]	   
                        if elm.name == "LANDUSE_PLOT_31":
                                if len(SM_PLOTS) < 31: elm.delete()
                                else: elm.text = SM_PLOTS[30][13]	   
                        if elm.name == "LANDUSE_PLOT_32":
                                if len(SM_PLOTS) < 32: elm.delete()
                                else: elm.text = SM_PLOTS[31][13]	   
                        if elm.name == "LANDUSE_PLOT_33":
                                if len(SM_PLOTS) < 33: elm.delete()
                                else: elm.text = SM_PLOTS[32][13]	   
                        if elm.name == "LANDUSE_PLOT_34":
                                if len(SM_PLOTS) < 34: elm.delete()
                                else: elm.text = SM_PLOTS[33][13]	   
                        if elm.name == "LANDUSE_PLOT_35":
                                if len(SM_PLOTS) < 35: elm.delete()
                                else: elm.text = SM_PLOTS[34][13]
                        if elm.name == "LANDUSE_PLOT_36":
                                if len(SM_PLOTS) < 36: elm.delete()
                                else: elm.text = SM_PLOTS[35][13]
                        if elm.name == "LANDUSE_PLOT_37":
                                if len(SM_PLOTS) < 37: elm.delete()
                                else: elm.text = SM_PLOTS[36][13]
        #===========================================================================================================
        arcpy.AddMessage("Setting QR Code Image")
        #-----------------------------------------------------------------------------------------------------------
        if not os.path.exists(os.path.dirname(ARCHIVE_PATH)):
                os.makedirs(os.path.dirname(ARCHIVE_PATH))    
        if len(SM_PLOTS) > 37:
                Excel_PIC = arcpy.mapping.ListLayoutElements(mxd, "PICTURE_ELEMENT", "Excel")[0]
                excelimgname = None
                for subdirs, dirs, files in os.walk(ARCHIVE_PATH):
                        for name in files:
                                basename = os.path.splitext(name)[0]
                                if basename == REQNUM + "_EXCEL":
                                        excelimgname = name 
                if excelimgname <> None:                
                        imagepath = ARCHIVE_PATH + excelimgname
                        excelwidth = Excel_PIC.elementWidth
                        excelheight = Excel_PIC.elementHeight
                        result = [x.strip() for x in Excel.split(',')]
                        if(result is not None and len(result) == 2):
                                arcpy.AddMessage(result)
                                excelwidth = float(result[0])
                                excelheight = float(result[1])
                        else:
                                excelraster = arcpy.sa.Raster(imagepath)
                                excelheight = excelraster.height
                                excelwidth = excelraster.width
                                excelwidth = (excelwidth * 2.54 ) / 96.0
                                excelheight = (excelheight * 2.54 ) / 96.0
                                
                       
                        
                        Excel_PIC.sourceImage = imagepath
                        Excel_PIC.elementPositionX = 1
                        Excel_PICy = pDataFrame.elementPositionY
                        Excel_PICHeight = pDataFrame.elementHeight
                        Excel_PIC.elementPositionY = (Excel_PICy + Excel_PICHeight) - excelheight
                        Excel_PIC.elementWidth = excelwidth
                        Excel_PIC.elementHeight = excelheight
                        #arcpy.AddMessage("Excel Size requested " + str(excelwidth) + " X " + str(excelheight))
                        #arcpy.AddMessage("Excel Size dim " + str(Excel_PIC.elementWidth) + " X " + str(Excel_PIC.elementHeight))
                else:
                        Excel_PIC.elementWidth = 0
                        Excel_PIC.elementHeight = 0
        else:
                Excel_PIC = arcpy.mapping.ListLayoutElements(mxd, "PICTURE_ELEMENT", "Excel")[0]
                Excel_PIC.elementWidth = 0
                Excel_PIC.elementHeight = 0
                                
        Legend_PIC = arcpy.mapping.ListLayoutElements(mxd, "PICTURE_ELEMENT", "Legend")[0]
        legendimgname = None
        for subdirs, dirs, files in os.walk(ARCHIVE_PATH):
                for name in files:                        
                        basename = os.path.splitext(name)[0]
                        if basename == REQNUM + "_LEGEND":
                                legendimgname = name
        
        if legendimgname <> None:
                imagepath = ARCHIVE_PATH + legendimgname                
                Legendwidth = Legend_PIC.elementWidth
                Legendheight = Legend_PIC.elementHeight
                result = [x.strip() for x in Legend.split(',')]
                if(result is not None and len(result) == 2):
                        arcpy.AddMessage(result)
                        Legendwidth = float(result[0])
                        Legendheight = float(result[1])
                else:
                        Legendraster = arcpy.sa.Raster(imagepath)
                        Legendheight = Legendraster.height
                        Legendwidth = Legendraster.width
                        Legendwidth = (Legendwidth * 2.54 ) / 96.0
                        Legendheight = (Legendheight * 2.54 ) / 96.0   
              
                
                Legend_PIC.sourceImage = imagepath
                Legend_PICx = pDataFrame.elementPositionX
                Legend_PICWidth = pDataFrame.elementWidth
                Legend_PIC.elementPositionX = (Legend_PICx + Legend_PICWidth) - Legendwidth
                #Legend_PIC.elementPositionY = 5.6
                Legend_PIC.elementWidth = Legendwidth
                Legend_PIC.elementHeight = Legendheight
                #arcpy.AddMessage("Legend Size requested " + str(Legendwidth) + " X " + str(Legendheight))
                #arcpy.AddMessage("Legend image dim " + str(Legend_PIC.elementWidth) + " X " + str(Legend_PIC.elementHeight))
        else:
                Legend_PIC.elementWidth = 0
                Legend_PIC.elementHeight = 0
                
        #===========================================================================================================
        arcpy.AddMessage("Exporting SITEMAP to PDF")
        #-----------------------------------------------------------------------------------------------------------
        mxd.saveACopy(ARCHIVE_PATH + REQNUM + "_AP" + ".mxd")
        arcpy.mapping.ExportToPDF(mxd, ARCHIVE_PATH + REQNUM + "_AP" + ".PDF",resolution=600,embed_fonts=True, image_quality="NORMAL")
        arcpy.SetParameterAsText(2, ARCHIVE_PATH + REQNUM + "_AP" + ".PDF")
        arcpy.SetParameterAsText(3, "Success")
else:
        arcpy.SetParameterAsText(3, "Failure")



           if elm.name == "LANDUSE_PLOT_11":
                                if len(SM_PLOTS) < 11: elm.delete()
                                else:
                                        if(SM_PLOTS[10][7] is not None and SM_PLOTS[10][8] is not None):
                                                elm.text = SM_PLOTS[10][7] + "/" + SM_PLOTS[10][8]
                                        else: elm.text = " "
                                        	   
                        if elm.name == "LANDUSE_PLOT_12":
                                if len(SM_PLOTS) < 12: elm.delete()
                                else:
                                        if(SM_PLOTS[11][7] is not None and SM_PLOTS[11][8] is not None):
                                               elm.text = SM_PLOTS[11][7] + "/" + SM_PLOTS[11][8]
                                        else: elm.text = " "
                                        	   
                        if elm.name == "LANDUSE_PLOT_13":
                                if len(SM_PLOTS) < 13: elm.delete()
                                else:
                                        if(SM_PLOTS[12][7] is not None and SM_PLOTS[12][8] is not None):
                                                elm.text = SM_PLOTS[12][7] + "/" + SM_PLOTS[12][8]
                                        else: elm.text = " "
                                        	   
                        if elm.name == "LANDUSE_PLOT_14":
                                if len(SM_PLOTS) < 14: elm.delete()
                                else:
                                        if(SM_PLOTS[13][7] is not None and SM_PLOTS[13][8] is not None):
                                               elm.text = SM_PLOTS[13][7] + "/" + SM_PLOTS[13][8]
                                        else: elm.text = " "
                                        	   
                        if elm.name == "LANDUSE_PLOT_15":
                                if len(SM_PLOTS) < 15: elm.delete()
                                else:
                                        if(SM_PLOTS[14][7] is not None and SM_PLOTS[14][8] is not None):
                                                elm.text = SM_PLOTS[14][7] + "/" + SM_PLOTS[14][8]
                                        else: elm.text = " "
                                        	   
                        if elm.name == "LANDUSE_PLOT_16":
                                if len(SM_PLOTS) < 16: elm.delete()
                                else:
                                        if(SM_PLOTS[15][7] is not None and SM_PLOTS[15][8] is not None):
                                                elm.text = SM_PLOTS[15][7] + "/" + SM_PLOTS[15][8]
                                        else: elm.text = " "
                                        	   
                        if elm.name == "LANDUSE_PLOT_17":
                                if len(SM_PLOTS) < 17: elm.delete()
                                else:
                                        if(SM_PLOTS[16][7] is not None and SM_PLOTS[16][8] is not None):
                                               elm.text = SM_PLOTS[16][7] + "/" + SM_PLOTS[16][8]
                                        else: elm.text = " "
                                        	   
                        if elm.name == "LANDUSE_PLOT_18":
                                if len(SM_PLOTS) < 18: elm.delete()
                                else:
                                        if(SM_PLOTS[17][7] is not None and SM_PLOTS[17][8] is not None):
                                               elm.text = SM_PLOTS[17][7] + "/" + SM_PLOTS[17][8]
                                        else: elm.text = " "
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                               