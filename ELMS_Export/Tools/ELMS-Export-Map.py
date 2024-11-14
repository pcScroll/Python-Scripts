# -*- ##################

import arcpy
import datetime
import uuid
import os,math,numpy
import json

#===========================================================================================================
arcpy.env.overwriteOutput = True
#===========================================================================================================
arcpy.AddMessage("Collecting input parameters")
#-----------------------------------------------------------------------------------------------------------


MapOutputType = arcpy.GetParameterAsText(0)
TempalteName = arcpy.GetParameterAsText(1)
MUNICIPALITYID = arcpy.GetParameterAsText(2)
ZONEID = arcpy.GetParameterAsText(3)
SECTORID = arcpy.GetParameterAsText(4)
PLOTID = arcpy.GetParameterAsText(5)
BUILDINGID = arcpy.GetParameterAsText(6)
FLOORID = arcpy.GetParameterAsText(7)
UNITID = arcpy.GetParameterAsText(8)
WIDTH= arcpy.GetParameterAsText(9)
HEIGHT= arcpy.GetParameterAsText(10)
DisplayDimensions = arcpy.GetParameterAsText(11)
Scale = arcpy.GetParameterAsText(12)
#===========================================================================================================
arcpy.AddMessage("Loading MapDocument")



#-----------------------------------------------------------------------------------------------------------
#Server Path
# DocumentPath = r"\\AD007APP06\\gisapplications\\arcgisdata1061\\ELMS_Export\\ELMS-Export.mxd"
# LayersPath  =  r"\\AD007APP06\\gisapplications\\arcgisdata1061\\ELMS_Export\\Layers\\"
# OutputPath  =  r"\\AD007APP06\\gisapplications\\arcgisdata1061\\ELMS_Export\\Output\\"
# TempPath  =  r"\\AD007APP06\gisapplications\arcgisdata1061\ELMS_Export\\Temp\\"




#Local Path
DocumentPath = r"D:\\MXDs\\ELMS_Export\\ELMS-Export.mxd"
LayersPath  =  r"D:\\MXDs\\ELMS_Export\\Layers\\"
OutputPath  =  r"D:\\MXDs\\ELMS_Export\\Output\\"
TempPath  =  r"D:\\MXDs\\ELMS_Export\\Temp\\"


mxd = arcpy.mapping.MapDocument(DocumentPath)
pDataFrame = arcpy.mapping.ListDataFrames(mxd, "Layers")[0]
pSpatialReference = pDataFrame.spatialReference
uid = str(uuid.uuid4().fields[-1])[:5]
#===========================================================================================================
arcpy.AddMessage(""+ TempalteName + " Info")
#-----------------------------------------------------------------------------------------------------------

TempalteName = "GetPlotMap"
MUNICIPALITYID = "1002"
ZONEID = "312"
SECTORID ="3112"
#PLOTID ="4246459"
PLOTID ="358011"
#PLOTID ="419166"
#BUILDINGID = "2196"
BUILDINGID = "424233"
FLOORID="223"
UNITID="554322"
WIDTH = 600
HEIGHT=600
MapOutputType="png"
DisplayDimensions = "True"
Scale = 500
#=========================================GetZoneMap======================================================
if TempalteName == "GetZoneMap":
        arcpy.AddMessage("Adding Map Layers..")
        ZONE_Layer = arcpy.mapping.Layer(LayersPath + "Zone\\Zone.lyr")
        arcpy.mapping.AddLayer(pDataFrame, ZONE_Layer, "BOTTOM")
        Sector_Layer = arcpy.mapping.Layer(LayersPath + "Sector\\Sector_Gray.lyr")
        arcpy.mapping.AddLayer(pDataFrame, Sector_Layer, "BOTTOM")
        if  ZONEID == "":
          arcpy.AddMessage(" ZONEID is required ")
          arcpy.SetParameterAsText(13, ""  )
          arcpy.SetParameterAsText(14, "Failed")
          arcpy.SetParameterAsText(15, "")
          arcpy.SetParameterAsText(16, "ZONEID is required")
          exit()


        query = "MUNICIPALITY_ID = '"+ MUNICIPALITYID +"'  AND UGISID = '" + ZONEID + "'"     
        Selected_Lyr = arcpy.mapping.Layer(LayersPath + "Zone\\SelectedZone.lyr") 
        Selected_Lyr.definitionQuery = query 
        rows = [row for row in arcpy.da.SearchCursor(Selected_Lyr,"*",query)]
        if len(rows) == 0:
          arcpy.AddMessage("No Result")
          arcpy.SetParameterAsText(13, "" )
          arcpy.SetParameterAsText(14, "Failed")
          arcpy.SetParameterAsText(15, "")
          arcpy.SetParameterAsText(16, " No Result! GetZoneMap: selected ZONEID is not exist ")
          exit()
        if SECTORID :
           query = "MUNICIPALITY_ID = '"+ MUNICIPALITYID +"'  AND UGISID = '" + SECTORID + "'  " 
           ##Add Sector layer
           Sector_Layer = arcpy.mapping.Layer(LayersPath + "Sector\\SelectedSector_Fill.lyr") 
           Sector_Layer.definitionQuery = query 
           arcpy.mapping.AddLayer(pDataFrame, Sector_Layer, "BOTTOM")
#==========================================================     
        arcpy.AddMessage("Setting map extent and scale")
        pDataFrame.extent = Selected_Lyr.getSelectedExtent() 

        if Scale != "":
          pDataFrame.scale = Scale
        else:          
          pDataFrame.scale *= 1.5
        arcpy.RefreshActiveView()       
#=========================================GetSectorMap======================================================      
if TempalteName == "GetSectorMap":

        arcpy.AddMessage("Adding Map Layers ..")
        ##Add sector layer
        Sector_Layer = arcpy.mapping.Layer(LayersPath + "Sector\\Sector.lyr")    
        arcpy.mapping.AddLayer(pDataFrame, Sector_Layer, "BOTTOM")  
        if  SECTORID == "":
          arcpy.AddMessage(" SECTORID is required ")
          arcpy.SetParameterAsText(13, ""  )
          arcpy.SetParameterAsText(14, "Failed")
          arcpy.SetParameterAsText(15, "")
          arcpy.SetParameterAsText(16, "SECTORID is required")
          exit()


        query = "MUNICIPALITY_ID = '"+ MUNICIPALITYID +"'  AND UGISID = " + SECTORID + " "         
        Selected_Lyr = arcpy.mapping.Layer(LayersPath + "Sector\\SelectedSector.lyr") 
        Selected_Lyr.definitionQuery = query 
        rows = [row for row in arcpy.da.SearchCursor(Selected_Lyr,"*",query)]
        ##check if there is result exist
        if len(rows) == 0:
          arcpy.AddMessage("No Result")
          arcpy.SetParameterAsText(13, "" )
          arcpy.SetParameterAsText(14, "Failed")
          arcpy.SetParameterAsText(15, "")
          arcpy.SetParameterAsText(16, " No Result! GetSectorMap: selected SECTORID is not exist ")
          exit()
        #arcpy.mapping.AddLayer(pDataFrame, Selected_Lyr, "BOTTOM")
        if PLOTID != "":
           query = "MUNICIPALITY_ID = '"+ MUNICIPALITYID +"' AND UGISID = " + PLOTID + "" 
           ##Add Plot layer
           Plot_Layer = arcpy.mapping.Layer(LayersPath + "Plot\\Plot.lyr") 
           Plot_Layer.definitionQuery = query 
           arcpy.mapping.AddLayer(pDataFrame, Plot_Layer, "BOTTOM")

            
        ##Add Road Centre Line layers
        RoadCentreLine_Layer = arcpy.mapping.Layer(LayersPath + "Road\\RoadCentreLine_Gray.lyr")
        arcpy.mapping.AddLayer(pDataFrame, RoadCentreLine_Layer, "BOTTOM")

        ##Add Road Edge layers
        RoadEdge_Layer = arcpy.mapping.Layer(LayersPath + "Road\\RoadEdge_Gray.lyr")
        arcpy.mapping.AddLayer(pDataFrame, RoadEdge_Layer, "BOTTOM")
#==========================================================    
        arcpy.AddMessage("Setting map extent and scale ..")
        pDataFrame.extent = Selected_Lyr.getSelectedExtent() 
        if Scale != "":
          pDataFrame.scale = Scale
        else:          
          pDataFrame.scale *= 1.5
        arcpy.RefreshActiveView()
#=========================================GetPlotMap======================================================
if TempalteName == "GetPlotMap":

        arcpy.AddMessage("Adding Map Layers ..")

        ##Add Road Edge layers
        RoadEdge_Layer = arcpy.mapping.Layer(LayersPath + "Road\\RoadEdge_Black.lyr")
        arcpy.mapping.AddLayer(pDataFrame, RoadEdge_Layer, "BOTTOM")

        ##Add Sector layer
        Sector_Layer = arcpy.mapping.Layer(LayersPath + "Sector\\Sector_Gray.lyr")    
        arcpy.mapping.AddLayer(pDataFrame, Sector_Layer, "BOTTOM")

        ##Add Zone layer
        Zone_Layer = arcpy.mapping.Layer(LayersPath + "Zone\\Zone_Gray.lyr")    
        arcpy.mapping.AddLayer(pDataFrame, Zone_Layer, "BOTTOM")

        ##Add CentreLine layer
        RoadCentreLine_Layer = arcpy.mapping.Layer(LayersPath + "Road\\RoadCentreLine_Black.lyr")
        arcpy.mapping.AddLayer(pDataFrame, RoadCentreLine_Layer, "BOTTOM")
      
        ##Add RoadEntrance layer
        RoadEntrance_Layer = arcpy.mapping.Layer(LayersPath + "Road\\RoadEntrance_Orange.lyr")
        arcpy.mapping.AddLayer(pDataFrame, RoadEntrance_Layer, "BOTTOM")

        ##Add RoadPavement layer
        RoadPavement_Layer = arcpy.mapping.Layer(LayersPath + "Road\\RoadPavement_Gray.lyr")
        arcpy.mapping.AddLayer(pDataFrame, RoadPavement_Layer, "BOTTOM")

        ##Add RoadParking layer
        RoadParking_Layer = arcpy.mapping.Layer(LayersPath + "Road\\RoadParking_blue.lyr")
        arcpy.mapping.AddLayer(pDataFrame, RoadParking_Layer, "BOTTOM")

        ##Add Landscape layer
        Landscape_Layer = arcpy.mapping.Layer(LayersPath + "Road\\Landscape_Green.lyr")
        arcpy.mapping.AddLayer(pDataFrame, Landscape_Layer, "BOTTOM")

        ##Add RoadLayby layer
        RoadLayby_Layer = arcpy.mapping.Layer(LayersPath + "Road\\RoadLayby_Green.lyr")
        arcpy.mapping.AddLayer(pDataFrame, RoadLayby_Layer, "BOTTOM")

        ##Add RoadHump layer
        RoadHump_Layer = arcpy.mapping.Layer(LayersPath + "Road\\RoadHump_Black.lyr")
        arcpy.mapping.AddLayer(pDataFrame, RoadHump_Layer, "BOTTOM")


        ##Add RoadCycleTrack layer
        RoadCycleTrack_Layer = arcpy.mapping.Layer(LayersPath + "Road\\RoadCycleTrack_brown.lyr")
        arcpy.mapping.AddLayer(pDataFrame, RoadCycleTrack_Layer, "BOTTOM")
        
        ##Add Plot outline layer
        Plot_Layer = arcpy.mapping.Layer(LayersPath + "Plot\\Plot_OutLine.lyr")     
        arcpy.mapping.AddLayer(pDataFrame, Plot_Layer, "BOTTOM")

        if  PLOTID == "":
          arcpy.AddMessage(" PLOTID is required ")
          arcpy.SetParameterAsText(13, ""  )
          arcpy.SetParameterAsText(14, "Failed")
          arcpy.SetParameterAsText(15, "")
          arcpy.SetParameterAsText(16, "PLOTID is required")
          exit()


        query = "MUNICIPALITY_ID = '"+ MUNICIPALITYID +"'  AND UGISID = '" + PLOTID + "'"     
        Selected_Lyr = arcpy.mapping.Layer(LayersPath + "Plot\\SelectedPlot.lyr") 
        Selected_Lyr.definitionQuery = query 
        rows = [row for row in arcpy.da.SearchCursor(Selected_Lyr,"*",query)]
        ##check if there is result exist
        if len(rows) == 0:
         arcpy.AddMessage("No Result")
         arcpy.SetParameterAsText(13, "" )
         arcpy.SetParameterAsText(14, "Failed")
         arcpy.SetParameterAsText(15, "")
         arcpy.SetParameterAsText(16, " No Result! GetPlotMap: selected PLOTID is not exist ")
         exit()
        arcpy.mapping.AddLayer(pDataFrame, Selected_Lyr, "BOTTOM")

#==========================================================    
        if DisplayDimensions.lower() == 'true':
          arcpy.AddMessage("Generate Plot Dimentions  ..")
          out_feature_Line = "PlotDimensions"
          out_feature_Points = "PlotPoints"
          arcpy.RefreshActiveView()
          in_memory_workspace = TempPath
          arcpy.env.workspace = in_memory_workspace         
          gdb_name =  uid +'.gdb'  
        #Create the geodatabase
        #gdb_File = arcpy.CreateFileGDB_management(OutputPath, gdb_name) 
          gdb_File = arcpy.CreateFileGDB_management(TempPath, gdb_name)
          if arcpy.Exists(TempPath+ gdb_name):
             
             template_polyline_fc = arcpy.CreateFeatureclass_management(gdb_File, out_feature_Line, "POLYLINE",spatial_reference=pSpatialReference)  
             template_points_fc = arcpy.CreateFeatureclass_management(gdb_File, out_feature_Points, "POINT",spatial_reference=pSpatialReference)
             Point_fnames = ["ID", "X", "Y"]  
             Point_fypes = ["LONG", "DOUBLE", "DOUBLE"]  
             for field_name, field_type in zip(Point_fnames, Point_fypes): arcpy.AddField_management(in_table=template_points_fc, field_name=field_name, field_type=field_type)
             Line_fnames = [ "Length"]  
             Line_fypes = ["DOUBLE"]  
             for field_name, field_type in zip(Line_fnames, Line_fypes): arcpy.AddField_management(in_table=template_polyline_fc, field_name=field_name, field_type=field_type)

             with arcpy.da.SearchCursor(Selected_Lyr, ["SHAPE@"]) as cursor:
              for row in cursor:
             # Get the geometry (SHAPE@) of the polygon feature
               polygon_geometry = row[0]
              # Check if the feature has a valid geometry
               if polygon_geometry:
            # Get the array of points (vertices) from the polygon
                points_array = polygon_geometry.getPart(0)
                point_feature_layers = []
            # Use an InsertCursor to insert each point into the output point feature class
                with arcpy.da.InsertCursor(template_points_fc, ["SHAPE@" ,"ID", "X" , "Y"]) as insert_cursor:
                 previous_point = None                
                 linesArr  =[]
                 for point in points_array:
                    insert_cursor.insertRow([point , point.ID, point.X , point.Y ])
                    if previous_point is not None:
                     if point != previous_point:
                        parray = arcpy.Array()
                        parray.add(arcpy.Point(previous_point.X, previous_point.Y))
                        parray.add(arcpy.Point(point.X, point.Y))
                        line = arcpy.Polyline(parray)
                        linesArr.append(line)         
                    previous_point = point

               with arcpy.da.InsertCursor(template_polyline_fc, ['SHAPE@' , 'Length']) as cursor:
                for line in linesArr: 
                 formatted_length = "{:.2f}".format(line.length)   
                 cursor.insertRow([line , formatted_length ])   
            # PlotDim_Lyr_path = '{}\\{}\\{}'.format(OutputPath, gdb_name ,out_feature_Line) 
             PlotDim_Lyr_path = '{}\\{}\\{}'.format(TempPath, gdb_name ,out_feature_Line) 
             PlotDimensions_lyr = arcpy.mapping.Layer(PlotDim_Lyr_path)         
             label_class = PlotDimensions_lyr.labelClasses[0]
             #label_class.expression = "[{}]".format('Length')
             label_class.expression   = '"{}" + [Length] +  "{}"'.format("<FNT size = '18'>","</FNT>") 
             PlotDimensions_lyr.showLabels = True
             arcpy.mapping.AddLayer(pDataFrame, PlotDimensions_lyr, "BOTTOM")
             
        #==========================================================    
        arcpy.AddMessage("Setting map extent and scale ..")
        pDataFrame.extent = Selected_Lyr.getSelectedExtent() 
        if Scale != "": 
         if Scale > pDataFrame.scale:
           pDataFrame.scale = Scale
         elif pDataFrame.scale >= 14000: 
           pDataFrame.scale += (pDataFrame.scale * 0.7)
         elif pDataFrame.scale >= 10000: 
           pDataFrame.scale += (pDataFrame.scale * 0.6)          
         elif pDataFrame.scale < 2000:         
           pDataFrame.scale = 2000
         else:pDataFrame.scale = 2000 
        elif pDataFrame.scale >= 14000: 
           pDataFrame.scale += (pDataFrame.scale * 0.7)
        elif pDataFrame.scale >= 10000: 
           pDataFrame.scale += (pDataFrame.scale * 0.6)
        elif pDataFrame.scale < 2000:         
           pDataFrame.scale = 2000
        else:pDataFrame.scale = 2000          
        arcpy.RefreshActiveView()

        if pDataFrame.scale <= 1000:
           Plot_Lyr = [layer for layer in arcpy.mapping.ListLayers(mxd) if layer.name == "Plot_OutLine"][0]  
           Plot_label_class = Plot_Lyr.labelClasses[0]
           Plot_label_class.expression   = '"{}" + [PLOTNUMBER] +  "{}"'.format("<FNT size = '10'>","</FNT>")                   
           if DisplayDimensions.lower() == 'true':
             PlotDim_lyr = [layer for layer in arcpy.mapping.ListLayers(mxd) if layer.name == "PlotDimensions"][0]                       
             PlotDim_label_class = PlotDim_lyr.labelClasses[0]
             PlotDim_label_class.expression   = '"{}" + [Length] +  "{}"'.format("<FNT size = '8'>","</FNT>")
           arcpy.RefreshActiveView() 
        
# #=====================================GetPlotThumbnail===================================================
if TempalteName == "GetPlotThumbnail":
      arcpy.AddMessage("Adding Map Layers ..")

      ##Add Road Edge layers
      RoadEdge_Layer = arcpy.mapping.Layer(LayersPath + "Road\\RoadEdge_Black.lyr")
      arcpy.mapping.AddLayer(pDataFrame, RoadEdge_Layer, "BOTTOM")
      ##Add Sector layer
      Sector_Layer = arcpy.mapping.Layer(LayersPath + "Sector\\Sector_Gray.lyr")    
      arcpy.mapping.AddLayer(pDataFrame, Sector_Layer, "BOTTOM")
      ##Add Zone layer
      Zone_Layer = arcpy.mapping.Layer(LayersPath + "Zone\\Zone_Gray.lyr")    
      arcpy.mapping.AddLayer(pDataFrame, Zone_Layer, "BOTTOM")
      ##Add CentreLine layer
      RoadCentreLine_Layer = arcpy.mapping.Layer(LayersPath + "Road\\RoadCentreLine_Black.lyr")
      arcpy.mapping.AddLayer(pDataFrame, RoadCentreLine_Layer, "BOTTOM")      
      ##Add RoadEntrance layer
      RoadEntrance_Layer = arcpy.mapping.Layer(LayersPath + "Road\\RoadEntrance_Orange.lyr")
      arcpy.mapping.AddLayer(pDataFrame, RoadEntrance_Layer, "BOTTOM")
      ##Add RoadPavement layer
      RoadPavement_Layer = arcpy.mapping.Layer(LayersPath + "Road\\RoadPavement_Gray.lyr")
      arcpy.mapping.AddLayer(pDataFrame, RoadPavement_Layer, "BOTTOM")
      ##Add RoadParking layer
      RoadParking_Layer = arcpy.mapping.Layer(LayersPath + "Road\\RoadParking_blue.lyr")
      arcpy.mapping.AddLayer(pDataFrame, RoadParking_Layer, "BOTTOM")
      ##Add Landscape layer
      Landscape_Layer = arcpy.mapping.Layer(LayersPath + "Road\\Landscape_Green.lyr")
      arcpy.mapping.AddLayer(pDataFrame, Landscape_Layer, "BOTTOM")
      ##Add RoadLayby layer
      RoadLayby_Layer = arcpy.mapping.Layer(LayersPath + "Road\\RoadLayby_Green.lyr")
      arcpy.mapping.AddLayer(pDataFrame, RoadLayby_Layer, "BOTTOM")
      ##Add RoadHump layer
      RoadHump_Layer = arcpy.mapping.Layer(LayersPath + "Road\\RoadHump_Black.lyr")
      arcpy.mapping.AddLayer(pDataFrame, RoadHump_Layer, "BOTTOM")
      ##Add RoadCycleTrack layer
      RoadCycleTrack_Layer = arcpy.mapping.Layer(LayersPath + "Road\\RoadCycleTrack_brown.lyr")
      arcpy.mapping.AddLayer(pDataFrame, RoadCycleTrack_Layer, "BOTTOM")      
      ##Add Plot outline layer
      Plot_Layer = arcpy.mapping.Layer(LayersPath + "Plot\\Plot_OutLine.lyr")     
      arcpy.mapping.AddLayer(pDataFrame, Plot_Layer, "BOTTOM")

      if  PLOTID == "":
          arcpy.AddMessage(" PLOTID is required ")
          arcpy.SetParameterAsText(13, ""  )
          arcpy.SetParameterAsText(14, "Failed")
          arcpy.SetParameterAsText(15, "")
          arcpy.SetParameterAsText(16, "PLOTID is required")
          exit()

      query = "MUNICIPALITY_ID = '"+ MUNICIPALITYID +"'  AND UGISID = '" + PLOTID + "'"     
      Selected_Lyr = arcpy.mapping.Layer(LayersPath + "Plot\\SelectedPlot.lyr") 
      Selected_Lyr.definitionQuery = query 
      rows = [row for row in arcpy.da.SearchCursor(Selected_Lyr,"*",query)]
      ##check if there is result exist
      if len(rows) == 0:
        arcpy.AddMessage("No Result")
        arcpy.SetParameterAsText(13, "" )
        arcpy.SetParameterAsText(14, "Failed")
        arcpy.SetParameterAsText(15, "")
        arcpy.SetParameterAsText(16, " No Result! GetPlotThumbnail: selected PLOTID is not exist ")
        exit()
      arcpy.mapping.AddLayer(pDataFrame, Selected_Lyr, "BOTTOM")

#==========================================================    
      if DisplayDimensions.lower() == 'true':
          
        arcpy.AddMessage("Generate Plot Dimentions  ..")
        in_memory_workspace = TempPath 
        arcpy.env.workspace = in_memory_workspace       
        out_feature_Line = "PlotDimensions"
        out_feature_Points = "PlotPoints"
        arcpy.RefreshActiveView()    
        gdb_name =  uid +'.gdb'  
        #Create the geodatabase
        #gdb_File = arcpy.CreateFileGDB_management(OutputPath, gdb_name) 
        gdb_File = arcpy.CreateFileGDB_management(in_memory_workspace, gdb_name)
        if arcpy.Exists(gdb_File):     
            template_polyline_fc = arcpy.CreateFeatureclass_management(gdb_File, out_feature_Line, "POLYLINE",spatial_reference=pSpatialReference)  
            template_points_fc = arcpy.CreateFeatureclass_management(gdb_File, out_feature_Points, "POINT",spatial_reference=pSpatialReference)
            Point_fnames = ["ID", "X", "Y"]  
            Point_fypes = ["LONG", "DOUBLE", "DOUBLE"]  
            for field_name, field_type in zip(Point_fnames, Point_fypes): arcpy.AddField_management(in_table=template_points_fc, field_name=field_name, field_type=field_type)
            Line_fnames = [ "Length"]  
            Line_fypes = ["DOUBLE"]  
            for field_name, field_type in zip(Line_fnames, Line_fypes): arcpy.AddField_management(in_table=template_polyline_fc, field_name=field_name, field_type=field_type)
            with arcpy.da.SearchCursor(Selected_Lyr, ["SHAPE@"]) as cursor:
             for row in cursor:
             # Get the geometry (SHAPE@) of the polygon feature
              polygon_geometry = row[0]
              # Check if the feature has a valid geometry
              if polygon_geometry:
            # Get the array of points (vertices) from the polygon
               points_array = polygon_geometry.getPart(0)
               point_feature_layers = []
            # Use an InsertCursor to insert each point into the output point feature class
               with arcpy.da.InsertCursor(template_points_fc, ["SHAPE@" ,"ID", "X" , "Y"]) as insert_cursor:
                previous_point = None                
                linesArr  =[]
                for point in points_array:
                    insert_cursor.insertRow([point , point.ID, point.X , point.Y ])
                    if previous_point is not None:
                     if point != previous_point:
                        parray = arcpy.Array()
                        parray.add(arcpy.Point(previous_point.X, previous_point.Y))
                        parray.add(arcpy.Point(point.X, point.Y))
                        line = arcpy.Polyline(parray)
                        linesArr.append(line)         
                    previous_point = point

               with arcpy.da.InsertCursor(template_polyline_fc, ['SHAPE@' , 'Length']) as cursor:
                for line in linesArr: 
                 formatted_length = "{:.2f}".format(line.length)   
                 cursor.insertRow([line , formatted_length ])   
           # PlotDim_Lyr_path = '{}\\{}\\{}'.format(OutputPath, gdb_name ,out_feature_Line) 
            PlotDim_Lyr_path = '{}\\{}\\{}'.format(TempPath, gdb_name ,out_feature_Line) 
            PlotDimensions_lyr = arcpy.mapping.Layer(PlotDim_Lyr_path)         
            label_class = PlotDimensions_lyr.labelClasses[0]
            #label_class.expression = "[{}]".format('Length')
            label_class.expression   = '"{}" + [Length] +  "{}"'.format("<FNT size = '18'>","</FNT>") 
            PlotDimensions_lyr.showLabels = True
            arcpy.mapping.AddLayer(pDataFrame, PlotDimensions_lyr, "BOTTOM")
           
        #==========================================================    
      arcpy.AddMessage("Setting map extent and scale ..")
      pDataFrame.extent = Selected_Lyr.getSelectedExtent() 
      if Scale != "": 
         if Scale > pDataFrame.scale:
           pDataFrame.scale = Scale
         elif pDataFrame.scale >= 14000: 
           pDataFrame.scale += (pDataFrame.scale * 0.7)
         elif pDataFrame.scale >= 10000: 
           pDataFrame.scale += (pDataFrame.scale * 0.6)
         elif pDataFrame.scale < 2000:         
           pDataFrame.scale = 2000
         else:pDataFrame.scale = 2000 
      elif pDataFrame.scale >= 14000: 
           pDataFrame.scale += (pDataFrame.scale * 0.7)
      elif pDataFrame.scale >= 10000: 
           pDataFrame.scale += (pDataFrame.scale * 0.6)
      elif pDataFrame.scale < 2000:         
           pDataFrame.scale = 2000
      else:pDataFrame.scale = 2000   
      arcpy.RefreshActiveView()
      if MapOutputType == "png":
       arcpy.AddMessage("Exporting MAP to PNG ..") 
       arcpy.mapping.ExportToPNG(mxd, OutputPath + uid + ".png",pDataFrame, df_export_width=int(WIDTH),  df_export_height=int(HEIGHT))
       arcpy.SetParameterAsText(13, uid  )
      else:
       arcpy.AddMessage("Exporting MAP to JPEG ..")
       arcpy.mapping.ExportToJPEG(mxd, OutputPath + uid + ".jpg",pDataFrame, df_export_width=int(WIDTH),  df_export_height=int(HEIGHT))
       arcpy.SetParameterAsText(13, uid )
      arcpy.AddMessage("Done")
      arcpy.SetParameterAsText(14, "Success")
      arcpy.SetParameterAsText(15, pDataFrame.scale)
      arcpy.SetParameterAsText(16, pDataFrame.extent.JSON)      
      exit()
  ##===========================================================================================================  
if TempalteName == "GetBuildingMap":
      arcpy.AddMessage("Adding Map Layers ..")
      ##Add Sector layer
      Sector_Layer = arcpy.mapping.Layer(LayersPath + "Sector\\Sector_Gray.lyr")    
      arcpy.mapping.AddLayer(pDataFrame, Sector_Layer, "BOTTOM")
      ##Add Plot outline layer
      Plot_Layer = arcpy.mapping.Layer(LayersPath + "Plot\\Plot_OutLine.lyr")     
      arcpy.mapping.AddLayer(pDataFrame, Plot_Layer, "BOTTOM")
      if BUILDINGID == "":
          arcpy.AddMessage(" BUILDINGID is required ")
          arcpy.SetParameterAsText(13, ""  )
          arcpy.SetParameterAsText(14, "Failed")
          arcpy.SetParameterAsText(15, "")
          arcpy.SetParameterAsText(16, "BUILDINGID is required")
          exit()

      query = "MUNICIPALITY_ID = '"+ MUNICIPALITYID +"'  AND BLD_UGISID = '" + BUILDINGID + "'"     
      Selected_Lyr = arcpy.mapping.Layer(LayersPath + "Building\\SelectedBuilding.lyr") 
      Selected_Lyr.definitionQuery = query 
      rows = [row for row in arcpy.da.SearchCursor(Selected_Lyr,"*",query)]
      ##check if there is result exist
      if len(rows) == 0:
        arcpy.AddMessage("No Result")
        arcpy.SetParameterAsText(13, "" )
        arcpy.SetParameterAsText(14, "Failed")
        arcpy.SetParameterAsText(15, "")
        arcpy.SetParameterAsText(16, " No Result! GetBuildingMap: selected BUILDINGID is not exist ")
        exit()
      arcpy.mapping.AddLayer(pDataFrame, Selected_Lyr, "BOTTOM")
      for row in rows:
        SECTORNUMBER = row[8]
        PLOTNUMBER = row[7]
      SelectedPlot_Layer = arcpy.mapping.Layer(LayersPath + "Plot\\SelectedPlot.lyr") 
      SelectedPlotQuery = "MUNICIPALITY_ID = '"+ MUNICIPALITYID +"'  AND SECTORTPSSNUMBER = '" + SECTORNUMBER + "' AND PLOTNUMBER = '" + PLOTNUMBER + "'"  
      SelectedPlot_Layer.definitionQuery = SelectedPlotQuery
      #arcpy.mapping.AddLayer(pDataFrame, SelectedPlot_Layer, "BOTTOM")
      arcpy.AddMessage("Setting map extent and scale ..")
      pDataFrame.extent = SelectedPlot_Layer.getSelectedExtent() 

      if Scale != "":
        pDataFrame.scale = Scale
      else:          
        pDataFrame.scale *= 1.5
      arcpy.RefreshActiveView()
# #=========================================GetUnitMap=================================================      
if TempalteName == "GetUnitMap":
     
      if  UNITID == "" or FLOORID  == "":
          arcpy.AddMessage("UNITID, FLOORID are required ")
          arcpy.SetParameterAsText(13, "" )
          arcpy.SetParameterAsText(14, "Failed")
          arcpy.SetParameterAsText(15, "")
          arcpy.SetParameterAsText(16, " UNITID, FLOORID are required")
          exit()

      arcpy.AddMessage("Adding Map Layers ..")
      ##Add Floor layer
      Floor_query = "MUNICIPALITY_ID = '"+ MUNICIPALITYID +"' AND FLR_UGISID = '" + FLOORID + "'"  
      Floor_Layer = arcpy.mapping.Layer(LayersPath + "Unit\\SelectedFloor.lyr") 
      Floor_Layer.definitionQuery = Floor_query
      Floorrows = [row for row in arcpy.da.SearchCursor(Floor_Layer,"*",Floor_query)]
      ##check if there is result exist
      if len(Floorrows) == 0:
        arcpy.AddMessage("No Result")
        arcpy.SetParameterAsText(13, "" )
        arcpy.SetParameterAsText(14, "Failed")
        arcpy.SetParameterAsText(15, "")
        arcpy.SetParameterAsText(16, " No Result! GetUnitMap: selected FLOORID is not exist ")
        exit()

      arcpy.mapping.AddLayer(pDataFrame, Floor_Layer, "BOTTOM")
      ##Add Selected Unit layer
      SelectedUnit_Layer = arcpy.mapping.Layer(LayersPath + "Unit\\SelectedUnit.lyr") 
      Unit_query = "MUNICIPALITY_ID = '"+ MUNICIPALITYID +"'  AND FLR_UGISID = '" + FLOORID + "' AND UNT_UGISID = '" + UNITID + "'"
      SelectedUnit_Layer.definitionQuery = Unit_query
      SelectedUnitrows = [row for row in arcpy.da.SearchCursor(SelectedUnit_Layer,"*",Unit_query)]
      ##check if there is result exist
      if len(SelectedUnitrows) == 0:
        arcpy.AddMessage("No Result")
        arcpy.SetParameterAsText(13, "" )
        arcpy.SetParameterAsText(14, "Failed")
        arcpy.SetParameterAsText(15, "")
        arcpy.SetParameterAsText(16, " No Result! GetUnitMap: selected UNITID is not exist ")
        exit()

      arcpy.mapping.AddLayer(pDataFrame, SelectedUnit_Layer, "BOTTOM")
      ##Add Unit layer
      Unit_Layer = arcpy.mapping.Layer(LayersPath + "Unit\\unit.lyr")
      Unit_Layer.definitionQuery = Floor_query  
      
      arcpy.AddMessage("Setting map extent and scale ..")
      pDataFrame.extent = Floor_Layer.getSelectedExtent() 
      Unitrows = [row for row in arcpy.da.SearchCursor(Floor_Layer,"*",Floor_query)]
      ##check if there is result exist
      if len(SelectedUnitrows) == 0:
        arcpy.AddMessage("No Result")
        arcpy.SetParameterAsText(13, "" )
        arcpy.SetParameterAsText(14, "Failed")
        arcpy.SetParameterAsText(15, "")
        arcpy.SetParameterAsText(16, " No Result! GetUnitMap: selected FLOORID is not exist ")
        exit()
      arcpy.mapping.AddLayer(pDataFrame, Unit_Layer, "BOTTOM")
      if Scale != "": 
        if Scale < pDataFrame.scale:
         pDataFrame.scale = Scale
      elif pDataFrame.scale <= 14000: 
           pDataFrame.scale += 10000
      elif pDataFrame.scale <= 10000: 
           pDataFrame.scale += 6000
      else:          
        pDataFrame.scale *= 3
      arcpy.RefreshActiveView()
# #=======================================GetPlotDimensionsString================================================
if TempalteName == "GetPlotDimensionsString":
     if  PLOTID == "":
           arcpy.AddMessage(" PLOTID is required ")
           arcpy.SetParameterAsText(13, ""  )
           arcpy.SetParameterAsText(14, "Failed")
           arcpy.SetParameterAsText(15, "")
           arcpy.SetParameterAsText(16, "PLOTID is required ")
           exit()

     query = "MUNICIPALITY_ID = '"+ MUNICIPALITYID +"'  AND UGISID = '" + PLOTID + "'"     
     Selected_Lyr = arcpy.mapping.Layer(LayersPath + "Plot\\SelectedPlot.lyr") 
     Selected_Lyr.definitionQuery = query 
     rows = [row for row in arcpy.da.SearchCursor(Selected_Lyr,"*",query)]
        ##check if there is result exist
     if len(rows) == 0:
       arcpy.AddMessage("No Result")
       arcpy.SetParameterAsText(13, "")
       arcpy.SetParameterAsText(14, "Failed")
       arcpy.SetParameterAsText(15, "")
       arcpy.SetParameterAsText(16, " No Result! GetPlotDimensionsString: selected PLOTID is not exist ")
       exit()
     arcpy.AddMessage("Generate Plot Dimentions  ..")
     in_memory_workspace =  TempPath
     arcpy.env.workspace = in_memory_workspace       
     out_feature_Line = "PlotDimensions"
     out_feature_Points = "PlotPoints"
     arcpy.RefreshActiveView()    
     gdb_name =  uid +'.gdb'  
     gdb_File = arcpy.CreateFileGDB_management(in_memory_workspace, gdb_name)
     if arcpy.Exists(gdb_File):     
            template_polyline_fc = arcpy.CreateFeatureclass_management(gdb_File, out_feature_Line, "POLYLINE",spatial_reference=pSpatialReference)  
            template_points_fc = arcpy.CreateFeatureclass_management(gdb_File, out_feature_Points, "POINT",spatial_reference=pSpatialReference)
            Point_fnames = ["ID", "X", "Y"]  
            Point_fypes = ["LONG", "DOUBLE", "DOUBLE"]  
            for field_name, field_type in zip(Point_fnames, Point_fypes): arcpy.AddField_management(in_table=template_points_fc, field_name=field_name, field_type=field_type)
            Line_fnames = [ "LengthM" ,"LengthF"]  
            Line_fypes = ["DOUBLE" ,"DOUBLE"]  
            for field_name, field_type in zip(Line_fnames, Line_fypes): arcpy.AddField_management(in_table=template_polyline_fc, field_name=field_name, field_type=field_type)
            with arcpy.da.SearchCursor(Selected_Lyr, ["SHAPE@"]) as cursor:
             for row in cursor:
              polygon_geometry = row[0]
              if polygon_geometry:
            # Get the array of points (vertices) from the polygon
               points_array = polygon_geometry.getPart(0)
               point_feature_layers = []
            # Use an InsertCursor to insert each point into the output point feature class
               with arcpy.da.InsertCursor(template_points_fc, ["SHAPE@" ,"ID", "X" , "Y"]) as insert_cursor:
                previous_point = None                
                linesArr  =[]
                for point in points_array:
                    insert_cursor.insertRow([point , point.ID, point.X , point.Y ])
                    if previous_point is not None:
                     if point != previous_point:
                        parray = arcpy.Array()
                        parray.add(arcpy.Point(previous_point.X, previous_point.Y))
                        parray.add(arcpy.Point(point.X, point.Y))
                        line = arcpy.Polyline(parray)
                        linesArr.append(line)         
                    previous_point = point
               LengArr = []
               with arcpy.da.InsertCursor(template_polyline_fc, ['FID@','SHAPE@' , 'LengthM' ,'LengthF']) as cursor:
                for  line in linesArr: 
                 formatted_length = "{:.2f}".format(line.length)
                 LengthFeet =  line.length * 10.7639
                 formatted_lengthFeet = "{:.2f}".format(LengthFeet)                 
                 obj = {'ID': linesArr.index(line), 'LengthM': formatted_length, 'LengthF': formatted_lengthFeet}
                 LengArr.append(obj)          
        #==========================================================    
     arcpy.AddMessage("Done")      
     arcpy.SetParameterAsText(13, uid)
     arcpy.SetParameterAsText(14, "Success")
     arcpy.SetParameterAsText(16, json.dumps(LengArr))
     del mxd ,in_memory_workspace
     exit()


#-----------------------------------------------------------------------------------------------------------

# mxd.saveACopy(OutputPath + TempalteName + ".mxd")
# arcpy.mapping.ExportToPDF(mxd, OutputPath + TempalteName + ".PDF",resolution=800,embed_fonts=True, image_quality="NORMAL")

dpi = int(WIDTH) / 11

page_layout = mxd.activeDataFrame

# Set the custom page size (in inches)
page_layout.pageSize = (int(WIDTH) / 96, int(HEIGHT) /96)

mxd.saveACopy(OutputPath + uid + ".mxd")
if MapOutputType == "png":
    arcpy.AddMessage("Exporting MAP to PNG ..") 
    arcpy.mapping.ExportToPNG(mxd, OutputPath + uid + ".png",pDataFrame, df_export_width=int(WIDTH),  df_export_height=int(HEIGHT))
    #arcpy.mapping.ExportToPNG(mxd, OutputPath + uid + ".jpg", data_frame="PAGE_LAYOUT", resolution = dpi )
    arcpy.SetParameterAsText(13, uid  )
else:
    arcpy.AddMessage("Exporting MAP to JPEG ..")
    arcpy.mapping.ExportToPNG(mxd, OutputPath + uid + ".png",pDataFrame, df_export_width=int(WIDTH),  df_export_height=int(HEIGHT))
    #arcpy.mapping.ExportToJPEG(mxd, OutputPath + uid + ".jpg", data_frame="PAGE_LAYOUT", resolution = dpi ) # 
    arcpy.SetParameterAsText(13, uid )

arcpy.AddMessage("Done")
arcpy.SetParameterAsText(14, "Success")
arcpy.SetParameterAsText(15, pDataFrame.scale)
arcpy.SetParameterAsText(16, pDataFrame.extent.JSON)


# Clean up resources
del mxd 















