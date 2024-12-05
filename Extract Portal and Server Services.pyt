# -*- coding: utf-8 -*-
import arcpy
import concurrent.futures
import requests
import pandas as pd
from arcgis.gis import GIS
from tqdm import tqdm 
import uuid
import os
import sys
import tkinter as tk
from tkinter import simpledialog

cert_path = r'C:\Users\alteranit\Documents\Cer\CECONY ROOT CA2 2017.crt'

def get_portal_token(portal_url, username, password):
    token_url = f"{portal_url}/sharing/rest/generateToken"
    token_params = {
        "username": username,
        "password": password,
        "referer": portal_url,  
        "f": "json"  
    }
    
# Request the token
    response = requests.post(token_url, data=token_params ,verify= cert_path)
    
    if response.status_code == 200:
        token_data = response.json()
        if 'token' in token_data:
            return token_data['token']
        else:
            arcpy.AddError(f"get_portal_token Error: Failed to get token: {token_data.get('messages', 'No token returned')}")
            sys.exit(1)
            
    else:
        arcpy.AddError(f"get_portal_token Error: Failed to get token: HTTP {response.status_code}")
        sys.exit(1)
    


def fetch_portal_service_data(service):
    try:
        return {
            'Title': service.title,
            'URL': service.url,
            'Type': service.type,
            'Owner': service.owner,
            'Id': service.id,
            'Tags': ", ".join(service.tags), 
            'AccessType': service.access,
            'Size': service.size,
            'Created': service.created,
            'Modified': service.modified
        }
    except Exception as e:
        arcpy.AddError(f"fetch_portal_service_data Error: Title': {service.title}, 'Error': {e}")
        sys.exit(1)



# def get_portal_services(portal_url, username, password):
#   try:
#     gis = GIS(portal_url, username, password)
#     portal_items = gis.content.search(query="type:Service", max_items=3000)  # Limit to 3000 services
#     portal_services = []
#     arcpy.AddMessage("fetching portal services...")
#     with concurrent.futures.ThreadPoolExecutor(max_workers=20) as executor:
#         for service_data in tqdm(executor.map(fetch_portal_service_data, portal_items), total=len(portal_items), desc="Fetching Portal Services", ncols=100):
#             portal_services.append(service_data)
#     arcpy.AddMessage(f"portal_services{portal_services}")
#     return portal_services
#   except Exception as e:
#     arcpy.AddError(f"get_portal_services Error: {e}")
#     sys.exit(1)
    

def get_portal_services(portal_url, username, password):
    try:
        # Connect to the portal
        gis = GIS(portal_url, username, password)
        portal_items = gis.content.search(query="type:Service", max_items=3000)
        portal_services = []
        arcpy.AddMessage("Fetching portal services...")


        with concurrent.futures.ThreadPoolExecutor(max_workers=20) as executor:
            futures = list(executor.map(fetch_portal_service_data, portal_items))
            for service_data in tqdm(futures, total=len(portal_items), desc="Fetching Portal Services", ncols=100):
                portal_services.append(service_data)
        arcpy.AddMessage(f"Fetched {len(portal_services)} portal services.")

        return portal_services
    except Exception as e:
        arcpy.AddError(f"get_portal_services Error: {e}")
        sys.exit(1)









def get_ArcServer_services(server_url, portal_token, folder_name=None):
    
    "list all services in the root folder and subfolders."
    all_services = []
    folder_url = f"{server_url}/services"
    
    if folder_name:
        folder_url = f"{folder_url}/{folder_name}" 

    params = {"token": portal_token,"f": "json"}

    try:
        # Send request to list services in the folder
        response = requests.get(folder_url, params=params ,verify= cert_path)
        response.raise_for_status()
        services_data = response.json()

        if 'services' in services_data:
            services = services_data['services']
            for service in services:
                service_name = service['serviceName']
                #owner = service.get('provider')
                folderName = folder_name if folder_name else 'root'
                service_type = service['type']
                status_url =  f"{server_url}/services/{folder_name}/{service_name}.{service_type}/status" if folder_name else f"{server_url}/services/{service_name}.{service_type}/status"
                service_url = f"{server_url}/services/{folder_name}/{service_name}" if folder_name else f"{server_url}/services/{service_name}"
                status_response = requests.get(status_url,params=params ,verify= cert_path)

                all_services.append({
                    'Service_name': service_name,
                     'FolderName': folderName,                   
                    'Service_type': service_type,
                    'Status': status_response.json().get("realTimeState"),
                    'Service_url': service_url
                })
        else:
            arcpy.AddMessage(f"No 'services' found for folder: {folder_name if folder_name else 'root'}")

        if 'folders' in services_data:
            if services_data['folders']:
                for subfolder in services_data['folders']:
                    arcpy.AddMessage(f"Found subfolder: {subfolder}")
                    all_services.extend(get_ArcServer_services(server_url, portal_token, subfolder))  # Recurse into subfolder
            else:
                arcpy.AddMessage(f"No subfolders found for folder '{folder_name if folder_name else 'root'}'.")


    except requests.exceptions.RequestException as e:
        arcpy.AddError(f"Error fetching services from folder '{folder_name if folder_name else 'root'}': {e}")
        sys.exit(1)
    except KeyError as e:
        arcpy.AddError(f"Missing expected key: {e}")
        sys.exit(1)
    except Exception as e:
        arcpy.AddError(f"Unexpected error: {e}")
        sys.exit(1)
    return all_services



# Export the fetched service data to Excel
def export_to_excel(portal_data,server_data ,output_path,Env_name):
    unique_id =  uuid.uuid4().hex[:8]
    filename= f"{Env_name}_Protal+Server_services_{unique_id}.xlsx"
    output_Dir = os.path.join(output_path,filename)
    
    with pd.ExcelWriter(output_Dir) as writer:
        portal_df = pd.DataFrame(portal_data)
        portal_df.to_excel(writer, sheet_name='Portal', index=False)

        server_df = pd.DataFrame(server_data)
        server_df.to_excel(writer, sheet_name='Server', index=False)

    arcpy.AddMessage(f"Data has been exported to {output_Dir}")


class Toolbox:
    def __init__(self):
        """Define the toolbox (the name of the toolbox is the name of the
        .pyt file)."""
        self.label = "Extract Portal and Server Services"
        self.alias = "Extract Portal and Server Services"

        # List of tool classes associated with this toolbox
        self.tools = [Tool]


class Tool:

    def __init__(self):
        """Define the tool (tool name is the name of the class)."""
        self.label = "Extract Portal and Server Services"
        self.description = ""

    def getParameterInfo(self):
        """Define the tool parameters."""
        Environment = arcpy.Parameter( displayName="Select Environment", name="environment", parameterType="Required", direction="Input")
        Environment.filter.type = "ValueList"
        Environment.filter.list = ["Dev","Test","UAT","Pre-Prod"]
        username = arcpy.Parameter( displayName="Username", name="username", parameterType="Required", direction="Input")
        #password = arcpy.Parameter( displayName="Password", name="password",datatype="GPString", parameterType="Required", direction="Input")
        outputDirectory = arcpy.Parameter( displayName="OutputDirectory", datatype="DEFolder", name="outputDirectory", parameterType="Required", direction="Input")
        
        params = [Environment, username, outputDirectory]
        return params

    def isLicensed(self):
        """Set whether the tool is licensed to execute."""
        return True

    def updateParameters(self, parameters):
        """Modify the values and properties of parameters before internal
        validation is performed.  This method is called whenever a parameter
        has been changed."""
        return

    def updateMessages(self, parameters):
        """Modify the messages created by internal validation for each tool
        parameter. This method is called after internal validation."""
        return

    def get_password(self):
        # Create the main Tkinter window (it will be hidden)
        root = tk.Tk()
        root.withdraw()  # Hide the root window

        # Prompt the user to enter the password (masked as *****)
        password = simpledialog.askstring("Password", "Enter your password:", show="*")
        
        # Return the entered password, or None if the user cancels
        return password
    
    def execute(self, parameters, messages):
               

        Env_List = {"Dev":"gdeafmw1","Test":"gstwebadaptor","UAT":"gutweba1","Pre-Prod":"gppwebadaptor"}
        Env_Name =  Env_List.get(parameters[0].valueAsText) 
        portal_url = f"https://{Env_Name}.conedison.net/portal"
        server_url = f"https://{Env_Name}.conedison.net/server/admin"

        username = parameters[1].valueAsText
        password = self.get_password()

        if password:
            #arcpy.AddMessage(f"Password entered: {password}")
            arcpy.AddMessage(f"Password entered, Processing...")
        else:
            arcpy.AddError("No password entered.")
            sys.exit(1)
    
      
        output_path = parameters[2].valueAsText
        arcpy.AddMessage(f"portal_url {portal_url}")
        arcpy.AddMessage(f"server_url {server_url}")
        arcpy.AddMessage(f"username {username}")
        arcpy.AddMessage(f"output_path {output_path}")

        portal_services = get_portal_services(portal_url, username, password)


        portal_token = get_portal_token(portal_url, username, password)
        arcpy.AddMessage("Portal Token acquired successfully.\n")
        arcpy.AddMessage("Listing all services (including services in folders):")
        arcpy.AddMessage("Fetching Server services...")
        server_services = get_ArcServer_services(server_url, portal_token)
        arcpy.AddMessage(f"Fetched {len(server_services)} Server services.")
    
        arcpy.AddMessage("Exporting data to Excel...")
        export_to_excel(portal_services, server_services,output_path ,parameters[0].valueAsText)


        return

    def postExecute(self, parameters):
        """This method takes place after outputs are processed and
        added to the display."""
        return
    



