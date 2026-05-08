import requests
from requests.auth import HTTPBasicAuth
import json
import urllib3
import subprocess
import os
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, Border, Side

def ping_ip(ip_address):
    try:
        # Run the ping command, '-c' is for the number of packets (works on Linux/macOS)
        # '-n' is used for Windows to specify number of packets
        response = subprocess.run(
            ["ping", "-n", "4", ip_address],  # Use "-c 4" for Linux/macOS
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True
        )
        
        # Check the return code to see if the ping was successful
        if response.returncode == 0:
            print(f"Ping to {ip_address} successful :)")
            return 1
            # print(response.stdout)
        else:
            print(f"Ping to {ip_address} failed :(")
            return 0
            # print(response.stderr)
    except Exception as e:
        print(f"An error occurred: {e}")


def fetch_input():
    file_path = "C:\Users\Naveenc\Desktop\Inventory\System_IPs.xlsx"
    df = pd.read_excel(file_path)
    suts = df.head
    #result = df[[ 'ILO_IP', 'ILO_USERNAME', 'ILO_PASSWORD']].values.tolist()
    result_dict = df.set_index('ILO_IP')[['ILO_USERNAME', 'ILO_PASSWORD']].to_dict(orient='index')
    # print(result_dict)
    return result_dict

def risget(url, ilo_user, ilo_password):
    # Suppress InsecureRequestWarning
    try:
        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
        response =  requests.get(url, auth=HTTPBasicAuth(str(ilo_user),str(ilo_password)),
                            verify=False, timeout=60)
        print(response.status_code)
        return response
    except:
        print("[Requests || risget]: Please check the SUT details.")

def succuess_response(response):
    if response is not None:
        if response.status_code == 200 or response.status_code == 201:
            return 1
    else:
        print(f"response is getting empty please check manually once")
        return 0

def get_platform(response):
    if succuess_response(response):
        obj_dict = json.loads(response.text)
        if obj_dict.get("Model", 0):
            platform = obj_dict.get("Model")
            #print("Platform: ", platform)
        #else:
            #print(obj_dict.get("Model", 0))
        return platform
    else:
        return "NA"

def get_storage_id_count(response):
    if succuess_response(response):
        obj_dict = json.loads(response.text)
        count = obj_dict.get("Members@odata.count", 0)
        if count:
            members = obj_dict.get("Members")
            # print(members)
            return members
        else:
            return None

def get_controller_name(response, file):
    if succuess_response(response):
        ctrl_name = []
        obj_dict = json.loads(response.text)
        controller_name = obj_dict.get("Name", "No Controller detected")
        file.write("-------------------------------------\n-------------------------------------\nController Name: " + controller_name + "\n-------------------------------------\n")
        ctrl_name.append(controller_name)
        return ctrl_name
    

def get_drives_list(response):
    if succuess_response(response):
        obj_dict = json.loads(response.text)
        drives_list = obj_dict.get("Drives")
        if not drives_list:
            return 0
        else:
            drive_list = []
            for drive in drives_list:
                drive_list.append(drive["@odata.id"])
#            print("No.of Drives attached:", len(drive_list))
            return drive_list
                        

def fetch_drive_details(response, file):
    if succuess_response(response):
        obj_dict = json.loads(response.text)
        drive_model = obj_dict.get("Model", "NA")
        drive_sn = obj_dict.get("SerialNumber", "NA")
        drive_data =  drive_model + " || " + drive_sn
        file.write("------------------------------\n"+drive_data + "\n")
        return drive_data


    
drive_list_temp = []
def start_inventory(ilo_ip, ilo_user, ilo_password):
    # args = sys.argv[1:]
    # print(args)
    # ilo_ip = args[0]
    # ilo_user = args[1]
    # ilo_password = args[2]
    file = open("Report.txt","a")
    controller_list, controller_name, drive_count, drive_list = [], [], [], []
    ilo_ip_l = [ilo_ip]
    
    https = "https://"
    base = "/redfish/v1/Systems/1"
    url = https + ilo_ip + base

    response = risget(url, ilo_user, ilo_password)

    platform = get_platform(response)
    platform_l = [platform]
    file.write("\nXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX\n")
    file.write("Data from the following server: \n")
    file.write("ILO IP: "+ ilo_ip +" || Platform: "+ platform + "\n")
    response = risget(url + "/Storage", ilo_user, ilo_password)
    storage_id_list = get_storage_id_count(response)
    
    if storage_id_list:
        for sid in storage_id_list:
            # print("id",id)
            url = https + ilo_ip +  sid["@odata.id"]
            #print(url)
            response = risget(url, ilo_user, ilo_password)
            if response is not None:
                controller_name = get_controller_name(response, file)
            else:
                print("Storage "+ url +"  getting empty")
            #print("Controller Name:", controller_name)
            file.write("Drive Model & Serial Numbers:- \n------------------------------\n")
            file.write("Model Number  || Serial Number" + "\n")
            drives = get_drives_list(response)
            drive_count.clear()
            if drives:
                drive_count.append(str(len(drives)))

                drive_list.clear()
                for drive in drives:
                    url = https + ilo_ip +  drive
                    #print(url)
                    response = risget(url, ilo_user, ilo_password)
                    drive_data = fetch_drive_details(response, file)
                    drive_list_temp.append(drive_data)
                drive_list.append(drive_list_temp)
            else:
               drive_count.append(str(drives))
               file.write("No drives detected\n")
            try:
                controller_list.append(controller_name + drive_count + drive_list)
            except:
                print("FAILED TO FETCH DETAILS TRY AGAIN..")
    sut_details = ilo_ip_l + platform_l + controller_list
    save_to_inventory_excel(
        ilo_ip=ilo_ip,
        platform=platform,
        controller_name=controller_name[0] if controller_name else "NA",
        drives=drive_list_temp
    )
    return sut_details


def save_to_inventory_excel(ilo_ip, platform, controller_name, drives, filename="Inventory.xlsx"):
    rows = []
    for drive in drives:
        if isinstance(drive, list) and len(drive) == 2:
            model, serial = drive
        else:
            try:
                model, serial = drive.split(" || ")
            except Exception:
                model, serial = str(drive), "NA"
        rows.append([ilo_ip, platform, controller_name, model, serial])
    df = pd.DataFrame(rows, columns=["ILO_IP", "Platform", "Controller_Name", "Model_Number", "Serial_Number"])
    if not os.path.exists(filename):
        # First run → create file with header
        df.to_excel(filename, index=False)
    else:
        # Append new rows without header
        with pd.ExcelWriter(filename, engine="openpyxl", mode="a", if_sheet_exists="overlay") as writer:
            df.to_excel(writer, index=False, header=False, startrow=writer.sheets["Sheet1"].max_row)
    # Formatting after writing
    wb = load_workbook(filename)
    ws = wb.active
    header_font = Font(bold=True)
    align_center = Alignment(horizontal="center", vertical="center")
    thin_border = Border(left=Side(style="thin"),
                         right=Side(style="thin"),
                         top=Side(style="thin"),
                         bottom=Side(style="thin"))
    # Format header row
    for col in ws[1]:
        col.font = header_font
        col.alignment = align_center
        col.border = thin_border
    # Format data cells
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=5):
        for cell in row:
            cell.alignment = align_center
            cell.border = thin_border
    # Auto column width
    for col in ws.columns:
        max_length = 0
        col_letter = col[0].column_letter
        for cell in col:
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = max_length + 2
    wb.save(filename)

