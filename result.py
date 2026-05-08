from inventory import *
import openpyxl

if __name__=="__main__":
    print("DO NOT CLOSE THE WINDOW UNTIL YOU FIND THE 'OK' MESSAGE")
    input_dict = fetch_input()
    data = []
    for key, value in input_dict.items():
        # print(key)
        # print(value)
        ilo_ip = key
        if ping_ip(ilo_ip):
            ilo_user = value["ILO_USERNAME"]
            ilo_password = value["ILO_PASSWORD"]
            #print(ilo_ip, ilo_user, ilo_password)
            print(f"Fetching the inventory from {ilo_ip}...")
            start_inventory(ilo_ip, ilo_user, ilo_password)
        else:
            print(f"!!!!!PLEASE VERIFY THE FOLLOWING IP: {ilo_ip}!!!!!")
    #print("=========================================================================================")
    #print("Printing from the rresult",data)
    print("Inventory fetched successfully!!")
    print("Path to 'Report': 'C:\\Inventory\\Report.txt'")
    print("Inventory.xlsx updated successfully")
    print("Now you are free to close the window!!!")
    print("OK")



