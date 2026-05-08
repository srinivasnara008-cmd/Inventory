from inventory import *
import openpyxl
from datetime import datetime

if __name__=="__main__":
    print("="*100)
    print("INVENTORY COLLECTION SYSTEM - STARTING")
    print("="*100)
    print("DO NOT CLOSE THE WINDOW UNTIL YOU FIND THE 'OK' MESSAGE")
    print()
    
    try:
        # Fetch input from Excel
        print("Step 1: Loading system credentials...")
        input_dict = fetch_input()
        
        if not input_dict:
            print("✗ No systems found in System_IPs.xlsx")
            sys.exit(1)
        
        print(f"✓ Found {len(input_dict)} system(s) to scan\n")
        
        # Initialize counters
        total_systems = len(input_dict)
        successful_scans = 0
        failed_scans = 0
        failed_ips = []
        
        # Process each system
        print("Step 2: Scanning systems...\n")
        for ilo_ip, credentials in input_dict.items():
            print("-" * 100)
            print(f"Processing: {ilo_ip}")
            print("-" * 100)
            
            # Verify connectivity
            if ping_ip(ilo_ip):
                try:
                    ilo_user = credentials["ILO_USERNAME"]
                    ilo_password = credentials["ILO_PASSWORD"]
                    
                    print(f"Fetching inventory from {ilo_ip}...")
                    result = start_inventory(ilo_ip, ilo_user, ilo_password)
                    
                    if result:
                        print(f"✓ Successfully fetched inventory for {ilo_ip}\n")
                        successful_scans += 1
                    else:
                        print(f"✗ Failed to fetch inventory for {ilo_ip}\n")
                        failed_scans += 1
                        failed_ips.append(ilo_ip)
                except Exception as e:
                    print(f"✗ Error processing {ilo_ip}: {e}\n")
                    failed_scans += 1
                    failed_ips.append(ilo_ip)
            else:
                print(f"✗ Cannot reach {ilo_ip} - skipping\n")
                failed_scans += 1
                failed_ips.append(ilo_ip)
        
        # Summary Report
        print("\n" + "="*100)
        print("INVENTORY COLLECTION SUMMARY")
        print("="*100)
        print(f"Total Systems: {total_systems}")
        print(f"✓ Successful Scans: {successful_scans}")
        print(f"✗ Failed Scans: {failed_scans}")
        
        if failed_ips:
            print(f"\nFailed IPs: {', '.join(failed_ips)}")
            print("Please verify credentials and network connectivity for failed systems.")
        
        print("\n" + "="*100)
        print("OUTPUT FILES GENERATED:")
        print("="*100)
        print(f"1. Report.txt - Detailed inventory report")
        print(f"2. Inventory.xlsx - Structured inventory data with {11} fields:")
        print("   - ILO_IP, Platform, Controller_Name, Controller_Model")
        print("   - Controller_Firmware, Drive_Model, Drive_Serial")
        print("   - Drive_Capacity, Drive_Status, Drive_Firmware, Drive_Type")
        print("="*100)
        print("\nInventory collection completed successfully!!")
        print("OK")
        
    except KeyboardInterrupt:
        print("\n✗ Process interrupted by user")
        sys.exit(1)
    except Exception as e:
        print(f"\n✗ Unexpected error: {e}")
        sys.exit(1)
