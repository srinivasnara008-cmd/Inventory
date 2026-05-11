## Excel Setup Guide - System_IPs.xlsx

### Overview
The `System_IPs.xlsx` file is the input file for the Inventory Collection System. It must contain ILO/iDRAC credentials with proper column headers.

### Required Column Headers

| Column Name | Description | Example |
|-------------|-------------|---------|
| **ILO_IP** | IP address of the iLO/iDRAC interface | 192.168.1.100 |
| **ILO_USERNAME** | Username for iLO/iDRAC authentication | admin |
| **ILO_PASSWORD** | Password for iLO/iDRAC authentication | password123 |

### Quick Start

#### Step 1: Create Template File
Run the header setup script to automatically create the template:

```bash
python add_headers.py
```

This will:
- ✅ Create `System_IPs.xlsx` with proper headers (if it doesn't exist)
- ✅ Add headers to existing file (if headers are missing)
- ✅ Apply professional formatting (bold headers, borders, colors)
- ✅ Provide sample data to get you started

#### Step 2: Edit the File
Open `System_IPs.xlsx` in:
- Microsoft Excel
- LibreOffice Calc
- Google Sheets

Replace the sample data with your actual server credentials:

| ILO_IP | ILO_USERNAME | ILO_PASSWORD |
|--------|--------------|--------------|
| 192.168.1.100 | admin | mypassword1 |
| 192.168.1.101 | admin | mypassword2 |
| 192.168.1.102 | admin | mypassword3 |

#### Step 3: Save and Run Inventory
```bash
python result.py
```

### File Format Details

**Location:** Root directory of the repository  
**Format:** Excel (.xlsx)  
**Encoding:** UTF-8  
**Sheet Name:** Can be any name (defaults to "Sheet1" or "Credentials")

### Important Notes

⚠️ **Security Warning:**
- Do NOT commit `System_IPs.xlsx` to Git if it contains real passwords
- Add it to `.gitignore` if sensitive
- Store credentials securely

✅ **Best Practices:**
- Use dedicated iLO/iDRAC service accounts
- Use strong, unique passwords
- Restrict file access on your system
- Keep the file on a secure, encrypted drive

### Troubleshooting

**Issue: Script says "Missing required column: ILO_IP"**
- Solution: Run `python add_headers.py` to add headers properly

**Issue: "Cannot connect to ILO IP"**
- Verify IP address is correct
- Check network connectivity with `ping <ILO_IP>`
- Verify credentials are correct

**Issue: File encoding problems**
- Ensure file is saved as `.xlsx` (Excel format)
- Not `.xls` or `.csv`
- Use Excel or LibreOffice to save

### Example Output Structure

After running `python result.py`, the system creates:

**1. Inventory.xlsx** - Contains structured data:
- ILO_IP
- Platform
- Controller_Name
- Controller_Model
- Controller_Firmware
- Drive_Model
- Drive_Serial
- Drive_Capacity
- Drive_Status
- Drive_Firmware
- Drive_Type

**2. Report.txt** - Contains detailed text report

### Related Files

- `inventory.py` - Main inventory collection logic (reads from System_IPs.xlsx)
- `result.py` - Orchestrator script
- `add_headers.py` - Template creation and formatting script
- `Inventory.xlsx` - Output file (generated automatically)
- `Report.txt` - Detailed text output

### Support

For more information, refer to:
- `0_user_manual.docx` - Complete user guide
- Project README on GitHub

---

**Last Updated:** 2026-05-11  
**Version:** 1.0
