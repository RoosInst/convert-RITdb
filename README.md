# RITdbVerify Conversion Scripts

Coverts Cassini RITdb Verify Datalogs to .csv and/or .xlsx format(s).

## Input RITdb Types

### RITdb Verify (Cassini Diag/Cal Exec)

Cassini generates Guru Objects with "ri.sys.Type=verfy" that can be copied via Guru Agent to the filesystem.  Enter the desired conversion script to the 'Post execution script" to be run immediately after being copied.

### RITdb Datalog (Cassini Test Exec)

The Cassini application creates RITdb files in the D:\RiApps\Data\ folder that is used by the Worksheet.  Cassini with Patch #TBD saves the RITdb datalog to Guru as ObjClass=RITdb.datalog in distinct pieces also called "windows" (type = begin, window, end). Guru Agent can be used to combine the pieces (type = begin, window, end) and all any of the conversion scripts with the *Post Execution Script*.  (Requires Guru Agent Editor v62+)

### Convert Wrapper

Utility script that performs conversion on every .ritdb in the same directory to both formats by default.  Use command options to specify a specific format (-c or -x).

## Output File Types

### CSV

Converts to comma separated .csv format.  "SystemCheck" limits are applied and Failed tests counted and added to the last line.

### XLSX

Converts to Excel .xlsx format. "SystemCheck" limits are applied and Failed tests counted and added to the last line.  Failed tests and the column header that contains the file name has a 'RED' style applied.

Requires openpyxl 'python3 -m pip install openpyxl'

## Installation & Use

1. Download this progect to a folder
2. run pip install --no-index --find-links /path/to/download/dir/ -r requirements.txt

Copy the .py (or .pyw Windows) file(s) to the same directory where the .ritdb files are located.  "convert-RITdb-verify" files require both "RITdbVerify2csv" and "RITdbVerify2xlsx" files for full functionaliy.

Linux and MacOS- only use .py files
Windows - only use .pyw files to suppress console window

Example Options and Arguments: (not available on every script)

    p | --pivot : transpose the rows and columns of the table
    s | --split : splits into separate output files per Wafer
    i | --ifile= : Input file name (required)
    o | --ofile= : Outpur file name

### Use Example

Converts all the .ritdb files in the current directory to both .csv and .xlsx files

    convert-RITdb-verify.pyw 

Converts the file RITdb-verify-samples/G9TQQ5YA-Ri7421B_DPVP_Verify_CF2_2021-11-30T19-21-46.ritdb file to .csv

    RTIdbVerify2csv.pyw -i RITdb-verify-samples/G9TQQ5YA-Ri7421B_DPVP_Verify_CF2_2021-11-30T19-21-46.ritdb

## Developers

Push requests and issues are welcome.  <https://github.com/RoosInst/convert-RITdb/issues>

### Python3 Setup

Run "python3 pip install -r requirements.txt" install or "python3 pip install --upgrade -r requirements.txt"  to upgrade.

Visual Studio Code configs included.
