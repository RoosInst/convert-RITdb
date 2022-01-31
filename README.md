# RITdbVerify Conversion Scripts

Coverts Cassini RITdb Verify Datalogs to other common desktop formats. (Requires Python3)

# Input Types

### RITdb Verify (Cassini Diag/Cal Exec)

Cassini generates Guru Objects with "ri.sys.Type=verfy" that can be copied via Guru Agent to the filesystem.  Enter the desired conversion script to the 'Post execution script" to be run immediately after being copied. 

### RITdb Datalog (Cassini Test Exec)

(Coming soon)
RITdb files are created at runtime by Cassini in the D:\Datalog folder.

Cassini Patch #TBD stores RITdb pieces in Guru.  Guru Agent can be used to combine the pieces (type = begin, window, end) and all any of the conversion scripts with the

## Output Files Types

### CSV

Converts to comma separated .csv format

### XLSX

Converts to Excel .xlsx format

## Verify Datalogs

Cassini Verify logs use the scripts that start with "RITdbVerify..."

## Convert Wrapper

Utility script that performs conversion on every .ritdb in the same directory to both formats by default.  Use command options to specify a specific format (-c or -x).

## Installation & Use

Copy the files to the same directory where the .ritdb files are located.

Linux and MacOS- use .py files
Windows - use .pyw files to suppress console window

Example Options and Arguments: (not available on every script)
- p | --pivot : transpose the rows and columns of the table
- s | --split : splits into separate output files per Wafer
- i | --ifile= : Input file name
- o | --ofile= : Outpur file name

# Developers

## Python3 Setup

Run "python3 pip install -r requirements.txt" install or "python3 pip install --upgrade -r requirements.txt"  to upgrade.