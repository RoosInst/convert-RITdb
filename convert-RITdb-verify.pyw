#!/usr/bin/python3
# Requires tqdm module, install cmd: py - m pip install tqdm
'''
Convert Cassini Verify results from RITdb format to CSV and/or XLSX format(s)

Usage: convert-RITdb-verify.pyw -a | -i <input RITdb file>
'''
import sys
import os
import getopt
import subprocess
from tqdm import tqdm
from datetime import datetime

csvScriptFile = 'RITdbVerify2csv.pyw'
xlsxScriptFile = 'RITdbVerify2xlsx.pyw'
start = datetime.now()


def convertRITdb(dbFileName, csvOut, xlsxOut):
    # one file -i --ifile=
    if dbFileName:
        if csvOut and xlsxOut:
            subprocess.run(csvScriptFile + ' -i ' + dbFileName + ' & ' +
                           xlsxScriptFile + ' -i ' + dbFileName, shell=True)
        elif csvOut:
            subprocess.run(csvScriptFile + ' -i ' + dbFileName, shell=True)
        elif xlsxOut:
            subprocess.run(xlsxScriptFile + ' -i ' + dbFileName, shell=True)
    # all .ritdb files in *this* directory
    else:
        dbFiles = ([f for f in os.listdir()
                    if f.endswith('.ritdb') and os.path.isfile(os.path.join(f))])
        print("Found " + str(len(dbFiles)) +
              ' RITdb files in ' + os.getcwd() + "...")
        pbar = tqdm(dbFiles)
        for file in pbar:
            pbar.set_description("Converting %s" % file)
            if csvOut and xlsxOut:
                subprocess.run(csvScriptFile + ' -i ' + file + ' & ' +
                               xlsxScriptFile + ' -i ' + file, shell=True)
            elif csvOut:
                subprocess.run(csvScriptFile + ' -i ' + dbFileName, shell=True)
            elif xlsxOut:
                subprocess.run(xlsxScriptFile + ' -i ' +
                               dbFileName, shell=True)
    print("DONE in " + str(datetime.now()-start))
    sys.exit()


if __name__ == '__main__':
    usage = ('''
        Usage: %s [-a | --all] [-c | --csv] [-x | --xlsx] [-i <input RITdb file>]" % (sys.argv[0])
        
        Options and Arguments:
        -a, --all       convert all .ritdb files, ignoring input file if included
        -c, --csv       convert input file to CSV
        -x, --xlsx      convert input file to XLSX
        -i, --ifile=    .ritdb file name (assume both formats if -c and/or -x are missing)
        -h              this help, option and argument order matters
        ''')

    dbFileName = None
    csvOut = True
    xlsxOut = True

    try:
        opts, args = getopt.getopt(
            sys.argv[1:], "hai:o:", ["all", "csv", "xlsx", "ifile=", "ofile="])
    except getopt.GetoptError as err:
        print(err)
        usage()
        sys.exit(2)
    for opt, arg in opts:
        if opt == '-h':
            usage()
            sys.exit()
        elif opt in ("-a", "--all"):
            csvOut = True
            xlsxOut = True
        elif opt in ("-c", "--csv"):
            csvOut = True
            xlsxOut = False
        elif opt in ("-x", "--xlsx"):
            csvOut = False
            xlsxOut = True
        elif opt in ("-i", "--ifile"):
            dbFileName = arg.strip()

    convertRITdb(dbFileName, csvOut, xlsxOut)
