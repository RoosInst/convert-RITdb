#!/usr/bin/python3
#
'''
Convert Cassini Verify results from RITdb to CSV & XLSX formats
requires tqdm module, install cmd: py -m pip install tqdm

Usage: RITdb-verifies2both.pyw -a | -i <input RITdb file>
'''
import sys
import os
import getopt
from tqdm import trange
# import RITdbVerify2csv
# import RITdbVerify2xlsx
import subprocess

import time
import sys

if __name__ == '__main__':
    usage = (
        "Usage: %s [-a | --all] [-i <input RITdb file>]" % (sys.argv[0]))

    dbFileName = ''
    csvFileName = ''
    xlsxFileName = ''

    try:
        opts, args = getopt.getopt(
            sys.argv[1:], "hai:o:", ["all", "ifile=", "ofile="])
    except getopt.GetoptError as err:
        print(err)
        usage()
        sys.exit(2)
    for opt, arg in opts:
        if opt == '-h':
            usage()
            sys.exit()
        elif opt in ("-a", "--all"):
            fileCount = len([f for f in os.listdir()
                             if f.endswith('.ritdb') and os.path.isfile(os.path.join(f))])
            print("Found " + str(fileCount) +
                  ' RITdb files in ' + os.getcwd() + "...")

            for file in os.listdir():
                if file.endswith('.ritdb'):

                    toolbar_width = len(file)
                    for fi in file:
                        # setup toolbar
                        sys.stdout.write("[%s]" % (" " * toolbar_width))
                        sys.stdout.flush()
                        # return to start of line, after '['
                        sys.stdout.write("\b" * (toolbar_width+1))

                        for i in range(toolbar_width):
                            subprocess.run('RITdb-verifies2csv.pyw -i ' + file +
                                           '& RITdb-verifies2xlsx.pyw -i' + file, shell=True)
                            # update the bar
                            sys.stdout.write("-")
                            sys.stdout.flush()

                        sys.stdout.write("]\n")  # this ends the progress bar
                    sys.exit()
                else:
                    continue

        elif opt in ("-i", "--ifile"):
            dbFileName = arg.strip()

    if not bool(csvFileName):   # csvFileName is empty
        baseName, _ = os.path.splitext(dbFileName)
        csvFileName = (baseName + '.csv')

 #   RITdbVerify2csv(dbFileName, csvFileName)
 #   RITdbVerify2xlsx(dbFileName, xlsxFileName)

    subprocess.run('RITdb-verifies2csv.pyw -i ' + dbFileName +
                   '& RITdb-verifies2xlsx.pyw -i' + dbFileName, shell=True)
