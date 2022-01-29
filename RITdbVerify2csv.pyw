#!/usr/bin/python3
#
'''
Convert Cassini Verify results from RITdb to CSV format

Usage: RITdbVerify2csv.py -i <input RITdb file> [-o <output CSV file>]
'''

import os
import sys
import getopt
import csv
import sqlite3 as sql
from sqlite3 import Error
import ctypes

# hides python terminal window when running from a Guru Agent on Windows
ctypes.windll.user32.ShowWindow(ctypes.windll.kernel32.GetConsoleWindow(), 0)

verify_reslutInfoQuery = '''
SELECT  n1.value AS 'Result Number',  n2.value AS 'Result Name',  n4.value AS 'Units',  n8.value AS 'U Limit',  n9.value AS 'L Limit',  n3.value AS 'RESULT_ID'
FROM ritdb1 n0
 JOIN ritdb1 n1 ON n0.entityID=n1.entityID AND n1.indexID='0' AND n1.name='RESULT_NUMBER'
 JOIN ritdb1 n2 ON n0.entityID=n2.entityID AND n2.indexID='0' AND n2.name='RESULT_NAME'
 JOIN ritdb1 n3 ON n0.entityID=n3.entityID AND n3.indexID='0' AND n3.name='RESULT_ID'
 JOIN ritdb1 n4 ON n0.entityID=n4.entityID AND n4.indexID='0' AND n4.name='RESULT_UNITS'
 JOIN ritdb1 n6 ON n6.indexID='0' AND n6.name='ENTITY_TYPE' AND n6.value='RESULT_LIMIT_SET'
 JOIN ritdb1 n7 ON n6.entityID=n7.entityID AND 'SystemCheck'=n7.value AND n7.indexID='0' AND n7.name='LIMIT_SET_NAME'
 LEFT JOIN ritdb1 n8 ON n0.entityID=n8.indexID AND n6.entityID=n8.entityID AND n8.name='UL'
 LEFT JOIN ritdb1 n9 ON n0.entityID=n9.indexID AND n6.entityID=n9.entityID AND n9.name='LL'
WHERE
 n0.name='ENTITY_TYPE' AND n0.value='RESULT_INFO'
ORDER BY n3.value ASC
'''

resultsQuery = '''
SELECT  n0.value * n3.value AS 'Result',  n0.value2,  n1.value,  n2.value
FROM ritdb1 n0
 JOIN ritdb1 n1 ON n0.entityID=n1.entityID AND n1.name='PART_RESULT_EVENT_ORDER'
 JOIN ritdb1 n2 ON n0.indexID=n2.entityID AND n2.name='RESULT_ORDER'
 LEFT JOIN ritdb1 n3 ON n2.entityID=n3.entityID AND n3.name='RESULT_SCALE'
WHERE
 n0.name='R'
ORDER BY n2.value, n1.value ASC
'''


def ritdb2csv(dbFileName, csvFileName):
    try:
        # Connect to database
        conn = sql.connect(dbFileName)
        cur = conn.cursor()
        # Query the top header
        nDevice = 1
        failedCount = 0
        # query the left table
        cur.execute(verify_reslutInfoQuery)
        leftTable = cur.fetchall()
        leftLabel = []
        for i in range(0, 5):
            leftLabel.append(cur.description[i][0])
        leftLabel.append(os.path.basename(baseName))
        # query result data, sort by PART_RESULT_EVENT_ORDER then RESULT_ORDER
        cur.execute(resultsQuery)

        # Open csv writer
        with open(csvFileName, "w", newline='') as csv_file:
            csv_writer = csv.writer(
                csv_file, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)

            # Write the top table to csv file
            csv_writer.writerow(leftLabel)

            # Write one row at a time
            data = cur.fetchone()     # get the first row of resultQuery data
            for ltRow in leftTable:
                dataLists = ['' for i in range(nDevice)]
                temp = []
                for i in range(0, 5):
                    temp.append(ltRow[i])

                # leftTable.testEntityID = data.testEntityID
                while data != None and ltRow[5] == data[3]:
                    dataLists[data[2] - 1] = data[0]  # device ID is 1 index
                    # Test UL and LL Limits and Count
                    if ltRow[3] and ltRow[3] < data[0]:
                        failedCount = failedCount + 1
                    elif ltRow[4] and ltRow[4] > data[0]:
                        failedCount = failedCount + 1
                    data = cur.fetchone()
                temp.extend(dataLists)
                csv_writer.writerow(temp)
                dataLists.clear()
                temp.clear()
            # Set Verify PASS/FAIL Result
            if failedCount > 0:
                csv_writer.writerow(['', 'FAIL', 'Count', '', '', failedCount])
            else:
                csv_writer.writerow(['', 'PASS', '', '', '', ''])

            csv_file.close()
            print("SUCCESS converting to CSV.", csvFileName)

    except Error as error:
        print("FAILED:", error, dbFileName)

    except Exception as e:
        print("ERROR:", e)

    # Close database connection
    finally:
        conn.close()
        sys.exit()
### end of ritdb2csv ###


if __name__ == '__main__':

    dbFileName = ''
    csvFileName = ''
    pivot = False

    argv = sys.argv[1:]
    usage = (
        "Usage: %s -i <input RITdb file> [-o <output CSV file>]" % (sys.argv[0]))

    opts, args = getopt.getopt(argv, "hpi:o:", ["ifile=", "ofile="])

    try:
        opts, args = getopt.getopt(
            argv, "hpi:o:", ["pivot", "ifile=", "ofile="])
    except getopt.GetoptError:
        print(usage)
        sys.exit(2)
    for opt, arg in opts:
        if opt == '-h':
            print(usage)
            sys.exit()
        elif opt in ("-i", "--ifile"):
            dbFileName = arg.strip()
        elif opt in ("-o", "--ofile"):
            csvFileName = arg.strip()

    if not bool(dbFileName):    # dbFileName is empty
        print(usage)
        sys.exit(2)

    if not bool(csvFileName):   # csvFileName is empty
        baseName, _ = os.path.splitext(dbFileName)
        csvFileName = (baseName + '.csv')

    ritdb2csv(dbFileName, csvFileName)
