#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Thu Jun 21 14:39:11 2018

@author: m006703
"""

from openpyxl import load_workbook
import os
import argparse

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument(
        '-i', '--inputDir', dest='basedir', required=True,
        help="Path to input directory of NIPS RTFs"
    )
    parser.add_argument(
        '-o', '--outpath', dest='outpath', required=True,
        help="Path to output file"
    )

    args = parser.parse_args()

    basedir = os.path.abspath(args.basedir)
    outpath = os.path.abspath(args.outpath)

    # Add / at the end if it is not included in the paths
    if basedir.endswith("/"):
        basedir = basedir
    else:
        basedir = basedir + "/"

    if outpath.endswith("/"):
        outpath = outpath
    else:
        outpath = outpath + "/"

    # Add all of the RTFs in a list
    fileList = os.listdir(basedir)

    sampleNameList = []
    initialQuantList = []
    finalQuantList = []
    averageSizeList = []
    runNameList = []

    for file in fileList:
        #Make sure it's a NIPS RTF
        if file[0:4] == 'NIPS':
            filePath = basedir + file
            #This script will work but the result file will be messed up if it's not in the form of NIPS###_RTF.xlsm
            runName = file.split('_')[0]
            print("Processing file: " + filePath)
            #Make sure the read_only=True is there, otherwise it takes forever!
            wb = load_workbook(filePath, data_only = True, read_only=True)

            quantSheet = wb['DNA Quantitation']

            for x in range (4, 999):
                try:
                    #Sometimes the NTC is NPP-NTC or NTC...
                    if 'NTC' in (quantSheet.cell(row = x, column = 2).value):
                        break
                    else:
                        sampleName = quantSheet.cell(row = x, column = 2).value
                        sampleNameList.append(sampleName)
                        try:
                            initialQuant = ("%.2f" % quantSheet.cell(row = x, column = 5).value)
                        except TypeError:
                            initialQuant = 'NA'
                        initialQuantList.append(initialQuant)
                        try:
                            finalQuant = ("%.2f" % quantSheet.cell(row = x, column = 7).value)
                        except TypeError:
                            finalQuant = 'NA'
                        finalQuantList.append(finalQuant)
                        # Gather the average size for every 8 samples
                        if ((x - 4) % 8) == 0:
                            try:
                                averageSize = quantSheet.cell(row = x, column = 9).value
                            except TypeError:
                                averageSize = 'NA'
                        averageSizeList.append(averageSize)
                        runNameList.append(runName)
                #Skip if there are any blank lines between samples and NTC
                except TypeError:
                    continue

    resultFile = open(outpath + "NIPSQuantResult.txt", 'w')
    resultFile.write("RunName\tSampleName\tInitialQuant\tFinalQuant\tAverageSize\n")

    for index in range (0, len(sampleNameList)):
        resultFile.write(runNameList[index] + "\t" + sampleNameList[index] + "\t" + initialQuantList[index] + "\t" + \
                         finalQuantList[index] + "\t" + str(averageSizeList[index]) + "\n")

    resultFile.close()

    print("Script is done running")
    print("Results are saved at " + outpath + "NipsQuantResult.txt")


if __name__ == "__main__":
    main()
