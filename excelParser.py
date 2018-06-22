#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Thu Jun 21 14:39:11 2018

@author: m006703
"""

from openpyxl import load_workbook
import os

basedir = '/Users/m006703/NIPS/RTF/'
fileList = os.listdir(basedir)

sampleNameList = []
initialQuantList = []
finalQuantList = []
averageSizeList = []
runNameList = []

for file in fileList:
    if file[0:4] == 'NIPS':
        filePath = basedir + file
        runName = file.split('_')[0]
        print("Processing file: " + filePath)
    
        wb = load_workbook(filePath, data_only = True, read_only=True)
        
        quantSheet = wb.get_sheet_by_name('DNA Quantitation')
    
        for x in range (4, 999):
            try:
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
                    
            except TypeError:
                continue
                
resultFile = open('/Users/m006703/NIPS/NIPSQuantResult.txt', 'w')
resultFile.write("RunName\tSampleName\tInitialQuant\tFinalQuant\tAverageSize\n")

for index in range (0, len(sampleNameList)):
    resultFile.write(runNameList[index] + "\t" + sampleNameList[index] + "\t" + initialQuantList[index] + "\t" + \
                     finalQuantList[index] + "\t" + str(averageSizeList[index]) + "\n")
    
resultFile.close()

print("Script is done running")