# encoding: utf-8
'''
Created on 2016��9��15��

@author: ruanhuabin
'''
from Screen import *
from constant import *
from util import extractPartData

def run_screen_process(inputFilename):
    rowTitleMust = [compoundNameTitle, mzExpectedTitle, libraryScoreTitle, measuredAreaTitle, ipTitle, lsTitle, rtMeasuredTitle, mzDeltaTitle]


    [screenDataBook, newScreenBook] = init(src=inputFilename)
    [missingNum, missingRowTitleDict] = checkFileValid(screenDataBook, rowTitleMust)
    if(missingNum > 0):
        print "Error: Some worksheet missing some column(s):"
        printDict(missingRowTitleDict)    
        raise ValueError("Error: Some worksheet missing some column(s), see missing column above")
    compoundNames = extractCompoundNameColumn(screenDataBook, newScreenBook)
      
    # 
    extractMZExpectColumn(screenDataBook, compoundNames, newScreenBook)  
    extractFormulaColumn(screenDataBook, compoundNames, newScreenBook)     
    genRTColumn(screenDataBook, compoundNames, newScreenBook)   
    genLibraryScoreColumn(screenDataBook, compoundNames, newScreenBook)
    genIPColumn(screenDataBook, compoundNames, newScreenBook)
    genLSColumn(screenDataBook, compoundNames, newScreenBook)
    genRTRangeColumn(screenDataBook, compoundNames, newScreenBook)
    genMZDeltaColumn(screenDataBook, compoundNames, newScreenBook)
      
    genMeasuredAreaColumn(screenDataBook, compoundNames, newScreenBook)
    genRTMeasuredColumn(screenDataBook, compoundNames, newScreenBook)
     
    writeWordBook(newScreenBook, "screen-filter.xlsx")
     
     
     
    excelProcess = popen4("start excel D:\workspace-excelprocess-final\ExcelProcessor/screen-filter.xlsx")
    print("Enter to finish")
    import sys
    line = sys.stdin.readline()
    #sleep(100)
    Popen("taskkill /F /im EXCEL.EXE",shell=True)

if __name__ == '__main__':
    print "This is main function"
    #inputDataBook = loadData("screen.xlsx")
    #extractPartData(wordBook=inputDataBook, rows = 50, output="middle-screen.xlsx")
    
    run_screen_process("screen.xlsx")
    pass