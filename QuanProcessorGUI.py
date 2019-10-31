# encoding: utf-8
'''
Created on 2016��9��14��

@author: ruanhuabin
'''
from Tkinter import *
from tkFileDialog import askopenfilename
from tkFileDialog import askdirectory
from _functools import partial

#from Screen import init

from Screen import init, printDict, checkFileValid, extractCompoundNameColumn,\
    extractMZExpectColumn, genRTColumn, genLibraryScoreColumn, genIPColumn,\
    genLSColumn, genRTRangeColumn, genMZDeltaColumn, genMeasuredAreaColumn,\
    genRTMeasuredColumn
from constant import *
from util import extractMZExpectData, writeWordBook
import tkMessageBox
from Quan import loadData, extractColumnToFile
#import thread
import time
from threading import Thread
from Logger import MyLogger
from util import getCurrTime
import os

class QuanFrame:
    def __init__(self, master):
        frame = Frame(master,width=400,height=600)
        frame.pack()   
        
        self.inputFilePath =StringVar()
        self.labelFolder = Label(frame,textvariable=self.inputFilePath).grid(row=0,columnspan=2)
        
        self.log_file_btn = Button(frame, text="Select Quan File", command=self.selectInputFile).grid(row=1, columnspan=2)
        
        self.outputFolderPath = StringVar()
        self.labelImageFile = Label(frame,textvariable = self.outputFolderPath).grid(row = 2, columnspan = 2)
        self.output_folder_btn = Button(frame, text="Output Folder", command=self.selectOutputFolder,width=40).grid(row=3, columnspan=2)
        
        self.blankText = StringVar()
        self.blankLabel = Label(frame,textvariable = self.blankText).grid(row = 4, columnspan = 2)
        
        self.output_filename_label = Label(frame, text="Output File Name:          ")#.grid(row=5, columnspan=2)
        self.output_filename_label.grid(row=5, columnspan=2)
       
        self.output_filename_entry = Entry(frame, width=30, bd = 4)
        self.output_filename_entry.grid(row=6, column=0, columnspan=2)
        
        
        f = frame
        xscrollbar = Scrollbar(f, orient=HORIZONTAL)
        xscrollbar.grid(row=12, column=0, sticky=N+S+E+W)
 
        yscrollbar = Scrollbar(f)
        yscrollbar.grid(row=11, column=1, sticky=N+S+E+W)
 
        self.text_field = Text(f, wrap=NONE,
                    xscrollcommand=xscrollbar.set,
                    yscrollcommand=yscrollbar.set)
        self.text_field.grid(row=11, column=0)
 
        xscrollbar.config(command=self.text_field.xview)
        yscrollbar.config(command=self.text_field.yview)
      

        
        
        self.blankText = StringVar()
        self.blankLabel = Label(frame,textvariable = self.blankText).grid(row = 7, columnspan = 2)
        
        self.start_run_btn = Button(frame, text="Start Processing", command=partial(self.startProcessing, self.text_field), width=40)
        self.start_run_btn.grid(row=8, columnspan=2)
        
        self.blankText = StringVar()
        self.blankLabel = Label(frame,textvariable = self.blankText).grid(row = 9, columnspan = 2)
        self.counter = 0
        
        self.inputDataBook = {}
        self.outputDataBook = {}
        self.inputFilename = ""
        
        self.logger = MyLogger("Quan-Logger").getLogger()
        
        self.isFinishProcessed = False
        

  

    def selectInputFile(self):
        filename = askopenfilename()
        self.inputFilePath.set(filename)
        self.inputFilename = filename
        
        
        dirName = os.path.dirname(filename)
        baseName = os.path.basename(filename)
        
        baseNamePrefix = baseName.split(".")[0]
        baseNamePrefix = baseNamePrefix + "-filter"
        
        outputFilename = baseNamePrefix + ".xlsx"
        
#         self.logger.info("dir name = " + dirName)
#         self.logger.info("base name = " + baseName)
        self.outputFolderPath.set(dirName)
        self.output_filename_entry.delete(0, END)
        self.output_filename_entry.insert(0, outputFilename)
        
      
        
       

    def selectOutputFolder(self):
        imageFolder = askdirectory()
        self.outputFolderPath.set(imageFolder)
        
    def run_thread(self, text_field, outputFilename):
        
        text_field.insert(INSERT, "[%s]: Start to Process file: %s\n" % (getCurrTime(), self.inputFilePath.get()))        
        self.counter = self.counter + 1
        
         
        rowTitleMust = [quanExpectRTTitle, quanAreaTitle, quanCompoundTitle, quanActualRTTitle, quanMZExpectedTitle]
        text_field.insert(INSERT, "[%s]: Start Loading Data: %s\n" %(getCurrTime(), self.inputFilePath.get()))
        self.counter += 1
        self.inputDataBook = loadData(self.inputFilename, self.logger, self.text_field)
        text_field.insert(INSERT, "[%s]: Finish Loading Data: %s\n" %(getCurrTime(), self.inputFilePath.get()))
        self.counter += 1
        
        #print "self.inputDataBook = ", self.inputDataBook        
        [missingNum, missingRowTitleDict] = checkFileValid(self.inputDataBook,  rowTitleMust )
        text_field.insert(INSERT, "[%s]: Check file validation complete\n" %(getCurrTime()))
        self.counter += 1
        if(missingNum > 0):
            self.logger.error("Some worksheet missing some column(s):")
            #printDict(missingRowTitleDict)
            self.logger.error(str(missingRowTitleDict))
            text_field.insert(INSERT, "[%s]: Error: Missing Columns: %s\n" %(getCurrTime(), str(missingRowTitleDict)))
            
            text_field.insert(INSERT, "[%s]: Error: Some worksheet missing some column(s), see missing column above\n" %(getCurrTime()))
            
            self.start_run_btn.configure(state=NORMAL)
            raise ValueError("Error: Some worksheet missing some column(s), see missing column above")
        
        
        text_field.insert(INSERT, "[%s]: Start to extracting data\n" %(getCurrTime())) 
        extractColumnToFile(self.inputDataBook, outputFilename, self.logger)
        text_field.insert(INSERT, "[%s]: Finish extracting data\n" %(getCurrTime()))
        text_field.insert(INSERT, "[%s]: Done Successfully, output file name is: %s\n" %(getCurrTime(), outputFilename))
        
        text_field.see("end")
        time.sleep(1)
        
        
        self.start_run_btn.configure(state=NORMAL)
        
    
    def startProcessing(self, text_field):
        if(self.inputFilePath.get() == ""):
            tkMessageBox.showerror("Error", "Please select a screen file to be processed")
            return
        if(self.outputFolderPath.get() == ""):
            self.outputFolderPath.set("./")

        if(self.output_filename_entry.get() == ""):           
            tkMessageBox.showerror("Error", "Please enter a filename for the output file" )
            return
            
        outputFolderPath = self.outputFolderPath.get()
        if(outputFolderPath[-1] != '/'):
            outputFolderPath = outputFolderPath + "/"
            
        outputFilename = outputFolderPath + self.output_filename_entry.get()
        
        filenameExtension = outputFilename.split(".")[-1]
        if(filenameExtension != "xlsx"):
            outputFilename = outputFilename + ".xlsx"
        
#         import thread
#         thread.start_new_thread(self.run_thread, (text_field,outputFilename))
  
        self.start_run_btn.configure(state=DISABLED)
        newThread = Thread(target = self.run_thread, args=(text_field, outputFilename))
        newThread.start()
          
          
        #newThread.join()
        
          
        #self.start_run_btn.configure(state=NORMAL)
               
       



def run_screen_processor():  
    mainFrame = Tk()
    mainFrame.title("Quan Data Processor")
    mainFrame.geometry("600x800")
    mainFrame.resizable(False, False)
    app = QuanFrame(mainFrame)
    mainFrame.mainloop()


if __name__ == '__main__':
    run_screen_processor()
#     mainThread = Thread(target = run_screen_processor())
#     mainThread.start()
    pass