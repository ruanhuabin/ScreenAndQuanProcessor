# encoding: utf-8
from Logger import MyLogger
import logging
from util import printDict
import math


def nonBlankLines(f):
    for l in f:
        line = l.rstrip()
        if line and line[0] != '#':
            yield line




def readTextFile(fileName, logger): 
    content = []  
    with open(fileName) as f:
        for line in nonBlankLines(f):
            content.append(line)
    
    return content

def loadTextFile(fileName, logger = 0):
    content = readTextFile(fileName, logger)
    headerLine = []
    if(len(content) > 0):
        headerLine = content[0].split('\t')
    
    headers = []
    dataDict = {}
    for header in headerLine:
        headers.append(header)
        dataDict[header] = {}
    
    
    count = 0    
    for item in content:
        item = item.split('\t')
        count += 1
        itemLen = len(item)
        #print count, ":", itemLen, ":" , item
        logger.debug("%d:%d:%s" % (count, itemLen, str(item)))
    #transpose the list of lists
    content2 = []
    for item in content:
        item = item.split('\t')
        content2.append(item)
        
    
    #Transpose the content matrix
    contentT = [[x[i] for x in content2] for i in range(len(content2[0]))]
    count = 0
    logger.debug("After transpose:")
    for item in contentT:        
        count += 1
        itemLen = len(item)
        #print count, ":", itemLen, ":" , item
        logger.debug("%d:%d:%s" % (count, itemLen, str(item)))
    
    for item in contentT:
        header = item[0]
        item.pop(0)
        dataDict[header] = item
        
    
    #print "Value in dataDict: ", dataDict
    #printDict(dataDict)
    #print dataDict
        
    return dataDict

def extractLipidIon(dataDict, logger):
    lipidIon = dataDict["LipidIon"]
    Rt = dataDict["Rt"]
    TopRT = dataDict["TopRT"]
    Formula = dataDict["Formula"]
    Grade = dataDict["Grade"]
    
    ms2Window = 0.2
    
    lipidFinal = []
    rtFinal = []
    topRTFinal = []
    formulaFinal = []
    gradeFinal = []
    
    origLen = len(lipidIon)
    for i in range(origLen):
        rtValue = Rt[i]
        topRTValue = TopRT[i]
        lipidName = lipidIon[i]
        formulaValue = Formula[i]
        gradeValue = Grade[i]
        
        
        diff = math.fabs(float(rtValue) - float(topRTValue))
        if( diff <= ms2Window):
            
            lipidFinal.append(lipidName)
            rtFinal.append(rtValue)
            topRTFinal.append(topRTValue)
            formulaFinal.append(formulaValue)
            gradeFinal.append(gradeValue)
            
            
            
            logger.info("Append to lipid Final: %s : %s: %s : %f" %(lipidName, rtValue, topRTValue, diff))
        else:
            
            logger.info("NOT to Append to lipid Final: %s : %s: %s : %f" %(lipidName, rtValue, topRTValue, diff))
            
    
    logger.info("LipidFinal = " + str(lipidFinal))
    logger.info("formulaFinal = " + str(formulaFinal))
    logger.info("topRTFinal = " + str(topRTFinal))
    logger.info("gradeFinal = " + str(gradeFinal))
    logger.info("Total lipidIon: %d, final added: %d" % (len(lipidIon), len(lipidFinal)))
    
    
            
        
        
    
        
    
    

if __name__ == '__main__':

    logger = MyLogger("Lipid-Logger", logging.INFO).getLogger()
    dataDict = loadTextFile("sample.txt", logger)
    extractLipidIon(dataDict, logger)
    pass