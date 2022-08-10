import os
import shutil
import sys
import logging
import pytesseract
from datetime import date
from pdf2image import convert_from_path

# To Launch: 
# 1. shift+ctrl+P to open comamnd pallete
# 2. click select interpreter
# 2. select interpreter C:\py\environments\autoPrint\Scripts\python.exe







poppler_path = os.path.join(os.path.dirname(__file__), 'poppler-0.68.0_x86/bin')
custom_config = r'--oem 3 --psm 6'

copyConvertTitle = '----------- COPY/CONVERT ----------'
errorMsgType = 'ERROR'



def run(): 
    global errorLog 
    errorLog = None

    global runningLog 
    runningLog = None

    global dateTodaySource 
    dateTodaySource = date.today().strftime("%m%d%Y")

    global dateTodayDest
    dateTodayDest = date.today().strftime("%Y%m%d")

    

    # check command line args 
    if len(sys.argv) > 2:
        log('Too many arguments specified, only accepts target date in form mmddyyyy - Shutting down', type = errorMsgType)
        print('len(sys.argv) two or: ',len(sys.argv) )
        
        sys.exit(0)
    elif len(sys.argv) == 2:
        dateTodaySource = sys.argv[1]
        dateTodayDest = dateTodaySource[4:]+dateTodaySource[:4]
        log('Running for user specified date: '+ dateTodaySource)
    else:
        log('Running for todays date: '+ dateTodaySource )

    global sourceDir
    sourceDir = r'\\mryflash\TempDecStore\iPub_Support'

    global destDir
    destDir = r'\\mryflash\renocsc$\Notices\A_Notices_for_CS'
    destDir = os.path.join(destDir,dateTodayDest)
    
    #check destination directory 
    if not os.path.isdir(destDir):
        log('Destination Directory: '+destDir+' does not exist - Shutting down', type = errorMsgType)
        sys.exit(0)

    log('Program Start')
    pclToPdf()
    filterRename() 


def deleteFile(filePath):
    if os.path.exists(filePath):
        os.remove(filePath)

def processDestDir(top):
    
    for fileName in os.listdir(top):
        filePath = os.path.join(top, fileName)
        if not os.path.isfile(filePath):
            continue
        if not fileName.endswith('.pdf'):
            continue
        yield (top, fileName)

def filterRename():
    global destDir

    log('\nFiltering/Renaming: '+destDir+' ---------------------------------------')

    for (folder, name) in processDestDir(destDir):

        try:
            filePath = os.path.join(folder,name)
            pages = convert_from_path(filePath, dpi=350, poppler_path=poppler_path)
            firstPageText = pytesseract.image_to_string(pages[0], config=custom_config)
            

            if(firstPageText.find("Company Copy") != -1):
                log('Deleted - company copy: '+name)
                deleteFile(filePath)
            elif not "Affidavit" in name:
                binderNbrLeadingIndex = firstPageText.find("Binder Number: ")
                if binderNbrLeadingIndex != -1:
                    binderNbrLeadingIndex = binderNbrLeadingIndex + 15
                    binderNbrTrailingIndex = firstPageText.find("\n",binderNbrLeadingIndex, len(firstPageText))

                    binderNbr = firstPageText[binderNbrLeadingIndex:binderNbrTrailingIndex]

                    if(firstPageText.find("Agent Copy") != -1):
                        newName = os.path.join(destDir, binderNbr+"_A.pdf") 
                    else:
                        newName = os.path.join(destDir, binderNbr+"_I.pdf") 
                    

                policyNbrLeadingIndex = firstPageText.find("Policy Number: ") 
                if policyNbrLeadingIndex != -1:
                    policyNbrLeadingIndex = policyNbrLeadingIndex + 15
                    policyNbrTrailingIndex = firstPageText.find("\n",policyNbrLeadingIndex, len(firstPageText))
                    
                    plcyNbr = firstPageText[policyNbrLeadingIndex:policyNbrTrailingIndex]

                    if(firstPageText.find("Agent Copy") != -1):
                        newName = os.path.join(destDir, plcyNbr+"_A.pdf") 
                    else:
                        newName =  os.path.join(destDir, plcyNbr+"_I.pdf") 

                os.rename(filePath, newName)
                log('Rename Successful: '+name+"   ->   "+ os.path.basename(newName))
        except Exception as ex:
            log('Rename Failed: '+name+'\n\tMessage:\n\t'+str(ex), type = errorMsgType)
        
# Finds today's file paths and returns them as a generator
def filePaths():
    global sourceDir
    global dateTodaySource

    try:
        
        dirList = os.listdir(sourceDir)
        

        # find directories marked with today's date that contain a Notices folder
        targetDirectory = [dir for dir in dirList if dateTodaySource == dir[:8] and  "Notices" in os.listdir(os.path.join(sourceDir,dir))]

        #this shouldn't happen, throw error
        if len(targetDirectory) == 0:
            raise NameError('Source directory: ' + os.path.join(sourceDir,dateTodaySource) + ' does not exist - Shutting down')
        elif len(targetDirectory) > 1:
            raise NameError('Multiple Source Directories found: '+', '.join(targetDirectory)+' - Shutting down')

        batFile = os.path.join(sourceDir, targetDirectory[0], "Notices", "PrintNotices"+targetDirectory[0]+".bat")

        for line in open(batFile, 'r'):
            words = line.split(' ')
            filePath = words[-1]
            if ("0_CIG_0" in filePath or "Affidavit" in filePath) and ".pdf" in filePath: 
                yield(filePath)

    except NameError as ex:
        log(ex, type = errorMsgType)
        sys.exit(0)
    except OSError as ex:
        log('Could not open source file: ' + batFile + ' - Shutting down', type = errorMsgType)
        sys.exit(0)

# copies pdf versions from source to dest directories
def pclToPdf():              
    global destDir

    log(copyConvertTitle)
    for filePath in filePaths():
        
        destPathTemp = os.path.join(destDir, os.path.basename(filePath))
        conversionMessage = filePath.strip()+"  ->  "+destPathTemp.strip()

        try:
            shutil.copyfile(filePath.strip(), destPathTemp.strip())
            log('Copy Successful: '+conversionMessage)
        except Exception as ex:
            log('Copy failed: '+conversionMessage+'\n\tMessage:\n\t'+str(ex), type = errorMsgType)

def setup_logger(name, log_file):
    """To setup as many loggers as you want"""
    formatter = logging.Formatter(fmt='%(asctime)s - %(levelname)s - %(message)s', datefmt='%d-%b-%y %H:%M:%S')
    handler = logging.FileHandler(log_file)        
    handler.setFormatter(formatter)

    logger = logging.getLogger(name)
    logger.setLevel(logging.INFO)
    logger.addHandler(handler)

    return logger    

def log(msg, type = None):
    global errorLog
    global runningLog
    # INITIALIZE LOGGERS -----------------------------------------------
    logDir = os.path.join(os.getcwd(),'logs')
    if not os.path.exists(logDir):
        os.mkdir(logDir)

    if type == errorMsgType:
        errorLogInit()
        errorLog.error(msg)
        runningLog.error(msg)
        print('ERROR - ',msg)
    else:
        runningLogInit()
        runningLog.info(msg)
        print(msg)

def errorLogInit():
    global errorLog
    global dateTodayDest
    if errorLog == None:
        # Create today's error log
        errorLogDir = os.path.join(os.getcwd(),'logs','errorLogs')
        if not os.path.exists(errorLogDir):
            os.mkdir(errorLogDir)
        errorLogPath = os.path.join(errorLogDir,'error_' + dateTodayDest + '.log' )
        errorLog = setup_logger('errorLog',errorLogPath)
    runningLogInit()

def runningLogInit():
    global runningLog
    global dateTodayDest
    if runningLog == None:
        # Create today's running log
        runningLogDir = os.path.join(os.getcwd(),'logs','runningLogs')
        if not os.path.exists(runningLogDir):
            os.mkdir(runningLogDir)
        runningLogPath = os.path.join(runningLogDir,'run_' + dateTodayDest + '.log')
        runningLog = setup_logger('runningLog',runningLogPath)


run()









