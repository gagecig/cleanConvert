import os
import logging
import pytesseract
from datetime import date
from pdf2image import convert_from_path

# To Launch: 
# 1. shift+ctrl+P to open comamnd pallete
# 2. click select interpreter
# 2. select interpreter C:\py\environments\autoPrint\Scripts\python.exe


sourceDir = r'\\mryflash\TempDecStore\iPub_Support'

poppler_path = os.path.join(os.path.dirname(__file__), 'poppler-0.68.0_x86/bin')
custom_config = r'--oem 3 --psm 6'




def processDestDir(top):
    if not os.path.isdir(top):
        exit(1)
    for fileName in os.listdir(top):
        filePath = os.path.join(top, fileName)
        if not os.path.isfile(filePath):
            continue
        if not fileName.endswith('.pdf'):
            continue
        yield (top, fileName)
            
def convert_pdf(src_path, dst_path):
    cmd = "gpcl6win64.exe -dNOPAUSE -sDEVICE=pdfwrite -sOutputFile=%s %s" % (dst_path, src_path)
    os.system(cmd)

# copies pdf versions from source to dest directories
def pclToPdf(destDir):              
    
    i = 0
    for filePath in filePaths():
        
        destPathTemp = os.path.join(destDir, os.path.basename(filePath).replace(".pcl",".pdf"))
        i = i + 1
        print(str(i)+": "+filePath.strip()+" -> "+destPathTemp.strip())
        convert_pdf(filePath.strip(), destPathTemp.strip())
        print("done.")

def run(dateTodayDest): 
    destDir = r'\\mryflash\renocsc$\Notices\A_Notices_for_CS'
    destDir = os.path.join(destDir,dateTodayDest)
    pclToPdf(destDir)
    filterRename(destDir)

def deleteFile(filePath):
    if os.path.exists(filePath):
        os.remove(filePath)

def filterRename(destDir):

    for (folder, name) in processDestDir(destDir):

        filePath = os.path.join(folder,name)
        pages = convert_from_path(filePath, dpi=350, poppler_path=poppler_path)
        firstPageText = pytesseract.image_to_string(pages[0], config=custom_config)
        print("\n")
        print("Checking   ", name)

        if(firstPageText.find("Company Copy") != -1):
            print("Company Copy Deleted")
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
                print("Renamed To ", os.path.basename(newName))
                os.rename(filePath, newName )

            policyNbrLeadingIndex = firstPageText.find("Policy Number: ") 
            if policyNbrLeadingIndex != -1:
                policyNbrLeadingIndex = policyNbrLeadingIndex + 15
                policyNbrTrailingIndex = firstPageText.find("\n",policyNbrLeadingIndex, len(firstPageText))
                
                plcyNbr = firstPageText[policyNbrLeadingIndex:policyNbrTrailingIndex]

                if(firstPageText.find("Agent Copy") != -1):
                    newName = os.path.join(destDir, plcyNbr+"_A.pdf") 
                else:
                   newName =  os.path.join(destDir, plcyNbr+"_I.pdf") 
                print("Renamed To ", os.path.basename(newName))
                os.rename(filePath, newName)
        
# Finds today's file paths and returns them as a generator
def filePaths():

    dirList = os.listdir(sourceDir)
    dateTodaySource = date.today().strftime("%m%d%Y")

    # find directories marked with today's date that contain a Notices folder
    targetDirectory = [dir for dir in dirList if dateTodaySource == dir[:8] and  "Notices" in os.listdir(os.path.join(sourceDir,dir))]

    #this shouldn't happen, throw error
    if len(targetDirectory) != 1:
        #throw error
        something = "something"


    batFile = os.path.join(sourceDir, targetDirectory[0], "Notices", "PrintNotices"+targetDirectory[0]+".bat")

    for line in open(batFile, 'r'):
        words = line.split(' ')
        filePath = words[-1]
        if "0_CIG_0" in filePath or "Affidavit" in filePath: 
            yield(filePath)

def setup_logger(name, log_file):
    """To setup as many loggers as you want"""
    formatter = logging.Formatter(fmt='%(asctime)s - %(levelname)s - %(message)s', datefmt='%d-%b-%y %H:%M:%S')
    handler = logging.FileHandler(log_file)        
    handler.setFormatter(formatter)

    logger = logging.getLogger(name)
    logger.setLevel(logging.INFO)
    logger.addHandler(handler)

    return logger    

#fetch today's date for destination folder
dateTodayDest = date.today().strftime("%Y%m%d")

logDir = os.path.join(os.getcwd(),'logs')
if not os.path.exists(logDir):
    os.mkdir(logDir)
    
# Create today's error log
errorLogDir = os.path.join(os.getcwd(),'logs','errorLogs')
if not os.path.exists(errorLogDir):
    os.mkdir(errorLogDir)
errorLogPath = os.path.join(errorLogDir,'error_' + dateTodayDest + '.log' )
errorLog = setup_logger('errorLog',errorLogPath)

# Create today's running log
runningLogDir = os.path.join(os.getcwd(),'logs','runningLogs')
if not os.path.exists(runningLogDir):
    os.mkdir(runningLogDir)
runningLogPath = os.path.join(runningLogDir,'run_' + dateTodayDest + '.log')
runningLog = setup_logger('runningLog',runningLogPath)

# run(dateTodayDest)









