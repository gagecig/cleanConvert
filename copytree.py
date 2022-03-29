import os
import pytesseract
from datetime import date
from pdf2image import convert_from_path

# To Launch: 
# 1. shift+ctrl+P to open comamnd pallete
# 2. click select interpreter
# 2. select interpreter C:\py\environments\autoPrint\Scripts\python.exe


sourcePath = r'\\mryflash\TempDecStore\iPub_Support'
destPath = r'\\mryflash\renocsc$\Notices\A_Notices_for_CS\gshawTemp'

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
def pclToPdf():              
    
    i = 0
    for filePath in filePaths():
        
        destPathTemp = os.path.join(destPath, os.path.basename(filePath).replace(".pcl",".pdf"))
        i = i + 1
        print(str(i)+": "+filePath.strip()+" -> "+destPathTemp.strip())
        convert_pdf(filePath.strip(), destPathTemp.strip())
        print("done.")

def run(): 
    pclToPdf()
    filterRename()
    

def deleteFile(filePath):
    if os.path.exists(filePath):
        os.remove(filePath)


def filterRename():

    for (folder, name) in processDestDir(destPath):

        filePath = os.path.join(folder,name)
        pages = convert_from_path(filePath, dpi=350, poppler_path=poppler_path)
        firstPageText = pytesseract.image_to_string(pages[0], config=custom_config)

        if(firstPageText.find("Company Copy") != -1):
            print("Delete Company Copy")
            deleteFile(filePath)
        elif not "Affidavit" in name:
            policyNbrLeadingIndex = firstPageText.find("Policy Number: ") + 15
            if policyNbrLeadingIndex != -1:
                policyNbrTrailingIndex = firstPageText.find("\n",policyNbrLeadingIndex, len(firstPageText))
                
                plcyNbr = firstPageText[policyNbrLeadingIndex:policyNbrTrailingIndex]

                if(firstPageText.find("Agent Copy") != -1):
                    os.rename(filePath, os.path.join(destPath, plcyNbr+"_A.pdf") )
                else:
                    os.rename(filePath, os.path.join(destPath, plcyNbr+"_I.pdf") )




# Finds today's file paths and returns them as a generator
def filePaths():

    dirList = os.listdir(sourcePath)
    dateToday = date.today().strftime("%m%d%Y")

    # find directories marked with today's date that contain a Notices folder
    targetDirectory = [dir for dir in dirList if dateToday == dir[:8] and  "Notices" in os.listdir(os.path.join(sourcePath,dir))]

    #this shouldn't happen, throw error
    if len(targetDirectory) != 1:
        #throw error
        something = "something"


    batFile = os.path.join(sourcePath, targetDirectory[0], "Notices", "PrintNotices"+targetDirectory[0]+".bat")

    for line in open(batFile, 'r'):
        words = line.split(' ')
        filePath = words[-1]
        if "0_CIG_0" in filePath or "Affidavit" in filePath: 
            yield(filePath)

    

        





run()








