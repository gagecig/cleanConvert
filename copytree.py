import os
import pytesseract
from PyPDF2 import PdfFileReader
from pdf2image import convert_from_path

# To Launch: 
# 1. shift+ctrl+P to open comamnd pallete
# 2. click select interpreter
# 2. select interpreter C:\py\environments\autoPrint\Scripts\python.exe


sourcePath = r'C:\py\data\testSource'
destPath = r'C:\py\data\testDest'
batchPath = r'C:\py\data\PCL_Files\Notices\Cancellation'
filePath = r'C:\py\data\testDest\0_CIG_0_178207K7886689.pdf'
poppler_path = os.path.join(os.path.dirname(__file__), 'poppler-0.68.0_x86/bin')
custom_config = r'--oem 3 --psm 6'


def processDir(top):
    if not os.path.isdir(top):
            exit(1)
    for fileName in os.listdir(top):
        filePath = os.path.join(top, fileName)
        if not os.path.isfile(filePath):
            continue
        if not fileName.endswith('.pcl'):
            continue
        yield (top, fileName)
            
def convert_pdf(src_path, dst_path):
    cmd = "gpcl6win64.exe -dNOPAUSE -sDEVICE=pdfwrite -sOutputFile=%s %s" % (dst_path, src_path)
    os.system(cmd)

def pclToPdf(src_path, dst_path):              
    try:
        os.mkdir(dst_path)
    except:
        pass
    
    i = 0
    for (folder, name) in processDir(src_path):
        # print("Folder:  ", folder)
        # print("Name:  ", name)
        src_path_tmp = os.path.join(src_path, folder, name)
        dst_path_tmp = os.path.join(dst_path, "%s.pdf" % (name[:-4]))
        i = i + 1
        print(str(i)+": "+src_path_tmp+" -> "+dst_path_tmp)
        convert_pdf(src_path_tmp, dst_path_tmp)
        print("done.")


    

def deleteFile(filePath):
    if os.path.exists(filePath):
        os.remove(filePath)


def filterRename(dirPath):

    for (folder, name) in processDir(dirPath):

        pages = convert_from_path(filePath, dpi=350, poppler_path=poppler_path)
        text = pytesseract.image_to_string(pages[0], config=custom_config)

        if(text.find("Insured Copy") != -1):
            deleteFile(filePath)
        else:
            plcyNbrInx = text.find("Policy Number: ")
            if plcyNbrInx != -1:
                plcyNbr = text[plcyNbrInx+15:plcyNbrInx+30]
                print(plcyNbr)
                print(os.path.join(dirPath, plcyNbr+".pdf"))
                os.rename(filePath, os.path.join(dirPath, plcyNbr+".pdf") )



        
        
        






for (folder, name) in processDir(batchPath):
    print("Folder: ",folder)
    print("Name: ",name)
#filterRename(destPath)








