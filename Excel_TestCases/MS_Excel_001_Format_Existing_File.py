import win32com.client as win32
from win32com.client import constants
import win32gui, win32con, win32api
import win32print as printer
from zipfile import ZipFile
import pyautogui as imageGrab
import logging, time
import os, datetime

# Check if folder exist or not
lstfolder_check = ['resources', 'screenshot_tc1', 'reports']
for fldr in lstfolder_check:
    if not os.path.exists('..\\' + fldr):
        os.makedirs('..\\' + fldr)

# folder paths
screenshot_path = '..\\screenshot_tc1\\'
resources_path = '..\\resources\\'
report_path = '..\\reports\\'


# Create a logger file
def logInfoError(log, msg):
    timestamp = datetime.datetime.today().strftime('%Y%m%d%H%M%S')
    logName = 'report'+timestamp+'.log'
    logging.basicConfig(filename=report_path+logName,
                    format='%(asctime)s %(message)s',
                    filemode='w')
    logger = logging.getLogger(logName)
    logger.setLevel(logging.DEBUG)
    if log == 'error':
        logger.error(msg)
    else:
        logger.info(msg)

# Create the Excel Application
try:
    xlApp = win32.gencache.EnsureDispatch('Excel.Application')
    logInfoError('info', 'Step1-Excel Application Invoked Successfully')
except Exception as e:
    logInfoError('error', e)
    raise e

# Function to Connect to ALM QC and Navigate to Particular folder
# and Download the file to resources folder path
def qcConnect_Donwloadfile(qcServer, qcUser, qcPassword, qcDomain,
                           qcProject, qcTC_Folder, TestCase_Name):
    # Connect to qc Server
    try:
        qcConn = win32.Dispatch('TDApiOle80.TDConnection.1')
        qcConn.InitConnectionEx(qcServer)
        qcConn.Login(qcUser, qcPassword)
        qcConn.Connect(qcDomain, qcProject)
        im = imageGrab.grab()
        im.save(screenshot_path + 'Step2_Connect_QCServer.png')
        logInfoError('info', 'Step2-Connect to QC Successfully')

        # Download file attached to test cases
        TreeObj = qcConn.TreeManager
        folder = TreeObj.NodeByPath(qcTC_Folder)
        testList = folder.FindTests(TestCase_Name)
        if (len(testList)) == 0:
            im = imageGrab.grab()
            im.save(screenshot_path + 'Step2_No_TC_Found.png')
            logInfoError('error', 'No TC found in folder-{}'.format(qcTC_Folder))
        else:
            for tst in range(len(testList)):
                teststorage = testList[tst].ExtendedStorage
                teststorage.ClientPath = resources_path + testList[tst].name
                teststorage.Load('', True)
                time.sleep(5)
                im = imageGrab.grab()
                im.save(screenshot_path + 'Step2_Downloadfile.png')
                logInfoError('info', 'Stpe2-Completed Download')
    except Exception as e:
        im = imageGrab.grab()
        im.save(screenshot_path + 'Step2_QCServer_Error.png')
        logInfoError('error', 'Could not Connect to QC Server-{}'.format(e))

# Function to extract the Zip file and Save the unzip file to resource path
def extract_ZipFile(resources_path):
    try:
        filelist = os.listdir(resources_path)
        for file in filelist:
            if file.startswith('Basic_Test'):
                with ZipFile(resources_path + file, 'r') as zip:
                    zip.printdir()
                    zip.extractall(resources_path)
        os.startfile(resources_path)
        time.sleep(1)
        im = imageGrab.grab()
        im.save(screenshot_path + 'Step2_Unzip_file.png')
        logInfoError('info', 'Step2-Unzip file {} successfully'.format(zip.infolist()[0].filename))
        return (zip.infolist()[0].filename)
    except Exception as e:
        im = imageGrab.grab()
        im.save(screenshot_path + 'Step2_File Not found.png')
        logInfoError('error','Step2-Could not Unzip file {}'.format(e))
        raise e

# Function to Type formula and copy the same and paste to alternate cells
def xl_type_CopyPaste(unzipfileName, xlpath):
    # Type cell with value and copy paste the value to different cells
    try:
        path = os.getcwd().replace('\'', '\\') + '\\'
        wb = xlApp.Workbooks.Open(path + xlpath + unzipfileName)
        ws = wb.Worksheets(1)
        xlApp.Visible = True
        im = imageGrab.grab()
        im.save(screenshot_path + 'Step1_Validate_Application.png')
        hwnd = win32gui.GetForegroundWindow()
        win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
        ws.Cells(28, 3).Value = '=COUNTA(C8:C24)'
        wb.Save()
        im = imageGrab.grab()
        im.save(screenshot_path + 'Step3to4_Type_Formula.png')
        logInfoError('info','Step3to4-Type formula to Cell C28 Successfully')
        time.sleep(1)
        wCol = 'E'
        for col in range(0, 5):
            ws.Range('C28').Copy()
            ws.Range(wCol + '28').PasteSpecial(Paste=constants.xlPasteValues)
            wCol = chr(ord(wCol) + 2)
        wb.Save()
        im = imageGrab.grab()
        im.save(screenshot_path + 'Step5_Copy_Cell_Paste.png')
        logInfoError('info','Step5-Copy C28, Paste Cell to Alternate Cells')
        return(wb, ws)
    except Exception as e:
        im = imageGrab.grab()
        im.save(screenshot_path + 'Step5_Copy_Cell_Paste_Error.png')
        logInfoError('error','Step5-Could not Open Excel {}'.format(e))
        raise e

# Function to format all the Copied cells
def xl_format_Cells(wb, ws):
    try:
        y = 'C'
        for j in range(0, 6):
            ws.Range(y + '28').HorizontalAlignment = 3
            strCol = '%02x%02x%02x' % (0, 165, 255)
            ws.Range(y + '28').Interior.Color = int(strCol, 16)
            for id in range(7, 13):
                ws.Range(y + '28').Borders(id).LineStyle = 1
                ws.Range(y + '28').Borders(id).Weight = 2
            y = chr(ord(y) + 2)
        wb.Save()
        im = imageGrab.grab()
        im.save(screenshot_path + 'Step6to9_Format_Cells.png')
        logInfoError('info','Step6to9-Format Cells and Highlight')
    except Exception as e:
        im = imageGrab.grab()
        im.save(screenshot_path + 'Step8to9_Couldnot_Format_Cells.png')
        logInfoError('error','Step6to9-Could not format the cells')
        raise e

# Function to Print and Close the Xl file
def print_file_close(wb, flpath, flname):
    try:
        Cur_printer = printer.GetDefaultPrinter()
        print(Cur_printer)
        printer.SetDefaultPrinter(Cur_printer)
        wb.Save()
        wb.Close()
        win32api.ShellExecute(0, 'print', flpath+flname, Cur_printer, ',', 0)
        im = imageGrab.grab()
        im.save(screenshot_path + 'Step10to11_Print_Close_File.png')
        logInfoError('info','Step10to11-Successfully printed file {} on printer {}'.format(flname, Cur_printer))
    except Exception as e:
        im = imageGrab.grab()
        im.save(screenshot_path + 'Step10to11_Print_error.png')
        logInfoError('error','Step10to11-Could not print file {} on printer {}'.format(flname, Cur_printer))
        raise e

if __name__ == '__main__':
    unzipfilename = extract_ZipFile(resources_path)
    xlws = xl_type_CopyPaste(unzipfilename, resources_path)
    xl_format_Cells(xlws[0], xlws[1])
    print_file_close(xlws[0], resources_path, unzipfilename)