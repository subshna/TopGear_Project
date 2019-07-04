import win32com.client as win32
from win32com.client import constants
import win32gui, win32con
from zipfile import ZipFile
import pyautogui as imageGrab
import logging, time
import os, datetime

# Check if folder exist or not
lstfolder_check = ['resources', 'screenshot_tc2', 'reports']
for fldr in lstfolder_check:
    if not os.path.exists('..\\' + fldr):
        os.makedirs('..\\' + fldr)

# folder paths
screenshot_path = '..\\screenshot_tc2\\'
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
            if file.startswith('Adobe'):
                with ZipFile(resources_path + file, 'r') as zip:
                    #zip.printdir()
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

# Function open the file
def validate_xl(unzipfileName, xlpath):
    # Type cell with value and copy paste the value to different cells
    try:
        path = os.getcwd().replace('\'', '\\') + '\\'
        wb = xlApp.Workbooks.Open(path + xlpath + unzipfileName)
        xlApp.Visible = True
        hwnd = win32gui.GetForegroundWindow()
        win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
        time.sleep(1)
        im = imageGrab.grab()
        im.save(screenshot_path + 'Step3_Open_Data_File.png')
        logInfoError('info', 'Step3-File Opened Successfully')

        # Select the Source Range
        ws = wb.Sheets('Sheet1')
        PivotRangeSrc = ws.UsedRange
        PivotRangeSrc.Select()
        return (wb, PivotRangeSrc)
    except Exception as e:
        logInfoError('error', 'Step3-Data file Not Found {}'.format(e))
        raise e

# Function to create the Pivot table based on the source and destination Range
def create_PivotTable(wb, PivotRangeSrc, PivotTblName, stepNo):
    try:
        # Select destination Range
        PivotSht = wb.Worksheets.Add()
        PivotRangeDest = PivotSht.Range('A1')
        PivotTableName = PivotTblName
    except Exception as e:
        logInfoError('error', 'Select Destination Error {}'.format(e))
        raise e

    # Create Pivot table
    try:
        PivotCache = wb.PivotCaches().Create(SourceType=constants.xlDatabase, SourceData=PivotRangeSrc)
        PivotTable = PivotCache.CreatePivotTable(TableDestination=PivotRangeDest, TableName=PivotTableName)
        time.sleep(0.5)
        im = imageGrab.grab()
        im.save(screenshot_path + stepNo+'_Highlight_All_Column.png')
        logInfoError('info', stepNo+'-Highlighted all columns')
        return (PivotTable, PivotSht)
    except Exception as e:
        logInfoError('error', stepNo+'-Range Selection error {}'.format(e))
        raise e

# Function to Drag the Columns and Sort
def select_PivotFields_Sort(PivotTable, *argv):
    try:
        PivotTable.PivotFields(argv[0]).Orientation = argv[2]
        DataField = PivotTable.AddDataField(PivotTable.PivotFields(argv[1]))
        DataField.NumberFormat = '##0.00'
        time.sleep(0.5)
        im = imageGrab.grab()
        im.save(screenshot_path + 'Step6_Pivot_Table_Loaded.png')
        logInfoError('info', 'Step6-Pivot table should loaded as per selection')

        # Sort field Hostname descending order
        PivotTable.PivotFields(argv[0]).AutoSort(constants.xlAscending, DataField)
        time.sleep(0.5)
        im = imageGrab.grab()
        im.save(screenshot_path + 'Step7_Pivot_Table_Sorted.png')
        logInfoError('info', 'Step7-The pivot table should load with sorting')
    except Exception as e:
        logInfoError('error', 'Step6to7-Pivot loading and sorting error {}'.format(e))
        raise e

# Function to Drag the Columns, Create chart, and Close the file
def select_PivotFields_Chart(PivotTable, PivotSht, wb, *argv):
    try:
        PivotTable.PivotFields(argv[0]).Orientation = argv[3]
        PivotTable.PivotFields(argv[0]).Position = 1
        PivotTable.PivotFields(argv[1]).Orientation = argv[3]
        PivotTable.PivotFields(argv[0]).Position = 2
        time.sleep(0.5)
        im = imageGrab.grab()
        im.save(screenshot_path + 'Step10_Columns_to_Axis.png')
        logInfoError('info', 'Step7-Columns are dragged successfully to Axis')

        DataField = PivotTable.AddDataField(PivotTable.PivotFields(argv[2]))
        DataField.NumberFormat = '##0.00'
        time.sleep(0.5)
        im = imageGrab.grab()
        im.save(screenshot_path + 'Step11_Columns_to_Axis.png')
        logInfoError('info', 'Step11-Columns are dragged successfully to Axis')

        # Create Chart
        chart = PivotSht.Shapes.AddChart2(201)
        time.sleep(0.5)
        im = imageGrab.grab()
        im.save(screenshot_path + 'Step12_Pivot_Chart.png')
        logInfoError('info', 'Step12-Pivot chart should be loaded')
        wb.Save()
        wb.Close()
    except Exception as e:
        logInfoError('error', 'Step10tp12-Columns and Chart not created {}'.format(e))
        raise e


# Call the main function
if __name__ == '__main__':
    PivotFields = {1 : 'UserName',
                   2 : 'Host Name',
                   3 : 'User Region',
                   4 : 'Count of Host Name'
    }
    PivotOrient = {1 : constants.xlRowField,
                   2 : constants.xlPageField,
                   3 : constants.xlColumnField,
    }
    unzipfilename = extract_ZipFile(resources_path)
    xlwb_Src = validate_xl(unzipfilename, resources_path)

    # Create first pivot table
    PivotTbl = create_PivotTable(xlwb_Src[0], xlwb_Src[1], 'PivotTable1','Step4to5')
    select_PivotFields_Sort(PivotTbl[0], PivotFields[1],PivotFields[2], PivotOrient[1])

    # Create Second Pivot table
    PivotTbl = create_PivotTable(xlwb_Src[0], xlwb_Src[1], 'PivotTable2', 'Step8to9')
    select_PivotFields_Chart(PivotTbl[0], PivotTbl[1], xlwb_Src[0], PivotFields[1],
                             PivotFields[3], PivotFields[2], PivotOrient[2])