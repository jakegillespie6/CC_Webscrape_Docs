from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium import webdriver
from selenium.webdriver import ChromeOptions, ActionChains
from chromedriver_py import binary_path
import re, time, os, warnings
import pandas as pd
import warnings 
import xlsb_converter as xlsb
import warnings
warnings.filterwarnings("ignore")

directory_write_to = r'C:\Users\jgillespie\Crown Castle USA Inc\West Region VZW VCA2 - SCAFs\Background Files'

directory = r"C:\Users\jgillespie\OneDrive - Crown Castle USA Inc\_SoCal ROW Compiled\RoW LiveSCAF Automation\Storage"
#PNW Directories
directory_pnw_amd = r"C:\Users\jgillespie\Crown Castle USA Inc\West Region VZW VCA2 - SCAFs\Live Consolidated SCAFs\_PNW\3 - Active AMD"
directory_pnw_contracted = r"C:\Users\jgillespie\Crown Castle USA Inc\West Region VZW VCA2 - SCAFs\Live Consolidated SCAFs\_PNW\2 - Contracted SCAF"
directory_pnw_active = r"C:\Users\jgillespie\Crown Castle USA Inc\West Region VZW VCA2 - SCAFs\Live Consolidated SCAFs\_PNW\1 - Active SCAF"
#RMR Directories
directory_rmr_amd = r"C:\Users\jgillespie\Crown Castle USA Inc\West Region VZW VCA2 - SCAFs\Live Consolidated SCAFs\_RMR\3 - Active AMD"
directory_rmr_contracted = r"C:\Users\jgillespie\Crown Castle USA Inc\West Region VZW VCA2 - SCAFs\Live Consolidated SCAFs\_RMR\2 - Contracted SCAF"
directory_rmr_active = r"C:\Users\jgillespie\Crown Castle USA Inc\West Region VZW VCA2 - SCAFs\Live Consolidated SCAFs\_RMR\1 - Active SCAF"
#DSW Direcotries
#RMR Directories
directory_dsw_amd = r"C:\Users\jgillespie\Crown Castle USA Inc\West Region VZW VCA2 - SCAFs\Live Consolidated SCAFs\_DSW\3 - Active AMD"
directory_dsw_contracted = r"C:\Users\jgillespie\Crown Castle USA Inc\West Region VZW VCA2 - SCAFs\Live Consolidated SCAFs\_DSW\2 - Contracted SCAF"
directory_dsw_active = r"C:\Users\jgillespie\Crown Castle USA Inc\West Region VZW VCA2 - SCAFs\Live Consolidated SCAFs\_DSW\1 - Active SCAF"

directory_norCal_contracted = r"C:\Users\jgillespie\Crown Castle USA Inc\West Region VZW VCA2 - SCAFs\Live Consolidated SCAFs\_NorCal\2 - Contracted SCAF"
directory_norCal_amd = r"C:\Users\jgillespie\Crown Castle USA Inc\West Region VZW VCA2 - SCAFs\Live Consolidated SCAFs\_NorCal\3 - Active AMD"
#SoCal Directories
directory_SoCal_contracted = r"C:\Users\jgillespie\Crown Castle USA Inc\West Region VZW VCA2 - SCAFs\Live Consolidated SCAFs\_SoCal\2 - Contracted SCAF"

dirMap = {directory_dsw_contracted:"DSW Contracted.xlsx", 
          directory_dsw_amd:"DSW Amd.xlsx", 
          directory_pnw_contracted:"PNW Contracted.xlsx", 
          directory_pnw_amd:"PNW Amd.xlsx", 
          directory_rmr_contracted:"RMR Contracted.xlsx", 
          directory_rmr_amd:"RMR Amd.xlsx", 
          directory_norCal_contracted:"NorCal Contracted.xlsx", 
          directory_norCal_amd:"NorCal Amd.xlsx", 
          directory_SoCal_contracted:"SoCal Contracted.xlsx"}


#Socal opportunity IDs
Socal_oppIDs = [60648790,60648791,68887104, 68935195, 60648792,60648794,60648798,60648801,60648805,60648806,60648807,60648808,60648809,60648814,60648825,60648853,60648854,60648855,60648857,60648859,60648862,60651855,60651856,60651857,60651860,60651877,60651880,60651882,60651884,60651886,60651887,60651898,60651900,60651902,60651906,60651907,60651913,60651914,60651985,60651989,60651990,60651991,61593777,67468143,68977187]
socal_amd_id = [62847816,62847817,62847822,62847823,62847824,62847825,62847826,62847827,62847828,62847835,62847836,62847837,62847838,62847839,62847842,62850776,62850777,62850779,62850780,62850781,62850782,62850784,62850785,62850786,62850787,62850789,62850790,62850795,64500916,65115779,65302031,65346787,65877768,65880767,66975927,67314986,68022964,69217167,69657356,69657359,69657360,69657362,69657363,69657376,69657377,69657412,69657413,69657417,69657422,69673527,69673530,69673531,69673532,69673533,69673628,69673631,69673632,69673675,69676302,69676303,69676305,69676306,69676307,69676308,69676309,69676310,69676311,69676312,69676321,69676323,69676341,69676346,69679288,69679289,69679292,69679340,69679341,69679501,69679502,69679570,69682548]
#NorCal opportunity IDs
norcal_oppIDs = [60528836,60528842,60528847,60528929,60531742,60531749,60531751,60531758,64497782,64656778]
norcal_amd_ID = [62763827,62763828,62766838,63513789,63573785,63573789,64080763,66339767,66663840,67512932,69679622]
#PNW opportunity IDs
pnw_oppIDs = [60576783,60603808,60654772,60654773,60657734,60660762,60660805,60663738,60663739,60663779,60663780,60663782,61776801,62304802,66904074,66904091,66904102,66907007,68857141]
pnw_amd_ID = [69028112,69031104,69031122,69031123,69502400,69508555,69796355,69796356,69805211,69805212,69805227,69805228]
#DSW opportunity IDs
dsw_oppIDs = [61077829,67023925,67023926,67023929,67023930,67023931,67023932,67023933,67023935,67026916,67026918,67026919,67026920,67026922,67026931,67026932,67026933,67026934,67026935,67026936,67498042,69124142]
dsw_amd_ID = [62790760, 67677918, 68887242]
#San Diego opportunity IDs
sd_oppIDs = [60651877]
#RMR opportunity IDs
rmr_oppIDs = [60600805,60600841,60600877,60600878,60603770]
rmr_amd_ID = [63006937,68112925,69524572,69524599,69535225,69538493,69538494,69538532,69544500,69598293]

def InitializeDriver(dir):
    service_object = Service(binary_path)
    prefs = {
        "download.default_directory": dir,
        "download.directory_upgrade": True,
        "download.prompt_for_download": False,
    }
    chromeOptions = ChromeOptions()
    chromeOptions.add_experimental_option("prefs",prefs)
    driver = webdriver.Chrome(service=service_object, options=chromeOptions)
    chromeOptions.add_argument('--headless=new')
    driver.get("http://ccireports.crowncastle.com/ibi_apps/WFServlet?IBIF_ex=/WFC/Repository/smallcell/Business_Managed_Reports/sc_classification_dashboard/VZW_VCA2_SCAF.fex&IBIMR_drill=IBFS,RUNFEX,IBIF_ex,false")
    time.sleep(3)
    driver.refresh()
    return driver

def downloadSCAF(driver, oppIDs,dir):
    try:
        openPaneBtn = WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="header"]/a')))
        oppInput = WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ui-id-1"]')))
        submitBtn = WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="promptPanel"]/div/div/div[1]/a[4]')))
        
    except:
        openPaneBtn = WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="header"]/a')))
        openPaneBtn.click()
        oppInput = WebDriverWait(driver,1).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ui-id-1"]')))

    for i,id in enumerate(oppIDs):
        try:
            oppInput.clear()
        except:
            openPaneBtn.click()
            time.sleep(1)
            oppInput.clear()
        oppInput.send_keys(id)
        submitBtn.click()

        j=0
        while (j <= i):
            time.sleep(5)
            j=0
            for file in os.listdir(dir):
                j+=1
    driver.close()


def assembleDataFrames(siteConfigAppList, siteDetailList, equipDetailList, dir):
    for filename in os.listdir(dir):
        f = os.path.join(dir,filename)
        if os.path.isfile(f):
            thisEquipment = pd.read_excel(f, 'Equip Detail',usecols="A:H", header=0)
            equipDetailList.append(thisEquipment)

            thisOppName = pd.read_excel(f,'Site Config App Form').iloc[2,1]
            thisOppNum = pd.read_excel(f,'Site Config App Form').iloc[1,15]
            thisSiteConfigApp = pd.read_excel(f, 'Site Config App Form',usecols="A:AN",header=6)
            thisSiteConfigApp['Opportunity Name'] = thisOppName
            thisSiteConfigApp['Opportunity ID'] = thisOppNum
            siteConfigAppList.append(thisSiteConfigApp)

            thisSiteDetail = pd.read_excel(f, 'Site Detail',usecols="A:AA", header=8)
            #thisSiteDetail.drop(index=thisEquipment.index[0:8],inplace=True)
            siteDetailList.append(thisSiteDetail)

def to_excel_(siteConfigAppList, equipDetailList, siteDetailList, path):
    try:
        siteConfigAppDF = pd.concat(siteConfigAppList)
        siteDetailDF = pd.concat(siteDetailList)
        equipDetailDF = pd.concat(equipDetailList)
    except:
        print("Empty Directory in: ", path)
        return

    with pd.ExcelWriter(path, mode="a", engine="openpyxl",if_sheet_exists='overlay') as writer:
        siteConfigAppDF.to_excel(writer, sheet_name="Site Config App Form", header=None, index=False, startrow=1)
        equipDetailDF.to_excel(writer, sheet_name="Equipment Detail", header=None, index=False, startrow=1) 
        siteDetailDF.to_excel(writer, sheet_name="Site Detail", header=None, index=False, startrow=1) 


def rm_SCAFS(directory):
    for file in os.listdir(directory):
        os.remove(os.path.join(directory,file))


for dir in dirMap:
    rm_SCAFS(dir)




driver = InitializeDriver(directory_SoCal_contracted)
downloadSCAF(driver, Socal_oppIDs, directory_SoCal_contracted)

driver = InitializeDriver(directory_norCal_contracted)
downloadSCAF(driver, norcal_oppIDs, directory_norCal_contracted)

driver = InitializeDriver(directory_pnw_contracted)
downloadSCAF(driver, pnw_oppIDs, directory_pnw_contracted)

driver = InitializeDriver(directory_dsw_contracted)
downloadSCAF(driver, dsw_oppIDs, directory_dsw_contracted)

driver = InitializeDriver(directory_rmr_contracted)
downloadSCAF(driver, rmr_oppIDs, directory_rmr_contracted)

#assembleDataFrames(siteConfigAppList=siteConfigAppList, siteDetailList=siteDetailList, equipDetailList=equipDetailList, dir=directory_norCal_contracted)
#to_excel_(siteConfigAppList=siteConfigAppList, siteDetailList=siteDetailList, equipDetailList=equipDetailList, filename=r'C:\Users\jgillespie\Crown Castle USA Inc\West Region VZW VCA2 - SCAFs\Background Files\NorCal SCAF.xlsx')


for dir in dirMap:
    siteConfigAppList = []
    siteDetailList = []
    equipDetailList = []
    dir_file = os.path.join(directory_write_to, dirMap[dir])
    print(dir_file)
    print(dir, '\n')

    assembleDataFrames(siteConfigAppList,siteDetailList, equipDetailList,  dir)
    to_excel_(siteConfigAppList, equipDetailList, siteDetailList, dir_file)

