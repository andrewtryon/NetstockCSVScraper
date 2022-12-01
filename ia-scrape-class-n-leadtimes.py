from dotenv import load_dotenv
load_dotenv()
from selenium import webdriver
from selenium.webdriver.common.by import By
import subprocess
import pandas as pd
import os
import shutil
import time
import pyodbc
import chromedriver_autoinstaller


# method to get the downloaded file name
def getDownLoadedFileName(waitTime):
    driver.execute_script("window.open()")
    # switch to new tab
    driver.switch_to.window(driver.window_handles[-1])
    # navigate to chrome downloads
    driver.get('chrome://downloads')
    # define the endTime
    endTime = time.time()+waitTime
    while True:
        try:
            # get downloaded percentage
            downloadPercentage = driver.execute_script(
                "return document.querySelector('downloads-manager').shadowRoot.querySelector('#downloadsList downloads-item').shadowRoot.querySelector('#progress').value")
            # check if downloadPercentage is 100 (otherwise the script will keep waiting)
            if downloadPercentage == 100:
                # return the file name once the download is completed
                return driver.execute_script("return document.querySelector('downloads-manager').shadowRoot.querySelector('#downloadsList downloads-item').shadowRoot.querySelector('div#content  #file-link').text")
        except:
            pass
        time.sleep(1)
        if time.time() > endTime:
            break



print('I sleep for 10 seconds while chromedriver, installs')
chromedriver_autoinstaller.install()  # Check if the current version of chromedriver exists
                                      # and if it doesn't exist, download it automatically,
                                      # then add chromedriver to path

time.sleep(10)

driver = webdriver.Chrome()

driver.get("https://fotronic.sageinvadv.net/session/new")
username = driver.find_element(By.ID, "login_email")
password = driver.find_element(By.ID, "login_password")
time.sleep(5)

username.send_keys(os.environ.get("netstock_login"))
password.send_keys(os.environ.get("netstock_pw"))
driver.find_element(By.ID, "login_button").click()
time.sleep(5)

#get supplier calculated data
driver.get("https://fotronic.sageinvadv.net/suppliers.csv")
f = os.path.join('C:\\Users\\andrew.tryon\\Downloads\\', 'suppliers.csv')
print(f)
time.sleep(6)
shutil.move(f, os.path.join('\\\\FOT00WEB\\Alt Team\\Kris\\GitHubRepos\\inventory-advisor-scrape\\downloads\\','suppliers.csv'))
time.sleep(6)
print(os.path.join('\\\\FOT00WEB\\Alt Team\\Kris\\GitHubRepos\\inventory-advisor-scrape\\downloads\\', 'suppliers.csv'))
suppliersdf = pd.read_csv(os.path.join('\\\\FOT00WEB\\Alt Team\\Kris\\GitHubRepos\\inventory-advisor-scrape\\downloads\\', 'suppliers.csv'))
suppliersdf['VendorCode'] = suppliersdf['Supplier code'].str.split('-', expand=True)[1]
#'Measured LT days'

scrapeDests = {
    'A/H': 'A-HIGH',
    'A/M': 'A-MEDIUM',
    'A/L': 'A-LOW',
    'B/H': 'B-HIGH',
    'B/M': 'B-MEDIUM',
    'B/L': 'B-LOW',
    'C/H': 'C-HIGH',
    'C/M': 'C-MEDIUM',
    'C/L': 'C-LOW',
    'X/X': 'OBSOLETE'
}

compileddf = pd.DataFrame(data=None)

for key in scrapeDests:
    print(key, '->', scrapeDests[key])

    driver.get("https://fotronic.sageinvadv.net/stockenquiry/1/stock_by_matrix_cell/" + key + ".csv")
    f = os.path.join('C:\\Users\\andrew.tryon\\Downloads\\', key.split('/')[1] + '.csv')
    print(f)
    time.sleep(6)
    shutil.move(f, os.path.join('\\\\FOT00WEB\\Alt Team\\Kris\\GitHubRepos\\inventory-advisor-scrape\\downloads\\', scrapeDests[key] + '.csv'))
    time.sleep(6)
    print(os.path.join('\\\\FOT00WEB\\Alt Team\\Kris\\GitHubRepos\\inventory-advisor-scrape\\downloads\\', scrapeDests[key] + '.csv'))

    df = pd.read_csv(os.path.join('\\\\FOT00WEB\\Alt Team\\Kris\\GitHubRepos\\inventory-advisor-scrape\\downloads\\', scrapeDests[key] + '.csv'))
    df['Classification'] = scrapeDests[key]
    
    compileddf = compileddf.append(df, sort=False)    
    print(df)

#grabbing Non-Stock for data dumping
for i in range(5):
    i = i + 1

    driver.get("https://fotronic.sageinvadv.net/stockenquiry/1/stock_by_matrix_cell/N/N.csv?page=" + str(i))
    f = os.path.join('C:\\Users\\andrew.tryon\\Downloads\\N.csv')
    print(f)
    time.sleep(6)
    shutil.move(f, os.path.join('\\\\FOT00WEB\\Alt Team\\Kris\\GitHubRepos\\inventory-advisor-scrape\\downloads\\N.csv'))
    time.sleep(6)
    print(os.path.join('\\\\FOT00WEB\\Alt Team\\Kris\\GitHubRepos\\inventory-advisor-scrape\\downloads\\N.csv'))

    df = pd.read_csv(os.path.join('\\\\FOT00WEB\\Alt Team\\Kris\\GitHubRepos\\inventory-advisor-scrape\\downloads\\N.csv'))
    df['Classification'] = 'NON-STOCK'
    
    compileddf = compileddf.append(df, sort=False)    
    print(df)    

driver.close()
driver.quit()

print(compileddf)
compileddf['VendorCode'] = compileddf['Supplier code'].str.split('-', expand=True)[1]

#saving
compileddf.to_excel(r'\\FOT00WEB\Alt Team\Kris\GitHubRepos\inventory-advisor-scrape\downloads\Compiled.xlsx')

#compileddf = compileddf.loc[(compileddf['Classification'] != 'NON-STOCK')] 
print(compileddf)

#Establish sage connection
sage_conn_str = os.environ.get(r"sage_conn_str").replace("UID=;","UID=" + os.environ.get(r"sage_login") + ";").replace("PWD=;","PWD=" + os.environ.get(r"sage_pw") + ";")  
sage_cnxn = pyodbc.connect(sage_conn_str, autocommit=True)

#Audit Lead Times
SageSQLquery = """
SELECT IM_ItemVendor.ItemCode, IM_ItemVendor.VendorNo, IM_ItemVendor.StandardLeadTime, CI_Item.InactiveItem, CI_Item.PrimaryVendorNo
FROM IM_ItemVendor, CI_Item
WHERE IM_ItemVendor.ItemCode = CI_Item.ItemCode
"""
print('Retrieving Sage IM_ItemVendor data')
VendorDF = pd.read_sql(SageSQLquery,sage_cnxn)
VendorDF = VendorDF.dropna(subset=['PrimaryVendorNo'])
VendorDF = VendorDF.loc[(VendorDF['VendorNo'] == VendorDF['PrimaryVendorNo'])] 
VendorDF = VendorDF.loc[(VendorDF['InactiveItem'] != 'Y')] 
VendorDF = pd.merge(suppliersdf, VendorDF, how='right', left_on=['VendorCode'], right_on=['PrimaryVendorNo']).reset_index(drop=True)

auditLTdf = pd.merge(compileddf, VendorDF, how='right', left_on=['Product code','VendorCode'], right_on=['ItemCode','VendorNo']).reset_index(drop=True)

auditLTdf.loc[auditLTdf['LT days'].isna(), 'LT days'] = auditLTdf['Measured LT days']
auditLTdf.loc[auditLTdf['LT days'] == 0, 'LT days'] = auditLTdf['Measured LT days']

auditLTdf = auditLTdf.loc[(auditLTdf['LT days'] != auditLTdf['StandardLeadTime'])] 
auditLTdf = auditLTdf.loc[(auditLTdf['VendorNo'] != 'TBD')] 
auditLTdf = auditLTdf.loc[(auditLTdf['LT days'] != '')] 
auditLTdf = auditLTdf.dropna(subset=['LT days'])

if auditLTdf.shape[0] > 0:
    auditLTdf.to_csv(r'\\FOT00WEB\Alt Team\Kris\GitHubRepos\inventory-advisor-scrape\VI\leadtimes.csv', sep = ',', index=False, header=False, columns =['ItemCode','PrimaryVendorNo','LT days','StandardLeadTime'])
    print('VIing LeadTimes')
    time.sleep(120)
    p = subprocess.Popen('Auto_LeadTimes_VIWI7L.bat', cwd=r"Y:\Kris\GitHubRepos\inventory-advisor-scrape\VI", shell=True)
    stdout, stderr = p.communicate()
    p.wait()
    print('Sage VI Complete!')
    #time.sleep(360)
else:
    print('No lead time changes!')

#Audit Classifcations
SageSQLquery = """
SELECT CI_Item.ItemCode, CI_Item.Category3, CI_Item.InactiveItem
FROM CI_Item"""
print('Retrieving Sage  data')
ItemDF = pd.read_sql(SageSQLquery,sage_cnxn)

auditClassdf = pd.merge(compileddf, ItemDF, how='inner', left_on=['Product code'], right_on=['ItemCode']).reset_index(drop=True)

#Only Things that change
auditClassdf = auditClassdf.loc[(auditClassdf['Classification'] != auditClassdf['Category3'])] 

#We can leave NON-STOCK as blanks
auditClassdf = auditClassdf.loc[(auditClassdf['Classification'] != 'NON-STOCK') | ((auditClassdf['Classification'] == 'NON-STOCK') & (auditClassdf['Classification'] != 'NON-STOCK'))]
auditClassdf.loc[auditClassdf['Classification'] == 'NON-STOCK', 'Classification'] = ''

#We can ignore inactive items
auditClassdf = auditClassdf.loc[(auditClassdf['InactiveItem'] != 'Y')] 
if auditClassdf.shape[0] > 0:
    auditClassdf.to_csv(r'\\FOT00WEB\Alt Team\Kris\GitHubRepos\inventory-advisor-scrape\VI\classifications.csv', sep = ',', index=False, header=False, columns =['ItemCode','Classification','Category3'])
    print('VIing Classifications')
    time.sleep(120)
    p = subprocess.Popen('Auto_Classifications_VIWI7M.bat', cwd=r"Y:\Kris\GitHubRepos\inventory-advisor-scrape\VI", shell=True)
    stdout, stderr = p.communicate()
    p.wait()
    print('Sage VI Complete!')
else:
    print('No Classifications changes!')    