from selenium import webdriver #Web driver activities
import logging #Activity logging
from configparser import ConfigParser #Configuration read
import os #Path
import sys #System paths
import re
import csv
import ctypes #message box
from selenium.webdriver.firefox.options import Options
import pyodbc
from selenium import webdriver
from selenium.webdriver.firefox.firefox_binary import FirefoxBinary
from selenium.webdriver.common.by import By #Element find
from selenium.webdriver.support import expected_conditions as EC #Error handling
from selenium.webdriver.support.ui import WebDriverWait #Web driver wait
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
import time #Sleep
from pathlib import Path #Path conversions
import pyautogui
from time import sleep # this should go at the top of the file
import pygetwindow
import shutil

download_dir = os.path.dirname(os.path.realpath(__file__))+'\\temp_downloads'
ConfigPath = os.path.dirname(os.path.realpath(__file__)) + '\\config.ini'
logtemplate = "An exception of type {0} occurred. Arguments:\n{1!r}"
firefox_location='./'

global wait, driver, webpage_url, user_name, password, sq_1, sq_2, sq_3, sq_4,onlyfiles
global sql_server_name, sql_user_name, sql_password, sql_db, SQLconnection

def read_config_file():
    global webpage_url, user_name, password, sq_1, sq_2, sq_3, sq_4
    global sql_server_name, sql_user_name, sql_password, sql_db
   
    try:
        #configuration entries
        config = ConfigParser()
        config.read(ConfigPath)
        webpage_url = config.get ("Data", "webpage_URL")
        user_name = config.get ("Data", "user_name")
        password = config.get ("Data", "password") 
        sql_server_name = config.get ("SQL", "server")
        sql_user_name = config.get ("SQL", "user")
        sql_password = config.get ("SQL", "password")
        sql_db = config.get ("SQL", "database")
        return(0)
    except Exception as ex:
        print('config file read error')
        return(-1)


def setup():
    global firefox_location, download_dir, webpage_url, driver, wait
    binary = FirefoxBinary(firefox_location)
    profile = webdriver.FirefoxProfile()
    profile.set_preference("browser.download.folderList", 2)
    profile.set_preference("browser.download.manager.showWhenStarting", False)
    profile.set_preference("browser.download.dir", download_dir)
    profile.set_preference("browser.helperApps.neverAsk.saveToDisk", "text/css, application/vnd.openxmlformats-officedocument.wordprocessingml.document, application/octet-stream, application/word, application/wordpad, image/png, image/bmp, image/jpeg, application/pdf, text/csv, text/html, text/plain, application/docx")
    profile.set_preference("pdfjs.disabled", True)
    profile.set_preference("plugin.disable_full_page_plugin_for_types", "text/css, application/vnd.openxmlformats-officedocument.wordprocessingml.document, application/octet-stream, application/word, application/wordpad, image/png, image/bmp,, image/jpeg, application/pdf, text/csv, text/html, text/plain, application/docx")
    options = Options()
    options.profile = profile
    driver = webdriver.Firefox(options=options)
    wait = WebDriverWait(driver, 10)
    driver.maximize_window()
    driver.get(webpage_url)
    time.sleep(1)
    wait = WebDriverWait(driver, 10)
    return 0


def rename_file(file_name):
    file_name_2 = file_name
    for each in file_name:
        if each == "'":
            file_name_2 = file_name.replace("'","''")
    return file_name_2

def updateToDatabase(src_id,dst_id,emp_name,fname):
    global SQLconnection
    try:
        cursor = SQLconnection.cursor()
        update_statement = """UPDATE [RoadSafe].[dbo].[LMS Files] SET [IsUploaded] = ?, [UploadedDateTime] = ? 
                        WHERE  [Employee_ID] =? AND [File_Name] = ? ;
                        """
        cursor.execute(update_statement,(1,time.strftime('%Y-%m-%d %H:%M:%S'),src_id,fname))
        SQLconnection.commit()   
        print(dst_id," : updated in Emp_Files")     
    except Exception as e:
        print('update to database error',e)
        return(-1)
    return 0

def searchAndUpload(ufiles, fld_name,  dst_id,emp_name, category,src_id):
    try:
        sea_flg = 0
        sea_cnt = 0
        while(1):
            try:
                sea_cnt = sea_cnt + 1
                try:
                    xpath = "//a[@href='//HCM.paycor.com/people/?companyId=194572&clientId=157905'][contains(text(),'Manage People')]" #user name
                    element = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
                    element.click()
                    time.sleep(3)
                except:
                    try:
                        xpath = "//a[@href='//HCM.paycor.com/people/?companyId=194573&clientId=148774'][contains(text(),'Manage People')]" #user name
                        element = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
                        element.click()
                        time.sleep(3)
                    except:
                        pass
                driver.get("https://hcm.paycor.com/people/?companyId=198349&clientId=153785")
                time.sleep(2)
                paycor_employee=dst_id
                print(paycor_employee)
                xpath  = "//input[contains(@id,'search_')]"
                driver.find_elements(By.XPATH, xpath).clear()
                element = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
                element.click()
                time.sleep(2)
                element.send_keys(paycor_employee)
                time.sleep(2)
                break
            except:
                if sea_cnt > 5:
                    sea_flg = 1
                    break
                else:
                    pass
        if sea_flg == 1:
            return(-1)

        xpath = "//div[contains(@id,'personCard_1')]"
        elements = driver.find_elements(By.XPATH,xpath)
        no_of_employees = len(elements)
        time.sleep(2)
        for i in range(0,no_of_employees):
            
            xpath1 = "//div[contains(@id,'personCard_1_"+str(i)+"')]/.//div[contains(@class,'empNumber')]"
            element1 = driver.find_element(By.XPATH, xpath1).get_attribute("title")
            if str(element1).strip()=='#'+str(paycor_employee):
                #xpath3 = ""
                xpath2 = "//div[contains(@id,'personCard_1_"+str(i)+"')]/.//span[@class='name secondaryCTA']"
                element2 = wait.until(EC.element_to_be_clickable((By.XPATH, xpath2)))
                element2.click()
                break
           
        time.sleep(10)
        try:
            xpath = '//*[@id="5000"]' # Personal
            element = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
            element.click()
            time.sleep(3)
        except:
            pass

        try:
            # Certificates
            xpath = "//a[@aria-label='Certifications, maydd menu item to favorites']//span[@class='node-leaf-text__3J_vd']"
            element = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
            element.click()
            time.sleep(2)
        except:
            try:
                xpath = '//*[@aria-label="Certifications, may Add menu item to favorites"]'
                element = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
                element.click()
                time.sleep(2)
            except:
                pass
        print(ufiles)
        
        for ufile in ufiles:
            down_flg = 0
            down_count = 0
            while(1):
                try:
                    try:
                        xpath = '//*[@class="Certification_alignCenter__CnrZp Certification_verticalOffset__yqhpu Certification_noRecordsBorder__XUgnW"]'
                        element = driver.find_element(By.XPATH,xpath)
                        txt = element.text
                        if txt == 'No certifications to display':
                            with open('No certificates.csv', 'a',encoding='utf-8') as f:
                                f.write(f'{src_id},"{emp_name}",{txt}\n')
                            return 'No Cert'
                    except:
                        print("looping file:",ufile)
                        down_count = down_count + 1
                        certificate_name = category[ufile]['mapping']
                        i=1
                        elements = wait.until(EC.presence_of_all_elements_located((By.XPATH, '//*[@class="FileChooser_fileChooserInput__d3Sgp"]')))
                        # input_element.send_keys(file_path)
                        for element in elements:
                            print("i valu is:",i)
                            
                            # print(element)
                            try:
                                xpath = '/html/body/div[1]/div/div/div[2]/div[3]/div/div['+str(i)+']/div/div/div/div[1]/div[1]/div[1]'
                                cert_path = driver.find_element(By.XPATH,xpath)
                                # for element in elements:
                                cert_name = cert_path.text
                                print(element.text)
                                print(f"Certificate found: {cert_name}, Expected: {certificate_name}")
                                k=1
                                xpath = '/html/body/div[1]/div/div/div[2]/div[3]/div/div['+str(i)+']/div/div/div/div[2]/div[4]/div[2]/div/div/a'
                                files_path = driver.find_elements(By.XPATH,xpath)
                                for file_path in files_path:
                                    print("file name before upload:",file_path.text)
                                    time.sleep(2)
                                    txt_before = file_path.text
                                    if txt_before != '':
                                        with open('files before upload.csv', 'a',encoding='utf-8') as f:
                                            f.write(f'{src_id},"{emp_name}","{txt_before}",{cert_name}\n')
                                
                                if cert_name.lower() == certificate_name.lower() or certificate_name.lower() in cert_name.lower():
                                    file_path = f'S:\\Cycle 1\\LMS\\Active\\{fld_name}\\{ufile}'
                                    print(f"File path to upload: {file_path}")
                                    try:
                                        element.send_keys(file_path)
                                        print("File uploaded successfully!")
                                        time.sleep(3)                 
                                        driver.refresh()
                                        time.sleep(5)
                                        # extracting file naem after uploading
                                        xpath = '/html/body/div[1]/div/div/div[2]/div[3]/div/div['+str(i)+']/div/div/div/div[2]/div[4]/div[2]/div/div/a'
                                        files_path = driver.find_elements(By.XPATH,xpath)
                                        for file_path in files_path:
                                            print("file name after upload:",file_path.text)
                                            txt_after = file_path.text
                                            time.sleep(2)
                                            with open('files after upload.csv', 'a',encoding='utf-8') as f:
                                                f.write(f'{src_id},"{emp_name}","{txt_after}",{cert_name}\n')
                                        value = dst_id+","+emp_name+","+","+ufile
                                        print(value)
                                        f= open('uploaded_Terminated.csv',"a")
                                        f.write(f'{dst_id},"{emp_name}","{ufile}"')
                                        f.write('\n')
                                        f.close()
                                        res = updateToDatabase(src_id,dst_id,emp_name,ufile)
                                        if res<0:
                                            return(-1)
                                        driver.refresh()
                                        time.sleep(5)
                                        i=i+1
                                        break
                                    except Exception as inner_exception:
                                        print(f"Error while sending file path: {inner_exception}")
                                else:
                                    print("Certificate name not found in sytem")
                                    # with open('f upload.csv', 'a',encoding='utf-8') as f:
                                    #     f.write(f'{src_id},"{emp_name}","{ufile}",{certificate_name}\n')
                                    i=i+1
                                    continue
                                
                            except Exception as f:
                                print("error while uploading document:",f)
                                               

                        
                        break
                except:
                    if down_count > 2:
                        down_flg = 1
                        break
                    else:
                        pass
            if down_flg == 1:
                return(-1)
    except Exception as g:
        print('employee name error',g)
        f= open('employee name error.csv',"a")
        f.write(dst_id)
        f.write('\n')
        f.close()
        try:
            xpath = '/html/body/div[1]/div/paycor-employee-navigation-bar/quick-links/button'
            element = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
            element.click()
            time.sleep(3)  
        except:
            pass
        return -1
    try:
        xpath = "//*[@id='Header_NavigateToEmployeeListButton']"
        element = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
        element.click()
        time.sleep(3)  
    except:
        try:
            xpath = "/html/body/div[1]/paycor-employee-navigation-bar/quick-links/button"
            element = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
            element.click()
            time.sleep(3)
        except: pass
    return 0

       

def login():
    global user_name, driver, password, wait, driver, sq_1, sq_2, sq_3, sq_4
    # window_before = driver.window_handles[0]
    time.sleep(3)
    # xpath = "//a[@href='https://hcm.paycor.com/authentication/signin']" #user name
    # element = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
    # element.click()
    time.sleep(5)
    # window_after = driver.window_handles[1]
    # driver.switch_to.window(window_after)
    xpath = "//input[@id='Username']" #user name 
    element = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
    element.click()
    time.sleep(2)
    element.send_keys(user_name)
    time.sleep(2)
    xpath = "//input[@id='Password']" #password
    element = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
    element.click()
    time.sleep(2)
    element.send_keys(password)
    time.sleep(2)    
    xpath = "//button[contains(text(),'Sign In')]" #login button
    element = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
    element.click()
    time.sleep(2)
    time.sleep(30)
    s=input()
    print("login success")


def connectSQL():
    global sql_server_name, sql_user_name, sql_password, sql_db, SQLconnection
    try:
        SQLconnection = pyodbc.connect('Driver={SQL Server};'
                            'Server=' + sql_server_name + ';'
                            'Database=' + sql_db + ';'
                            'UID=' + sql_user_name + ';'
                            'PWD=' + sql_password + ';'
                            'Trusted_Connection=no;')
        return(0)
    except Exception as ex:
        print('SQL connection error')
        return(-1)

def main():
    global SQLconnection
    read_config_file()
    connectSQL()
    setup()
    login()
    emp_ids=set()
    cnt=0
    emp_flg_start=1
    rows=[]
    select_statement = "SELECT [Employee_ID],[Employee_Name],[Status],[IsUploaded]"
    select_statement = select_statement + "FROM [RoadSafe].[dbo].[LMS List];"

    cursor = SQLconnection.cursor()
    cursor.execute(select_statement)
    employee_details = cursor.fetchall()
    cursor.close()
    
    for i in range(2000,len(employee_details)):
        each_employee = employee_details[i]
        # emp_ids.add(each_employee[0])
        emp_name = each_employee[1].strip()
        # mapping = each_employee[3].strip()
        src_id = each_employee[0]
        dst_id = src_id
        isuploaded = each_employee[3]
        fld_name = emp_name+"("+src_id+")"
        # path = 'Y:\\All Downloaded Documents\\'
        path = 'S:\\Cycle 1\\LMS\\Active\\'
        directory_contents = os.listdir(path)
        emp_report = {}
        #print(isuploaded)
        # li=['100433',]
        if not isuploaded:
            ufiles = []
            onlyfiles = []
            emp_flg=0
            #print(fld_name)
            cat={}
            if fld_name in directory_contents:
                if os.path.isdir(path+fld_name):
                    print(fld_name)
                    emp_ids.add(dst_id)
                    select_statement1 = """SELECT [Employee_ID],[Employee_Name],[File_Name],[Mapping],[IsUploaded],[UploadedDateTime]
                                         FROM [RoadSafe].[dbo].[LMS Files] WHERE [Employee_ID] = ? ;"""
                    cursor = SQLconnection.cursor()
                    cursor.execute(select_statement1,(src_id))
                    files_details = cursor.fetchall()
                    cursor.close()
                    for each_file_db in files_details:
                        if str(each_file_db[4]).strip()=='None':
                            if each_file_db[3] not in ['Personal Goals','Practical Assessment','Documents']:
                                onlyfiles.append(each_file_db[2])
                                cat[each_file_db[2]]={}
                                cat[each_file_db[2]]['mapping']=each_file_db[3]
                            print(cat)

                    # files_with_size = [ (ufile, os.stat(os.path.join(path,fld_name, ufile)).st_size) for ufile in onlyfiles]
                    files_with_size = []
                    for ufile in onlyfiles:
                        print(ufile)
                        if "''"in ufile:
                            ufile=ufile.replace("''","'")
                        full_path = os.path.join(path, fld_name, ufile)
                        file_size = os.stat(os.path.join(path, fld_name, ufile)).st_size
                        files_with_size.append((ufile, file_size))
                    for ufile,usize in files_with_size:
                        if ufile=='readme.txt':
                            continue
                        ext = os.path.splitext(ufile)[1].lower()
                        if len(ext)<10 and ext!='.jfif' and ext!='.p' and ext!='' and ext!='.aspx' and ext!='.part' and ext!='.0' and ext!='. 9' and ext!='. 7' and ext!='.10' and ext!='.13' and ext!='.02' and ext!='.08' and ext!='. 1' and ext!='. 6' and ext!='. 2' and ext!='. 3' and ext!='. 8' and ext!='.24' and ext!='.04' and usize!=0:
                            ufiles.append(ufile)
                    
                print(ufiles)
                if not ufiles:
                    cursor = SQLconnection.cursor()
                    sql_update = """UPDATE [RoadSafe].[dbo].[LMS List] SET [IsUploaded] = ?
                        WHERE  [Employee_ID] =? ;"""
                    cursor.execute(sql_update,(1,dst_id))
                    SQLconnection.commit()
                    cursor.close()
                    print("No files for this employee and updated databse:",src_id)
                    continue
                res = searchAndUpload(ufiles, fld_name, dst_id,emp_name, cat,src_id)
                if res == 'No Cert':
                    cursor = SQLconnection.cursor()
                    sql_update = """UPDATE [RoadSafe].[dbo].[LMS List] SET [IsUploaded] = ?
                        WHERE  [Employee_ID] =? ;"""
                    cursor.execute(sql_update,(1,dst_id))
                    SQLconnection.commit()
                    cursor.close()
                    print(dst_id," : updated in Emp_List_157905")
                elif res == 0:
                    print(dst_id," : updated in Emp_List_157905")            
                elif res<0:
                    print('Employee error '+fld_name)
                    continue
main()
