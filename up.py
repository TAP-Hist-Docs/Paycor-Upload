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
        update_statement = """UPDATE [Six_Robblees].[dbo].[Emp_Files_Upload] SET [IsUploaded] = ?, [UploadedDateTime] = ? 
                        WHERE  [Employee_ID] =? AND [File_Name] = ? ;
                        """
        cursor.execute(update_statement,(1,time.strftime('%Y-%m-%d %H:%M:%S'),src_id,fname))
        SQLconnection.commit()   
        print(dst_id," : updated in Emp_Files_Upload")     
    except Exception as e:
        print('update to database error',e)
        return(-1)
    return 0

def searchAndUpload(ufiles, fld_name, lastname, dst_id,firstname, category,src_id):
    try:
        sea_flg = 0
        sea_cnt = 0
        while(1):
            try:
                sea_cnt = sea_cnt + 1
                try:
                    xpath = "//a[@href='//HCM.paycor.com/people/?companyId=214370&clientId=174248'][contains(text(),'Manage People')]" #user name
                    element = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
                    element.click()
                    time.sleep(3)
                except:
                    try:
                        xpath = "//a[@href='//HCM.paycor.com/people/?companyId=214370&clientId=174248'][contains(text(),'Manage People')]" #user name
                        element = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
                        element.click()
                        time.sleep(3)
                    except:
                        pass
                driver.get("https://hcm.paycor.com/people/?companyId=214370&clientId=174248")
                time.sleep(2)
                paycor_employee=dst_id
                print(paycor_employee)
                xpath  = "//input[contains(@id,'search_')]"
                driver.find_elements(By.XPATH, xpath).clear()
                element = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
                element.click()
                time.sleep(3)
                element.send_keys(paycor_employee)
                time.sleep(2)

                """Employee ID not Available in Paycor"""
                
                try:
                    element = driver.find_element(By.XPATH, "//span[@class='grapheinPro_LightItalic']")

                    # Extract and print the text
                    text = element.text
                    if text =="No people to display.":
                        with open('personal EmployeeNotInPaycor.csv', mode = 'a') as f:
                            writer = csv.writer(f)
                            writer.writerow([src_id, lastname, firstname])

                        return 0

                except:
                    print("Inside Exception")

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
            xpath = "//*[@id='3000'][@aria-label='Position'][@aria-expanded='false']"
            xpath = '//*[text()="Assignment"]'
            element = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
            element.click()
            time.sleep(3)
        except:
            pass

        try:
            # xpath = "/html/body/div[1]/div/div[2]/div/div/div[2]/nav/ul/nav[4]/ul/li[3]/a/span/span[2]"
            xpath = '//span[text() ="Documents"]'
            element = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
            element.click()
            time.sleep(3)
        except:
            try:
                xpath = "/html/body/div[1]/div/div[3]/div[2]/div/div[2]/nav/ul/nav[3]/ul/li[1]/a/span/span[2]"
                element = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
                element.click()
                time.sleep(3)
            except:
                pass
        try:
            xpath = "//div[@id='unfiled']//div[@class='style_documentFolder_name__2D6Yd']"
            # xpath = '/html/body/div[1]/div[2]/div[2]/div/div/div[2]/div/div[3]/div[3]/div/div[1]/svg[1]'
            element=driver.find_element(By.XPATH,xpath)
            element.click()
            time.sleep(2)
        except Exception as h:
            print("error h is:", h)
            try:
                xpath = "//div[@id='unfiled']//div[@class='style_documentFolder_name__2D6Yd']"
                element=driver.find_element(By.XPATH,xpath)
                element.click()
                time.sleep(2)
            except Exception as g:
                print("error g is:",g)
        xpath = "//*[contains(@class,'table-content_row__31sbr')]"
        no_of_documents = driver.find_elements(By.XPATH,xpath)
        initial_count = len(no_of_documents)
        print("Befor upload count:",initial_count)
        with open("personal Before upload count.csv",'a',encoding='utf-8') as f:
            f.write(f'"{fld_name}",{initial_count}\n')
        print(ufiles)
        
        for ufile in ufiles:
            down_flg = 0
            down_count = 0
            while(1):
                try:
                    # xpath = "//div[@id='unfiled']//div[@class='style_documentFolder_name__2D6Yd']"
                    # element=driver.find_element(By.XPATH,xpath)
                    # element.click()
                    # time.sleep(2)
                    xpath = "//*[contains(@class,'table-content_row__31sbr')]"
                    no_of_documents = driver.find_elements(By.XPATH,xpath)
                    initial_count = len(no_of_documents)
                    

                    down_count = down_count + 1
                    time.sleep(2)
                    # xpath='/html/body/div[1]/div[2]/div[2]/div/div[1]/button/div/span'
                    xpath='//button[@class="button_button__3-VDN button_button--primary__YkS-a"]'
                    element = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
                    element.click()
                    time.sleep(4)

                    xpath= "//input[contains(@id,'input_')]"
                    element = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
                    element.click()
                    f_name = os.path.splitext(ufile)[0]
                    element.send_keys(f_name)
                    time.sleep(3)
                    print(f_name)
                    emp_name=firstname+" "+lastname

                    xpath = "//*[@id='category']"
                    element = wait.until(EC.visibility_of_element_located((By.XPATH, xpath)))
                    element.click()
                    time.sleep(2)
                    # text = category[ufile]['catt']  #"AllEmployeeDocuments"
                    text = 'Personal'
                    print(text)
                    actions = ActionChains(driver)
                    actions.move_to_element(element)
                    actions.click()
                    time.sleep(2)
                    actions.send_keys(text)
                    actions.send_keys(Keys.RETURN)
                    actions.perform()
                    time.sleep(2)
                    xpath = "//*[@id='selectedFile']"
                    element  = driver.find_element(By.XPATH, xpath)
                    # file_path = r"Z:\Six Robblees\Document\Documents_Employment\"+fld_name+"\\"+ufile
                    file_path = rf"Z:\Six Robblees\Document\Documents_Personal\{fld_name}\{ufile}"
                    print(f"File Path = {file_path}")
                    # element.send_keys("Z:/Six Robblees/Document/Documents_Employment/"+fld_name+"/"+ufile)
                    try:

                        element.send_keys(file_path)
                        time.sleep(3)

                    except Exception as er:
                        print(f"\n\n Wrong File path = {er} - {fld_name}")
                
                    # element.send_keys("Z:\\Six Robblees\\Document\\Documents_Employment\\"+fld_name+"\\"+ufile)
                    # time.sleep(2)

                    # xpath = "/html/body/div[1]/div[2]/div[2]/form/div/div[7]/div/div/div/div/div[2]/label/div[1]"
                    # xpath = '//*[@id="boxEmployeePublicDocument"]'
                    # element = wait.until(EC.visibility_of_element_located((By.XPATH, xpath)))
                    # element.click()
                    # time.sleep(2)


                    """Employee Visiblilty Radio Button"""
                    try:
                        xpath = '//*[@id="boxEmployeeDocument_Administrator"]/svg/g/path'
                        element  = driver.find_element(By.XPATH, xpath)
                        element.click()
                        time.sleep(2)


                    except Exception as e:
                        print(f"\n\n Employee visibility Error = {e}")
                    
                    
                    """Click on Save"""
                    try:
                        xpath = "//html//body//div[1]//div[2]//div[2]//form//div//div[8]//div//div//div//button[2]"
                        xpath = "//button[@type='submit']"
                        element = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
                        element.click()
                        time.sleep(5)
                    except:
                        xpath = "//html//body//div[6]//div[2]//div[2]//button[2][contains(text(),'Save')]"
                        element = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
                        element.click()
                        time.sleep(5)
                        
                    value = dst_id+","+emp_name+","+","+ufile
                    print(value)
                            
                    f= open('personal uploaded.csv',"a")
                    f.write(f'{dst_id},"{emp_name}","{ufile}"')
                    f.write('\n')
                    f.close()
                    res = updateToDatabase(src_id,dst_id,emp_name,ufile)
                    driver.refresh()
                    time.sleep(5)
                    xpath = "//div[@id='unfiled']//div[@class='style_documentFolder_name__2D6Yd']"
                    element=driver.find_element(By.XPATH,xpath)
                    element.click()
                    time.sleep(2)
                    xpath = "//*[contains(@class,'table-content_row__31sbr')]"
                    no_of_documents = driver.find_elements(By.XPATH,xpath)
                    final_count = len(no_of_documents)
                    print("After upload count:",final_count,'\nBefore upload count:',initial_count)
                    
                    filename = "personal report.txt"
                    with open(filename, "a+", encoding="utf-8") as file:
                        if res == 0 and final_count == initial_count + 1:
                            file.write(dst_id+", Successfully updated the database and uploaded to paycor " + ufile + "\n")
                        elif res == 0 and final_count != initial_count + 1:
                            file.write(dst_id+", Successfully updated the database but not uploaded to paycor " + ufile + "\n")
                        elif res < 0 and final_count == initial_count + 1:
                            file.write(dst_id+", Not updated the database but uploaded to paycor " + ufile + "\n")
                        else:
                            file.write(dst_id+", Not updated the database and not uploaded to paycor " + ufile + "\n")
                    # xpath = "//div[@id='unfiled']//div[@class='style_documentFolder_name__2D6Yd']"
                    # element=driver.find_element(By.XPATH,xpath)
                    # element.click()
                    # time.sleep(2)

                    if res<0:
                        return(-1)
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
        f= open('personal employee name error.csv',"a")
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
        xpath = "/html/body/div[1]/paycor-employee-navigation-bar/quick-links/button"
        element = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
        element.click()
        time.sleep(3)
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
    select_statement = "SELECT [Employee_ID],[Last_Name],[First_Name],[IsUploaded_personal]"
    select_statement = select_statement + "FROM [Six_Robblees].[dbo].[Emp_list_Upload];"

    cursor = SQLconnection.cursor()
    cursor.execute(select_statement)
    employee_details = cursor.fetchall()
    cursor.close()
    
    for i in range(270,len(employee_details)):
        each_employee = employee_details[i]
        # emp_ids.add(each_employee[0])
        firstname = each_employee[2].strip()
        lastname = each_employee[1].strip()
        src_id = each_employee[0]
        # dst_id = each_employee[3]
        dst_id = src_id
        isuploaded = each_employee[3]
        fld_name = lastname+" "+firstname+" ("+src_id+")"
        # path = 'Y:\\All Downloaded Documents\\'
        path = 'Z:\\Six Robblees\\Document\\Documents_Personal\\'
        directory_contents = os.listdir(path)
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
                    select_statement1 = """SELECT [Employee_Id],[Employee_Name],[File_Name],[IsUploaded],[UploadedDateTime]
                                         FROM [Six_Robblees].[dbo].[Emp_Files_Upload_Personal] WHERE [Employee_Id] = ? ;"""
                    cursor = SQLconnection.cursor()
                    cursor.execute(select_statement1,(src_id))
                    files_details = cursor.fetchall()
                    cursor.close()
                    for each_file_db in files_details:
                        if str(each_file_db[3]).strip()=='None':
                            # if each_file_db[2]=='Paystubs':
                                onlyfiles.append(each_file_db[2])
                                # cat[each_file_db[3]]={}
                                # cat[each_file_db[3]]['catt']=each_file_db[2]

                    # files_with_size = [ (ufile, os.stat(os.path.join(path,fld_name, ufile)).st_size) for ufile in onlyfiles]
                    files_with_size = []
                    print(onlyfiles)
                    for ufile in onlyfiles:
                        #print(ufile)
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
                res = searchAndUpload(ufiles, fld_name, lastname, dst_id, firstname, cat,src_id)
                if res == 0:
                    cursor = SQLconnection.cursor()
                    sql_update = """UPDATE  [Six_Robblees].[dbo].[Emp_list_Upload] SET [IsUploaded_personal] = ?
                        WHERE  [Employee_ID] =? ;"""
                    cursor.execute(sql_update,(1,dst_id))
                    SQLconnection.commit()
                    cursor.close()
                    print(dst_id," : updated in Emp_list_Upload")
    
                elif res<0:
                    print('Employee error '+fld_name)
                    continue
main()