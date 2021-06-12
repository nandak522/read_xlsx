from selenium import webdriver
 
from selenium.webdriver.chrome.webdriver import WebDriver

#from selenium.webdriver.common.keys import keys
 
from selenium.common.exceptions import NoSuchElementException
 
import xlrd
 
import openpyxl
 
import pandas as pd
 
import time
 
 
 
options = webdriver.ChromeOptions()
 
options.add_experimental_option("excludeSwitches", ["enable-logging"])
 
driver = webdriver.Chrome(options=options, executable_path=r'C:\Users\karthik.ramani\Downloads\New_Install\chromedriver.exe')
 
driver.maximize_window()
 
sheet_name = "Sheet1"
xlsx_path = (r'C:\Users\karthik.ramani\read_xlsx\Source\Techem.xlsx')
 
sheetDataFrame = pd.read_excel(xlsx_path, sheet_name)
 
print("Sleeping for 10 seconds")
time.sleep(20)
 
driver.get("https://wf5.myhcl.com/OpsHi5/SurveyPage_Wf.aspx?subSurId=22319")
 
#driver.get("https://wf5.myhcl.com/OpsHi5/index.aspx?PageName=Default.aspx")
 
print("Sleeping for 10 seconds")
time.sleep(30)
 
x = driver.find_elements_by_name("loginfmt")
 
x[0].send_keys("karthik.ramani@hcl.com")
 
print("Sleeping for 5 seconds")
time.sleep(10)
 
driver.find_elements_by_id("idSIButton9")[0].click()
 
print("Sleeping for 5 seconds")
time.sleep(10)
 
#y = driver.find_elements_by_id("i0118")
 
# print("Sleeping for 5 seconds")
# time.sleep(5)
 
#y[0].send_keys("")
 
# print("Sleeping for 10 seconds")
# time.sleep(5)
 
#driver.find_elements_by_id("idSIButton9")[0].click()
 
print("Sleeping for 60 seconds")
time.sleep(40)
 
# Insert in the form the Audit evidence value one single data
 
 
 
#for i in sheet.iter_rows(min_row=2, max_row=9, values_only=True):
 
# for i in sheet.range(1, 10):
start = time.time()
i = 2
for index, row in sheetDataFrame.iterrows():
    print(row['List of audit evidence/Remarks'])
    prevsymbol = row["List of audit evidence/Remarks"]
    #i +=1
    if prevsymbol == None:
 
        pass
 
    else:
 
        try:
            
            print("Sleeping for 20 seconds Audit Evidence Data will be copied now..")
            time.sleep(10)
            elementId = "grdView_ServiceDelivery_ctl{0:0=2d}_txtareaAuditEvidence".format(i)
            driver.find_element_by_id(elementId).clear()
            #driver.find_element_by_id("grdView_ServiceDelivery_ct102_txtareaAuditEvidence").clear()
            
            print("Sleeping for 20 seconds")
            
            time.sleep(10)
 
            # driver.find_element_by_id("grdView_ServiceDelivery_ctl02_txtareaAuditEvidence").send_keys(sheet["G2"].value)
            
            #if elementId != str('NA'):
            driver.find_element_by_id(elementId).send_keys(prevsymbol)
            

            #fillna(NA, inplace=True)
            #######
 
            # categories_map = {
 
            #     "Service Delivery": service_delivery_handler,
 
            #     "People": people_handler,
 
            # }
 
            # for key in categories_map:
 
            #     func_name = categories_map[key]
 
            #     func_name()
 
            #######
 
        except NoSuchElementException:
 
            print("List of audit evidence/Remarks :(txtareaAuditEvidence) Not Found")
 
        # Insert in the form the Observation deviation value one single data
    print(row['Observation on deviation'])
    prevsymbol1 = row["Observation on deviation"]
 
    if prevsymbol1 == None:
 
        pass
 
    else:
 
        try:
 
            print("Sleeping for 10 seconds")
            time.sleep(10)
            elementobsId = "grdView_ServiceDelivery_ctl{0:0=2d}_txtareaObsDeviation".format(i)
            driver.find_element_by_id(elementobsId).clear()
            #driver.find_element_by_id("grdView_ServiceDelivery_ctl02_txtareaObsDeviation").clear()
 
            print("Sleeping for 10 seconds")
            time.sleep(10)
 
            #driver.find_element_by_id("grdView_ServiceDelivery_ctl02_txtareaObsDeviation").send_keys(prevsymbol1)
            driver.find_element_by_id(elementobsId).send_keys(prevsymbol1)
            
 
        except NoSuchElementException:
 
            print("Observation on deviation :() txtareaObsDeviation Not Found")
            
            
     # Insert in the form the Recommendation value one single data 
    print(row['Recommendation'])
    prevsymbol2 = row["Recommendation"]
    
    if prevsymbol2 == None:
 
        pass
 
    else:
 
        try:
 
            print("Sleeping for 10 seconds")
            time.sleep(10)
            elementrecId = "grdView_ServiceDelivery_ctl{0:0=2d}_txtareaActionPlan".format(i)
            driver.find_element_by_id(elementrecId).clear()
            #driver.find_element_by_id("grdView_ServiceDelivery_ctl02_txtareaObsDeviation").clear()
 
            print("Sleeping for 10 seconds")
            time.sleep(10)
 
            #driver.find_element_by_id("grdView_ServiceDelivery_ctl02_txtareaObsDeviation").send_keys(prevsymbol2)
            driver.find_element_by_id(elementrecId).send_keys(prevsymbol2)
 
        except NoSuchElementException:
 
            print("Observation on deviation :() txtareaObsDeviation Not Found")
 
        #grdView_ServiceDelivery_ctl03_txtTargetDate
        #grdView_ServiceDelivery_ctl02_ddl_AuditorStatus
 
    # click Save button to change the content
        #driver.find_element_by_xpath("//*[@id='btn_Save_SD']").click()
        i +=1 
        print("Sleeping for 10 seconds")
        time.sleep(10)
        end=time.time()
        hours, rem = divmod(end-start, 3600)
        minutes, seconds = divmod(rem, 60)
        print("{:0>2}:{:0>2}:{:05.2f}".format(int(hours),int(minutes),seconds))
        #print(end-start)
 
 
def service_delivery_handler():
 
    pass
 

#def AuditorStatus(prevsymbol, prevsymbol1):
    #if prevsymbol and prevsymbol1 == 0
     driver.find_element_by_xpath("//select[@name='element_name']/option[text()='option_text']").click() 
        #return;
 
 
def retrieve_column_value(column_name):
 
    excel_data_df = pd.read_excel(
 
        xlsx_path,
 
        sheet_name=sheet_name,
 
        usecols=[
 
            'Category ',
 
            'S.No ',
 
            'Parameters ',
 
            'Category ',
 
            'Filter ',
 
            'Guidelines ',
 
            'List of audit evidence/Remarks ',
 
            'Observation on deviation ',
 
            'Recommendation ']
 
    )
 
    # To get the column header names
 
    # print(excel_data_df.columns.ravel())
 
    return excel_data_df[[column_name]]
