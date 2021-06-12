

from selenium import webdriver

from selenium.webdriver.chrome.webdriver import WebDriver

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
xlsx_path = (r'C:\Users\karthik.ramani\read_xlsx\Source\Test_copy.xlsx')

sheetDataFrame = pd.read_excel(xlsx_path, sheet_name)

print("Sleeping for 10 seconds")
time.sleep(10)

driver.get("https://wf5.myhcl.com/OpsHi5/SurveyPage_Wf.aspx?subSurId=22319")

#driver.get("https://wf5.myhcl.com/OpsHi5/index.aspx?PageName=Default.aspx")

print("Sleeping for 10 seconds")
time.sleep(10)

x = driver.find_elements_by_name("loginfmt")

x[0].send_keys("karthik.ramani@hcl.com")

print("Sleeping for 5 seconds")
time.sleep(5)

driver.find_elements_by_id("idSIButton9")[0].click()

print("Sleeping for 5 seconds")
time.sleep(5)

#y = driver.find_elements_by_id("i0118")

# print("Sleeping for 5 seconds")
# time.sleep(5)

#y[0].send_keys("")

# print("Sleeping for 10 seconds")
# time.sleep(5)

#driver.find_elements_by_id("idSIButton9")[0].click()

print("Sleeping for 30 seconds")
time.sleep(30)

# Insert in the form the Audit evidence value one single data



#for i in sheet.iter_rows(min_row=2, max_row=9, values_only=True):

# for i in sheet.range(1, 10):
i = 2
for row in sheetDataFrame.iterrows():

    i +=1

    prevsymbol = row["List of audit evidence/Remarks "].value

    if prevsymbol == None:

        pass

    else:

        try:

            print("Sleeping for 20 seconds")
            time.sleep(20)

            elementId = "grdView_ServiceDelivery_ctl{0:0=2d}_txtareaAuditEvidence".format(i)
            driver.find_element_by_id(elementId).clear()

            print("Sleeping for 20 seconds")
            time.sleep(20)

            # driver.find_element_by_id(elementId).send_keys(sheet["G2"].value)

            driver.find_element_by_id(elementId).send_keys(prevsymbol)

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

            print("List of audit evidence/Remarks :(txtareaAuditEvidence) Not Found")

        # Insert in the form the Observation deviation value one single data

    prevsymbol = row["Observation on deviation "].value

    if prevsymbol == None:

        pass

    else:

        try:

            print("Sleeping for 10 seconds")
            time.sleep(10)

            driver.find_element_by_id("grdView_ServiceDelivery_ctl02_txtareaObsDeviation").clear()

            print("Sleeping for 10 seconds")
            time.sleep(10)

            driver.find_element_by_id("grdView_ServiceDelivery_ctl02_txtareaObsDeviation").send_keys(row["Observation on deviation "].value)

        except NoSuchElementException:

            print("Observation on deviation :() txtareaObsDeviation Not Found")

    # click Save button to change the content

    print("Sleeping for 10 seconds")
    time.sleep(10)

driver.find_element_by_xpath("//*[@id='btn_Save_SD']").click()



def service_delivery_handler():

    pass



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
