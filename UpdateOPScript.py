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

sleep(10)

driver.get("https://wf5.myhcl.com/OpsHi5/SurveyPage_Wf.aspx?subSurId=22319")

#driver.get("https://wf5.myhcl.com/OpsHi5/index.aspx?PageName=Default.aspx")

sleep(30)

x = driver.find_elements_by_name("loginfmt")

x[0].send_keys("karthik.ramani@hcl.com")

sleep(10)

driver.find_elements_by_id("idSIButton9")[0].click()

sleep(5)

#y = driver.find_elements_by_id("i0118")

sleep(5)

#y[0].send_keys("")

# sleep(5)

#driver.find_elements_by_id("idSIButton9")[0].click()

sleep(60)

# Insert in the form the Audit evidence value one single data



#for i in sheet.iter_rows(min_row=2, max_row=9, values_only=True):

# for i in sheet.range(1, 10):
start = time.time()
i = 2
for index, row in sheetDataFrame.iterrows():
    print(row['List of audit evidence/Remarks'])
    prevsymbol = row["List of audit evidence/Remarks"]
    #i +=1
    if prevsymbol is not None:
        try:
            sleep(10)
            elementId = "grdView_ServiceDelivery_ctl{0:0=2d}_txtareaAuditEvidence".format(i)
            driver.find_element_by_id(elementId).clear()
            sleep(10)
            driver.find_element_by_id(elementId).send_keys(prevsymbol)
        except NoSuchElementException:
            print("List of audit evidence/Remarks :(txtareaAuditEvidence) Not Found")

        # Insert in the form the Observation deviation value one single data
    print(row['Observation on deviation'])
    prevsymbol1 = row["Observation on deviation"]

    if prevsymbol1 is not None:
        try:
            sleep(10)
            elementobsId = "grdView_ServiceDelivery_ctl{0:0=2d}_txtareaObsDeviation".format(i)
            driver.find_element_by_id(elementobsId).clear()
            sleep(10)
            driver.find_element_by_id(elementobsId).send_keys(prevsymbol1)
        except NoSuchElementException:

            print("Observation on deviation :() txtareaObsDeviation Not Found")

    # Insert in the form the Recommendation value one single data
    print(row['Recommendation'])
    prevsymbol2 = row["Recommendation"]

    if prevsymbol2 is not None:
        try:
            sleep(10)
            elementrecId = "grdView_ServiceDelivery_ctl{0:0=2d}_txtareaActionPlan".format(i)
            driver.find_element_by_id(elementrecId).clear()
            sleep(10)
            driver.find_element_by_id(elementrecId).send_keys(prevsymbol2)
        except NoSuchElementException:
            print("Recommendation :() txtareaObsDeviation Not Found")

    # click Save button to change the content
    #driver.find_element_by_xpath("//*[@id='btn_Save_SD']").click()

    # We are setting the dropdwn to Green/Partial on every iteration (which means on every row in the table.)
    # prevsymbol => will never be empty or "-"
    assert len(prevsymbol) and len(prevsymbol1) and len(prevsymbol2), "Invalid Data"

    dropdownId = "grdView_ServiceDelivery_ctl{0:0=2d}_ddl_AuditorStatus".format(i)
    if not prevsymbol1:
        # set it to Green
        el = driver.find_element_by_id(dropdownId)
        for option in el.find_elements_by_tag_name('option'):
            if option.text == 'Done':
                option.click()
                break
        # driver.find_element_by_xpath("//select[@id='" +dropdownId+ "']/option[text()='Done']").click()
    else:
        if len(prevsymbol1) == 1 and prevsymbol1[0] == "-":
            # set it to Green
            el = driver.find_element_by_id(dropdownId)
            for option in el.find_elements_by_tag_name('option'):
                if option.text == 'Done':
                    option.click()
                    break
            # driver.find_element_by_xpath("//select[@id='" +dropdownId+ "']/option[text()='Done']").click()
        else:
            # set it to Partial
            el = driver.find_element_by_id(dropdownId)
            for option in el.find_elements_by_tag_name('option'):
                if option.text == 'Partial':
                    option.click()
                    break
            # driver.find_element_by_xpath("//select[@id='" +dropdownId+ "']/option[text()='Partial']").click()

    targetDateId = "grdView_ServiceDelivery_ctl{0:0=2d}_txtTargetDate".format(i)
    from datetime import datetime
    from dateutil import relativedelta
    now = datetime.now()
    future_date_from_now = now + relativedelta.relativedelta(months=1)
    future_date_with_15th_day = datetime(year=future_date_from_now.year, month=future_date_from_now.month, day=15)
    driver.find_element_by_id(targetDateId).send_keys(future_date_with_15th_day.strftime("%d %b %Y"))

    i +=1
    sleep(10)
    end=time.time()
    hours, rem = divmod(end-start, 3600)
    minutes, seconds = divmod(rem, 60)
    print("{:0>2}:{:0>2}:{:05.2f}".format(int(hours),int(minutes),seconds))
    #print(end-start)
    print("Moving to the next row...")


def service_delivery_handler():
    pass

#def AuditorStatus(prevsymbol, prevsymbol1):
    #if prevsymbol and prevsymbol1 == 0
    #  driver.find_element_by_xpath("//select[@name='element_name']/option[text()='option_text']").click()
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


def sleep(seconds):
    print("Sleeping for %s seconds" % seconds)
    time.sleep(seconds)
