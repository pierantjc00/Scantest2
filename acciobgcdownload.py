from logging import error
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
import csv
import time
import sys
import traceback
from datetime import date
from datetime import datetime, timedelta
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

# Red Arrow Online credentials
username = "jpierantozzi"
password = "Accio9294!"
accountname = "Foley"
rootdir = "C:\\DW\\ACCIO\\"
def write_logs(error_message, stack_trace):
    # print the error message and stack
    #print (error_message + '\n' + stack_trace)
    with open(rootdir + 'acciodownloaderrors.log', 'a') as outfile:
        outfile.write(error_message + '\n' + stack_trace)

def write_log_message(error_message):
    # print the error message and stack
    #print (error_message + '\n' + stack_trace)
    with open(rootdir + 'acciodownloaderrors.log', 'a') as outfile:
        outfile.write((datetime.now()).strftime("%m/%d/%Y %H:%M:%S") + ' - ' + error_message + '\n')

def write_table_to_file(pAppendFlag, pCorrectionDate, pHTMLTable):
    fileWriteParameter = 'w'
    if (pAppendFlag):
        fileWriteParameter = 'a'

    try:
        with open(rootdir + datetime.now().strftime("%Y%m%d") + 'CorrectionSummary' + '.csv', fileWriteParameter, newline='') as csvfile:
            outrow = []
            wr = csv.writer(csvfile)
            #only write header row if overwriting file, or first row of appending
            if not (pAppendFlag):
                outrow.append('PersonFullName')    
                outrow.append('CorrectionDate')
                outrow.append('NumberOfCorrections')     
                outrow.append('WeekNumber')
                wr.writerow(outrow)
            outrow = []
            for row in pHTMLTable.find_elements(By.TAG_NAME, "tr"):
                cellCounter = 0
                outrow = []
                for d in row.find_elements(By.TAG_NAME, "td"):
                    cellCounter += 1
                    if (cellCounter == 1):
                        outrow.append(d.text)    
                    if (cellCounter == 2):
                        outrow.append(pCorrectionDate.strftime("%m/%d/%Y"))    
                    if (cellCounter == 5):
                        outrow.append(d.text)
                if (len(outrow) > 0):
                    if (outrow[0] != 'Totals:'):        
                        #outrow.append(pCorrectionDate.isocalendar().week)
                        outrow.append(pCorrectionDate.strftime("%U"))
                        wr.writerow(outrow)
    except Exception as e:
        error_message = str(e)
        stack_trace = str(traceback.format_exc())
        write_logs(error_message, stack_trace)
        driver.close()
        exit("Unable to create export CorrectionSummary.csv file")

    write_log_message('After write file')

def export_data(pfromDate, ptoDate):
    noDataForDate = 0
    headerWritten = 0
    write_log_message('In export data')    

    if (pfromDate == ptoDate):
        try:
            #Find the From Date field on correction summary report screen and populate with current date
            #fromDate = driver.find_element(By.XPATH, ("/html/body/div[3]/div/form/input[1]"));    
            
            fromDateElement = driver.find_element('xpath', "//input[@name='dtFrom']");
            fromDateElement.clear()
            fromDateElement.send_keys(pfromDate.strftime("%m/%d/%Y"));

            #Find the To Date field on correction summary report screen and populate with current date
            #toDate = driver.find_element(By.XPATH, ("/html/body/div[3]/div/form/input[2]"));
            toDateElement = driver.find_element('xpath', "//input[@name='dtTo']");
            toDateElement.clear()
            toDateElement.send_keys(ptoDate.strftime("%m/%d/%Y"));
        except Exception as e:
            error_message = str(e)
            stack_trace = str(traceback.format_exc())
            write_logs(error_message, stack_trace)
            driver.close()
            exit("Unable to complete setting To and From date")

        write_log_message('After set to and from date')
        try:
            #Find Refresh button and click
            #refreshButton = driver.find_element(By.XPATH, ("/html/body/div[3]/div/form/input[3]"));
            refreshButton = driver.find_element('xpath', "//input[@type='submit']");
            refreshButton.submit()
        except Exception as e:
            error_message = str(e)
            stack_trace = str(traceback.format_exc())
            write_logs(error_message, stack_trace)
            driver.close()
            exit("Unable to complete reloading report with current To/From Date")

        #wait for screen to refresh
        time.sleep(3)

        try:
            #Find table and write to CSV
            #table = driver.find_element(By.XPATH, ("/html/body/div[6]/table"));
            table = driver.find_element('xpath', "//table");
            #print(table)
            time.sleep(10)
        except Exception as e:
            error_message = str(e)
            stack_trace = str(traceback.format_exc())
            write_logs(error_message, stack_trace)
            driver.close()
            exit("Unable to find table on page")
        
        write_table_to_file(False, pfromDate,table)
    else:

        write_log_message('In process multiple days')

        for i in range(3):
            d = pfromDate+timedelta(days=i) #+ timedelta(days=x) for x in range((ptoDate-pfromDate).days + 1)]:
            write_log_message('Processing date: ' + d.strftime("%m/%d/%Y"))
            try:
                #Find the From Date field on correction summary report screen and populate with current date
                fromDateElement = driver.find_element('xpath', "//input[@name='dtFrom']");
                fromDateElement.clear()
                fromDateElement.send_keys(d.strftime("%m/%d/%Y"));

                #Find the To Date field on correction summary report screen and populate with current date
                toDateElement = driver.find_element('xpath', "//input[@name='dtTo']");
                toDateElement.clear()
                toDateElement.send_keys(d.strftime("%m/%d/%Y"));
            except Exception as e:
                error_message = str(e)
                stack_trace = str(traceback.format_exc())
                write_logs(error_message, stack_trace)
                driver.close()
                exit("Unable to complete setting To and From date")

            write_log_message('After set to and from date')
            try:
                #Find Refresh button and click
                #refreshButton = driver.find_element(By.XPATH, ("/html/body/div[3]/div/form/input[3]"));
                refreshButton = driver.find_element('xpath', "//input[@type='submit']");
                refreshButton.submit()
            except Exception as e:
                error_message = str(e)
                stack_trace = str(traceback.format_exc())
                write_logs(error_message, stack_trace)
                driver.close()
                exit("Unable to complete reloading report with current To/From Date")

            #wait for screen to refresh
            time.sleep(3)

            try:
                #Find table and write to CSV
                #table = driver.find_element(By.XPATH, ("/html/body/div[6]/table"));
                table = driver.find_element('xpath', "//table");
                #print(table)
                time.sleep(10)
            except Exception as e:
                error_message = str(e)
                stack_trace = str(traceback.format_exc())
                write_logs(error_message, stack_trace)
                if ('no such element' in error_message ):
                    noDataForDate = 1
                else:
                    driver.close()
                    exit("Unable to find table on page")

            if (noDataForDate != 1):
                if (headerWritten == 0):
                    write_table_to_file(False, d,table)
                    headerWritten = 1
                else:
                    write_table_to_file(True, d,table)
            
            noDataForDate = 0

# Set log location
chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument('--no-proxy-server')
chrome_options.add_argument('service_args=["--verbose", "--log-path=C:\\dw\\accio\\acciodownloadchromeerrors.log"')
driver = webdriver.Chrome(chrome_options)

# initialize the Chrome driver and navigate to home page
try:
    driver.get("https://foleyservices.acciodata.com/sysops/sysop_home5.html")
except Exception as e:
    error_message = str(e)
    stack_trace = str(traceback.format_exc())
    write_logs(error_message, stack_trace)
    driver.close()
    exit("Not able to reach url")

write_log_message('After get login page')    
try:
    #Get User ID field and populate
    element = driver.find_element('xpath', "//input[@id='account']")
    element.send_keys(accountname)

    #Get Password field and populate
    element = driver.find_element('xpath', "//input[@id='userid']")
    element.send_keys(username)

    #Get Password field and populate
    element = driver.find_element('xpath', "//input[@id='password']")
    element.send_keys(password)
    
    time.sleep(3)
    
    #Submit form to login
    element = driver.find_element('xpath', '//*[@id="login_button"]')
    #element = driver.find_element('xpath','//button[text()="Sign In"]')
    element.click()
except Exception as e:
    error_message = str(e)
    stack_trace = str(traceback.format_exc())
    write_logs(error_message, stack_trace)
    driver.close()
    exit("Not able to reach url")

write_log_message('After login')

try:
    time.sleep(10)
    driver.get("https://foleyservices.acciodata.com/sysops/create_researchercompletion_report.html")
    WebDriverWait(driver=driver, timeout=10).until(
        lambda x: x.execute_script("return document.readyState === 'complete'")
    )
except Exception as e:
    error_message = str(e)
    stack_trace = str(traceback.format_exc())
    write_logs(error_message, stack_trace)
    driver.close()
    exit("Not able to reach url")

write_log_message('After navigate to export page')
time.sleep(10)

toDate = (datetime.now()+timedelta(days=-1)) #.strftime("%m/%d/%Y")
fromDate = toDate+timedelta(days=-3)
if (datetime.now().weekday() == 0):
    fromDate = (datetime.now()+timedelta(days=-3)) #.strftime("%m/%d/%Y");

usernamelist = "eromano, malma, jwright, adilisio, lgrant, cjones, ggotora, ptrueland, alandes, towenby, acollins, gperaza"

try:
            
    fromDateElement = driver.find_element('xpath', "//input[@id='fromdate']")
    #fromDateElement.clear()
    fromDateElement.send_keys(fromDate.strftime("%m/%d/%Y"));

    toDateElement = driver.find_element('xpath', "//input[@id='todate']")
    #toDateElement.clear()
    toDateElement.send_keys(toDate.strftime("%m/%d/%Y"));
    
    #usernameListElement = driver.find_element('xpath', "//input[@name='username']")
    usernameListElement = driver.find_element('xpath', '//*[@id="userNameRow"]/td[2]/div/input')
    #usernameListElement.clear()
    usernameListElement.send_keys(usernamelist);

    accountListElement = driver.find_element('xpath', '//*[@id="accountNumbersRow"]/td[2]/div/input')
    accountListElement.send_keys('Foley');

    time.sleep(10)
        
    #Submit form to generate report
    createReportButton = driver.find_element('xpath', '//*[@id="headerForm"]/input[1]');
    createReportButton.click()
    
    time.sleep(10)

    #Submit form to login
    downloadReportButton = driver.find_element('xpath', '//*[@id="resultsdivid"]/div[3]/div[1]/div/button');
    downloadReportButton.click()

    time.sleep(10)
    
except Exception as e:
    error_message = str(e)
    stack_trace = str(traceback.format_exc())
    write_logs(error_message, stack_trace)
    driver.close()
    exit("Not able to reach url")

#write_log_message('After researcher')

try:
    time.sleep(10)
    driver.get("https://foleyservices.acciodata.com/sysops/create_reviewercompletion_report.html")
    WebDriverWait(driver=driver, timeout=10).until(
        lambda x: x.execute_script("return document.readyState === 'complete'")
    )
except Exception as e:
    error_message = str(e)
    stack_trace = str(traceback.format_exc())
    write_logs(error_message, stack_trace)
    driver.close()
    exit("Not able to reach url")

write_log_message('After navigate to export page')
time.sleep(10)

toDate = (datetime.now()+timedelta(days=-1)) #.strftime("%m/%d/%Y")
fromDate = toDate+timedelta(days=-3)
if (datetime.now().weekday() == 0):
    fromDate = (datetime.now()+timedelta(days=-3)) #.strftime("%m/%d/%Y");

usernamelist = "eromano, malma, jwright, adilisio, lgrant, cjones, ggotora, ptrueland, alandes, towenby, acollins, gperaza"

try:
            
    fromDateElement = driver.find_element('xpath', "//input[@id='fromdate']")
    #fromDateElement.clear()
    fromDateElement.send_keys(fromDate.strftime("%m/%d/%Y"));

    toDateElement = driver.find_element('xpath', "//input[@id='todate']")
    #toDateElement.clear()
    toDateElement.send_keys(toDate.strftime("%m/%d/%Y"));
    
    #usernameListElement = driver.find_element('xpath', "//input[@name='username']")
    usernameListElement = driver.find_element('xpath', '//*[@id="userNameRow"]/td[2]/div/input')
    #usernameListElement.clear()
    usernameListElement.send_keys(usernamelist);

    accountListElement = driver.find_element('xpath', '//*[@id="accountNumbersRow"]/td[2]/div/input')
    accountListElement.send_keys('Foley');

    time.sleep(10)
        
    #Submit form to generate report
    createReportButton = driver.find_element('xpath', '//*[@id="headerForm"]/input[1]');
    createReportButton.click()
    
    time.sleep(10)

    #Submit form to login
    downloadReportButton = driver.find_element('xpath', '//*[@id="resultsdivid"]/div[3]/div[1]/div/button');
    downloadReportButton.click()

    time.sleep(10)
    
except Exception as e:
    error_message = str(e)
    stack_trace = str(traceback.format_exc())
    write_logs(error_message, stack_trace)
    driver.close()
    exit("Not able to reach url")

#export_data(fromDate, toDate)
write_log_message('Done')

# close the driver
driver.close()
exit(0)