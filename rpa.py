import os
import subprocess
import time
from selenium import webdriver

import pandas as pd
from datetime import datetime

from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait


FPATH='/Users/right/Desktop/geckodriver.exe'
SITE='https://ec.europa.eu/taxation_customs/vies/vatRequest.html'
FILE='/Users/right/Desktop/vat_codes.xlsx'
df = pd.read_excel(FILE)
df.columns=['VAT', 'Status', 'Date']
df['Country']= df['VAT'].astype(str).str[:2]
df['VATID']= df['VAT'].astype(str).str[2:]
print(df)

country=df['Country'].tolist()
vatid=df['VATID'].tolist()

# print(country)
# print(vatid)

def switch_demo(argument):
    switcher = {
        "DE": 6,
        "GB": 13,
        "SE": 26,
        "DK": 7,
        "CY": 4,

    }
    #print (switcher.get(argument, "Invalid month"))
    return switcher.get(argument,"Invalid month")

browser = webdriver.Firefox()
browser.get(SITE)
i=0
browser.maximize_window()

def mymain():
    global i
    try:
        for (c,v) in zip(country, vatid):
            print("\n" + c,v)
            element = WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.ID, "countryCombobox")))

            select_c = Select(browser.find_element_by_id("countryCombobox"))
            nr = switch_demo(c)

            if not nr == "Invalid month":
                select_c.select_by_index(nr)
                time.sleep(1)
                el = browser.find_element_by_id('number').send_keys(v)
                time.sleep(1)
                btn_submit= browser.find_element_by_id('submit').click()
                time.sleep(3)
                """Check the Writing"""
                message = browser.find_element_by_xpath("/html/body/div[2]/div[4]/div/div/div[2]/div/fieldset/table/tbody/tr[1]/td/b/span")
                mymessage = message.text
                print(mymessage)
                # df['Status'].loc[i]=mymessage
                # df['Date'].loc[i]=time.Now()
                df.loc[df.index[i], 'Status'] = mymessage
                now = datetime.now()
                df.loc[df.index[i], 'Date'] = now
                """GO BACK"""
                btn_back=  browser.find_element_by_xpath("/html/body/div[2]/div[4]/div/div/div[2]/div/fieldset/p/a")
                btn_back.click()
                time.sleep(2)
                i = i + 1
                #print(i)
            else:
                msg="Incorrect VAT ID, please check the ISO country code"
                print(msg)
                df.loc[df.index[i], 'Status'] = msg
                now = datetime.now()
                df['Date']=now
                i = i + 1

                #print(i)

        print("\n Please find bellow the Data Frame\n")
        df.columns=['VAT codes', 'Status', 'Date', 'Country', 'VATID']
        print(df)

        df.to_excel(FILE, columns=["VAT codes", "Status", "Date"],index=False)

    except NoSuchElementException:
         print("The page did not open for 3 minutes")
         df['Status']="The site is down"
         now = datetime.now()
         df['Date']=now

    finally:
         #browser.quit()
         browser.close()
         print("\n That was it!\n Thank you for the challenge!\n")
mymain()
