import pandas as pd
import numpy as np
import os
import re
import time
import warnings
warnings.filterwarnings('ignore')
from PIL import Image
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import NoAlertPresentException,TimeoutException
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.options import Options
from datetime import date
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.alert import Alert
import pygetwindow as gw
import pyautogui
import matplotlib.pyplot as plt
import math
import streamlit as st
import streamlit_option_menu
from streamlit_option_menu import option_menu


st._config.set_option('themebase','dark')

st.set_page_config(layout='wide')

title_txt = '''<h1 style='font-size : 55px;text-align:center;color:purple;background-color:lightgrey;'>BHIMA JEWELLERY</h1>'''
st.markdown(title_txt,unsafe_allow_html=True)

st.write(" ")


def init_driver():
    chrome_options = Options()
    chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    chrome_options.add_experimental_option('useAutomationExtension', False)    
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    driver = webdriver.Chrome(service=Service(
        r"C:\Users\DELL\Documents\eway\chromedriver-win64\chromedriver.exe",
        log_path="NUL"), 
        options=chrome_options)
    return driver

if st.button(":rainbow[Eway-Bill]",use_container_width=True):
    try:
        if "driver" not in st.session_state:
            st.session_state.driver = init_driver()
            driver = st.session_state.driver
            driver.get("https://ewaybillgst.gov.in")
            wait = WebDriverWait(driver, 20)
    
    
        # login page
        login_link = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Login")))
        login_link.click()

        user_name_input = wait.until(EC.presence_of_element_located((By.ID, "txt_username")))
        password_input = driver.find_element(By.ID,'txt_password')
        user_name_input.send_keys("BHIMAMADURAI")
        password_input.send_keys("Eway@1234")

        # captcha
        captcha_img = wait.until(EC.presence_of_element_located((By.XPATH, "//img[contains(@src, 'Captcha')]")))
        captcha_img.screenshot("captcha.png")
        st.image(Image.open("captcha.png"), caption="Captcha")

        # show input box for captcha
        st.session_state.captcha_value = ""
    except:
        st.info('Please Try again - Chrome Issue')

# callback for captcha submit
def submit_captcha():
                
    driver = st.session_state.driver
    wait = WebDriverWait(driver, 20)
    captcha_field = wait.until(EC.presence_of_element_located((By.ID, "txtCaptcha")))
    captcha_field.clear()
    captcha_field.send_keys(st.session_state.captcha_value)
    driver.find_element(By.ID, "btnLogin").click()
    try:
        WebDriverWait(driver,3).until(EC.alert_is_present())
        alert_captcha = driver.switch_to.alert
        msg =alert_captcha.text
        alert_captcha.accept()
        if "Invalid Captcha" in msg:
            st.error(f"Invalid Captcha ❌ ({msg})")
        elif "OTP has been sent" in msg:
            st.success(f"Captcha submitted: {st.session_state.captcha_value}")
        else:
            st.info(f"Alert:{msg}")
    except:
        st.success(f"Captcha submitted: {st.session_state.captcha_value}")    
        

def submit_otp():
    driver = st.session_state.driver
    wait = WebDriverWait(driver, 20)
    try:
        WebDriverWait(driver, 5).until(EC.alert_is_present())
        alert = driver.switch_to.alert
        st.info(f"Alert: {alert.text}")
        alert.accept()
        time.sleep(2)
    except TimeoutException:
        st.info("No initial alert after requesting OTP")
    otp_field = wait.until(EC.presence_of_element_located((By.ID, "OtpTxt")))
    otp_field.clear()
    otp_field.send_keys(st.session_state.otp_value)
    driver.find_element(By.ID, "btnsubmit").click()
    try:
        WebDriverWait(driver,5).until(EC.alert_is_present())
        alerts = driver.switch_to.alert
        mssg= alerts.text
        alerts.accept()
        if "Invalid " in mssg:
            st.error(f"Invalid Captcha ❌ ({mssg})")
        else:
            st.info(f"Alert:{mssg}")
    except TimeoutException:
        pass               
        

    try:
        ewaybill_menu = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "e-Waybill")))
        ewaybill_menu.click()
        st.success("Navigated to e-Waybill menu")
    

        ewb_gold_link = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "EWB for Gold")))
        ewb_gold_link.click()
        st.success("EWB for Gold page opened ✅")
    except:
        st.info('Please try again')


if "driver" in st.session_state:
    st.text_input("Enter Captcha", key="captcha_value", on_change=submit_captcha)
    otp_val = st.text_input("Enter OTP", key="otp_value", on_change=submit_otp)

    try:
        file = st.file_uploader ("Upload the file",type=['xlsx','xlsm','xls'])
    except:
        st.info('Please Upload the Correct File')

    if 'data_d' not in st.session_state:

        st.session_state['data_d'] = {}

    if st.button(":rainbow[Submit]") and file is not None:
        file_path = file
        xls = pd.ExcelFile(file_path)
        st.session_state.file_path = file_path
        st.session_state.sheets_lst = xls.sheet_names
        st.session_state.selected_sheet = xls.sheet_names[0]  # default

# Only render sheet picker if file was submitted
    if "sheets_lst" in st.session_state:
        selected_sheet = st.selectbox(
            "Choose a sheet",
            st.session_state.sheets_lst,
            index=st.session_state.sheets_lst.index(st.session_state.selected_sheet),
        )
        st.session_state.selected_sheet = selected_sheet

        if st.button(":rainbow[submit]"):
            df = pd.read_excel(st.session_state.file_path, sheet_name=selected_sheet, header=None)
            df = df.dropna(how="all", axis=0).dropna(how="all", axis=1)
            df = df.reset_index(drop=True)
            df.iloc[:, 0] = df.iloc[:, 0].astype(str).str.replace(r"[\n\t]+", " ", regex=True).str.strip()

            st.session_state.df = df             

            st.write("sheet loaded successfully")

          
            st.session_state.df  = st.session_state.df.dropna(how="all",axis=0).dropna(how="all",axis=1)
            st.session_state.df  =st.session_state.df .reset_index(drop=True)
            st.session_state.df.iloc[:,0] = st.session_state.df .iloc[:,0].astype(str).str.replace(r"[\n\t]+"," ",regex=True).str.strip()

            new_lst = [st.session_state.df .loc[j,i] for i in st.session_state.df .columns for j in st.session_state.df .index]
            new_lst = [str(i).replace('\n','') for i in new_lst]
            new_lst = [x for x in new_lst if x != 'nan']
            new_d = {i:col for i,col in enumerate(new_lst)}

            gst = ''
            for i in new_lst:
                if str(i).isalnum():
                    if str(i).endswith('J'):
                        gst = gst+str(i)

            st.session_state['data_d']['GSTIN'] = gst

            hsn_lst =[]
            for i in new_lst:
                if str(i).isdigit():
                    if str(i).startswith('71'):
                        hsn_lst.append(str(i))

            hsn_lst = hsn_lst.pop()

            st.session_state['data_d']['hsn_no']= hsn_lst

            unit_lst = ""
            for i in new_lst:
                if str(i).startswith('Gross'):
                    unit_lst = unit_lst+i
            ind = unit_lst.find('Gms')
            ind_1 = unit_lst[ind:]
            ind_2 = "".join(re.findall('[a-zA-Z]',ind_1))
            
            st.session_state['data_d']['units'] = ind_2

            for i in new_lst:
                if 'STOCK' in str(i):
                    stock_str = str(i)
            st_1 = " ".join(stock_str[stock_str.find('STOCK'):].split(" ")[0:2])
            
            st.session_state['data_d']['sub_type'] = st_1

            ornaments_lst = []
            for i in new_lst:
                if any(x in str(i) for x in ['ORNAMENTS', 'COIN', 'BULLION']):
                    ornaments_lst.append(i)
            if len(ornaments_lst)>1:
                ornments = ornaments_lst.pop()
                ornaments_lst_1 = "".join(ornaments_lst)
                print(ornaments_lst_1)
            else:
                ornaments_lst_1 = "".join(ornaments_lst)

            st.session_state['data_d']['item'] = ornaments_lst_1

            new_d = {i:col for i,col in enumerate(new_lst)}

            gross_wt = ''
            for x in new_d:
                if new_d[x].startswith('Gross'):        
                    gross_wt = str(round(float(gross_wt+new_d[x+1]),2))
            
            st.session_state['data_d']['gross_wt'] = gross_wt

            value = ''
            for x in new_d:
                if new_d[x].startswith('Value of Supply'):        
                    value = value+new_d[x+1]
            value_1 = float(value)
            value_2 = round(value_1,2)
            value_3 = str(value_2)
            value = value_3

            st.session_state['data_d']['value'] = value

            for x in new_d:
                if 'DC' in new_d[x]:
                    doc_no = new_d[x]
            doc_no_1 = doc_no.split("-")
            doc_no_2 = [i.strip() for i in doc_no_1]
            doc_no_2.remove('DC')
            doc_no_3 = "-".join(doc_no_2)

            st.session_state['data_d']['doc_no'] = doc_no_3
            
            for x in new_d:
                if 'MADURAI - HO' in new_d[x]:
                    address =  new_d[x+1]
            ad = address.split(",")
            add_1 = ad[0].strip()
            add_2 = ad[1].strip()
            
            st.session_state['data_d']['from_office_add_1'] = add_1
            st.session_state['data_d']['from_office_add_2'] = add_2
            
            for x in new_d:
                if '625016' in new_d[x]:
                    city = new_d[x]
            cty = city.split("-")
            place= cty[0].strip()
            pin_code = cty[1].strip()
            
            st.session_state['data_d']['from_city'] = place
            st.session_state['data_d']['from_city_pincode'] = pin_code
            
            to_office = []
            for x in new_d:
                if "BHIMA" in new_d[x]:
                    to_office.append(new_d[x])
        
            for i in to_office:
                if 'HO' in i:
                    to_office.remove(i)
            to_office_1 = "".join(to_office).strip()
            
            st.session_state['data_d']['to_office'] = to_office_1

            for  x in new_d:
                if to_office_1 in new_d[x]:
                    to_office_add_1 = new_d[x+1].strip()
                    to_office_add_1 = re.sub(r'[^A-Za-z0-9 ]+', ' ', to_office_add_1)
            
            st.session_state['data_d']['to_office_add_1'] = to_office_add_1

            for  x in new_d:
                if to_office_1 in new_d[x]:
                    to_office_add_2 = new_d[x+2].strip()
                    to_office_add_2 = re.sub(r'[^A-Za-z0-9 ]+', ' ', to_office_add_2)
            
            st.session_state['data_d']['to_office_add_2'] = to_office_add_2

            city_name = "".join(to_office).split()[-1]

            
            city_lst = ['MADURAI','RAJPALAYAM','DINDIGUL','TRICHY','SALEM','RAJAPALYAM','RAJAPALAYAM','DINDUGAL']

            city_names =[new_d[x] for x in new_d if city_name in new_d[x]]

            if len(city_names)>1:
                st.session_state['data_d']['to_office_city'] = city_names[-1].strip()
            else:
                st.session_state['data_d']['to_office_city'] = "".join(city_names).split()[-1].strip()

            st.session_state['data_d']['to_office_city'] = re.sub(r'[^a-zA-Z]','',st.session_state['data_d']['to_office_city']).strip()

            if st.session_state['data_d']['to_office_city'] == 'MADURAI':
                st.session_state['data_d']['to_office_add_1'] = '137'
                st.session_state['data_d']['to_office_add_2'] = 'WEST MASI STREET'
            else:
                st.session_state['data_d']['to_office_add_1'] = st.session_state['data_d']['to_office_add_1']
                st.session_state['data_d']['to_office_add_2'] = st.session_state['data_d']['to_office_add_2']

            pincode_d = {'TRICHY' : '620002','MADURAI' : '625001','RAJAPALYAM':'626117','DINDIGUL':'624001','DINDUGUL':'624001',
                        'RAJAPALAYAM':'626117'}   
            pincode_d['DINDUGAL'] = '624001'
            pincode_d['RAJPALAYAM'] = '626117'

            if st.session_state['data_d']['to_office_city'] in city_lst:
                st.session_state['data_d']['pincode'] = pincode_d[st.session_state['data_d']['to_office_city']]    
            
            st.success("✅ File processed and dictionary stored.")
            
            df_1= pd.DataFrame(list(st.session_state['data_d'].items()),columns=['Items','Values'])
            st.dataframe(df_1)

    if st.button(":rainbow[Fill_Form]"):

        if "driver" in st.session_state and "data_d" in st.session_state:

            try:

                driver = st.session_state.driver
                wait = WebDriverWait(driver, 20)

                pdt_inpt = wait.until(EC.presence_of_element_located((By.ID, "txtProductName_1")))
                pdt_inpt.clear()

                item_des = wait.until(EC.presence_of_element_located((By.ID, "txt_Description_1")))
                item_des.clear()

                hsn_inpt = wait.until(EC.presence_of_element_located((By.ID, "txt_HSN_1")))
                hsn_inpt.clear()

                quant_inpt = wait.until(EC.presence_of_element_located((By.ID, "txt_Quanity_1")))
                quant_inpt.clear()

                units = wait.until(EC.presence_of_element_located((By.ID, "txt_Unit_1")))
                units.clear()
                
                value = wait.until(EC.presence_of_element_located((By.ID, "txt_TRC_1")))
                value.clear()

                outward_radio = wait.until(EC.element_to_be_clickable((By.ID, "ctl00_ContentPlaceHolder1_rbtOutwardInward_0")))
                outward_radio.click()
                #st.info(outward_radio.is_selected())
                st.success('Outward is Selected')
                
                wait.until(EC.invisibility_of_element_located((By.CSS_SELECTOR, "div.ui-widget-overlay.ui-front")))
                others_radio = wait.until(EC.element_to_be_clickable((By.ID, "ctl00_ContentPlaceHolder1_rbtOutSubType_8")))
                others_radio.click()
                #st.info(others_radio.is_selected())
                st.success('Others_radio is Selected')

                others_input = wait.until(EC.presence_of_element_located((By.ID, "txtSpecify")))
                others_input.clear()
                others_input.send_keys('STOCK TRANSFER ISSUE')
                st.success('Successefully Entered STOCK TRANSFER')

                doc_type_dropdown = wait.until(EC.presence_of_element_located((By.ID, "ddlDocType")))
                select = Select(doc_type_dropdown)
                select.select_by_value('CHL')
                st.success("Document Type set to Delivery Challan")

                doc_no_input = wait.until(EC.presence_of_element_located((By.ID, "txtDocNo")))
                doc_no_input.clear()
                doc_no_input_value  = st.session_state['data_d']['doc_no']
                doc_no_input.send_keys(doc_no_input_value)
                st.success("Document no.Entered")

                current_date = date.today()
                formatted_date = current_date.strftime("%d/%m/%Y")

                current_date = date.today()
                formatted_date = current_date.strftime("%d/%m/%Y")
                date_input = wait.until(EC.presence_of_element_located((By.ID, 'txtDocDate')))
                date_input.send_keys(formatted_date)
                date_input.send_keys(Keys.TAB)
                st.success(f"Document Date entered: {formatted_date}")

                trans_type = wait.until(EC.presence_of_element_located((By.ID, "ddlTransType")))
                select = Select(trans_type)
                select.select_by_value('1')
                st.success("Transction_type set to Regular")

                firm_name_input = wait.until(EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_txtFromTrdName")))
                firm_name_input.clear()
                firm_name_input.send_keys("BHIMA JEWELLERY MADURAI - HO")
                st.success('Successfully entered HO Name')

                to_firm_gstnum_input = wait.until(EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_txtToGSTIN")))
                to_firm_gstnum_input.send_keys(st.session_state['data_d']['GSTIN'])
                st.info('Successfully entered GSTIN')

                time.sleep(2)

                addrss_1_input = wait.until(EC.presence_of_element_located((By.ID, "txtFromAddr1")))
                addrss_1_input.clear()
                addrss_1_input.send_keys(st.session_state['data_d']['from_office_add_1'])
                st.success('Successfully entered HO Address')

                addrss_2_input = wait.until(EC.presence_of_element_located((By.ID, "txtFromAddr2")))
                addrss_2_input.clear()
                addrss_2_input.send_keys(st.session_state['data_d']['from_office_add_2'])
                st.success('Successfully entered HO Address')

                pincode_input = wait.until(EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_txtFromPincode")))
                pincode_input.clear()
                pincode_input.send_keys(st.session_state['data_d']['from_city_pincode'])
                st.success('Successfully entered HO Address')

                to_firm_gstnum_input = wait.until(EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_txtToGSTIN")))
                #to_firm_gstnum_input.clear()
                to_firm_gstnum_input.send_keys(st.session_state['data_d']['GSTIN'])
                print('Successfully entered')

                time.sleep(4)

                to_name_input = wait.until(EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_txtToTrdName")))
                to_name_input.clear()
                to_name_input_value  = st.session_state['data_d']['to_office']
                to_name_input.send_keys(to_name_input_value.upper())
                st.success('Successfully entered Consignee Branch Name')

                time.sleep(2)        

                to_office_add = wait.until(EC.presence_of_element_located((By.ID, "txtToAddr1")))
                to_office_add.clear()
                to_office_add_value= st.session_state['data_d']['to_office_add_1']
                to_office_add.send_keys(to_office_add_value.upper())
                st.success('Successfully entered Consignee Branch Address1')

                to_office_address = wait.until(EC.presence_of_element_located((By.ID, "txtToAddr2")))
                to_office_address.clear()
                to_office_address_value= st.session_state['data_d']['to_office_add_2']
                to_office_address.send_keys(to_office_address_value.upper())
                st.success('Successfully entered Consignee Branch Address2')

                to_place_input = wait.until(EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_txtToPlace")))
                to_place_input.clear()
                to_place_input_value = st.session_state['data_d']['to_office_city']
                to_place_input.send_keys(to_place_input_value.upper())
                st.success('Successfully entered Consignee Branch City')

                to_pincode_input = wait.until(EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_txtToPincode")))
                to_pincode_input.clear()
                to_pincode_input.send_keys(st.session_state['data_d']['pincode'])
                st.success('Successfully entered Consignee Branch Pincode')

                item = st.session_state['data_d']['item']
                pdt_inpt = wait.until(EC.presence_of_element_located((By.ID, "txtProductName_1")))
                pdt_inpt.clear()
                pdt_inpt.send_keys(item)
                pdt_inpt.send_keys(Keys.TAB)

                
                item_des = wait.until(EC.presence_of_element_located((By.ID, "txt_Description_1")))
                item_des.clear()
                item_des.send_keys(st.session_state['data_d']['item'])
                item_des.send_keys(Keys.TAB)

                hsn_inpt = wait.until(EC.presence_of_element_located((By.ID, "txt_HSN_1")))
                hsn_inpt.clear()
                hsn_inpt.send_keys(st.session_state['data_d']['hsn_no'])
                hsn_inpt.send_keys(Keys.TAB)

                quant = st.session_state['data_d']['gross_wt']
                quant_inpt = wait.until(EC.presence_of_element_located((By.ID, "txt_Quanity_1")))
                quant_inpt.clear()
                quant_inpt.send_keys(quant)
                quant_inpt.send_keys(Keys.TAB)

                units = wait.until(EC.presence_of_element_located((By.ID, "txt_Unit_1")))
                units.clear()
                units.send_keys(st.session_state['data_d']['units'])
                units.send_keys(Keys.TAB)

                value_inpt= float(st.session_state['data_d']['value'])
                value = wait.until(EC.presence_of_element_located((By.ID, "txt_TRC_1")))
                value.clear()
                value.send_keys(value_inpt)
                value.send_keys(Keys.TAB)

                tax_select = wait.until(EC.presence_of_element_located((By.ID, "SelectCSGST_1")))
                select = Select(tax_select)
                select.select_by_value("1.500")
                time.sleep(4)

                cgst_inpt = wait.until(EC.presence_of_element_located((By.ID, "txtCGST")))
                sgst_inpt = wait.until(EC.presence_of_element_located((By.ID, "txtSGST")))
                cgst_val = cgst_inpt.get_attribute("value")
                sgst_val = sgst_inpt.get_attribute("value")
                if cgst_val:
                    cgst_val = int(round(float(cgst_val),0))
                if sgst_val:
                    sgst_val = int(round(float(sgst_val),0))

                time.sleep(3)

                grand_total = wait.until(EC.presence_of_element_located((By.ID, "txtTotInvVal")))
                grand_tot_val = grand_total.get_attribute("value")
                grand_tot_val = math.floor(float(grand_tot_val))
                grand_total.clear()
                grand_total.send_keys(str(grand_tot_val))
                st.success('All Item Details Updated')

            
            except TimeoutException:
                st.error("⚠️ OTP verification failed: e-Waybill menu not found.")

