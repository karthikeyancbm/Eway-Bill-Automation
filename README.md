# Eway-Bill-Automation

# Problem Statement:
  This project aim is to automating Eway Bill in eway gst portal by uploading the excel file which has required fields those need to be entered in the eway gst portal.

# Approach:

**Data Cleaning:**

1.Cleaned the data and extracted the required fields and stored it in the dictionary.

2.Finding and mapping the respective html ids and tags with the corresponding inputs.

3.Developed the frontend using streamlit.

4.Created a batch file to fecilitate to run in all nodes in LAN.

5.Flow: 
  * The user need to upload the excel file and click the subit button,the user will get required fields in the dictionary format.
  * Then need to click Eway bill button,then the user has to enter the captcha and otp and after submitting OTP,all fields in the portal will get filled automatically and end up in the printer form.

