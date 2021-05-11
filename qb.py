################
'''
The goal of this program is to automate data entry into QuickBooks.
Using Webdriver and Selenium, the program signs into your QuickBooks account for you
and proceeds by reading in a CSV file and entering in the bills. 

The program is equipped to handle multiple different BIll #'s within a single CSV.

To get around on the Bill page, we Tab our way around each textbox grab the active element 
and perform an operation as seen below. 

This code can be configured to work for invoices, or any data that needs to be entered,
by finding the appropriate HTML tags correspoding to the weboage

Everytime I sleep the program, it is equivalent to WebDriverWait. QuickBooks takes awhile to 
load and the program would crash if it tries to execute before the page loads.

To run the program:
~ In the required areas ~
1. Input your username and password
2. Input the path to chromedriver on your computer
3. Input the path to your CSV

Happy Automating!
'''
################
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
import pandas as pd


#INPUT PATH TO CHROMEDRIVER HERE
path_to_driver = "./chromedriver" 

driver = webdriver.Chrome(path_to_driver)
driver.get("https://quickbooks.intuit.com/sign-in-offer/")

driver.maximize_window()

#INPUT USERNAME AND PASSWORD HERE
username = ""
password = ""


#QUICKBOOKS SIGN IN
#find the login and password id, grab it and send keys 
login_user = driver.find_element_by_id("ius-userid")
login_user.click()
login_user.send_keys(username)

sign_in_pword = driver.find_element_by_id("ius-password")
sign_in_pword.click()
sign_in_pword.send_keys(password)
sign_in_pword.send_keys(Keys.RETURN)

time.sleep(15)


#READ IN CSV 
#INPUT PATH TO CSV HERE 
df = pd.read_csv(r'')


#GO TO MENU AND CLICK BILL
#find bill and create tags and click it 
menu_button = driver.find_element_by_class_name("left-nav-default-create")
menu_button.click()
time.sleep(5)
bill_link = driver.find_element_by_id("bill")
bill_link.click()

time.sleep(10)


#VENDOR DETAILS
#find vendor id and start entering data from csv 
vendor = driver.switch_to.active_element
vendor.click()

#grab_vendor = df["*Vendor"].values[0]
#vendor.send_keys(grab_vendor)
vendor.send_keys(Keys.ARROW_DOWN) 
time.sleep(5)
vendor.send_keys(Keys.ARROW_DOWN*3) 

vendor.send_keys(Keys.RETURN)
vendor.send_keys(Keys.TAB * 3)


#BILL DATE DETAILS
#find bill id and enter data from csv
bill_date = driver.switch_to.active_element
bill_date.click()
bill_date.send_keys(Keys.BACKSPACE*10) #delete prefilled info 
time.sleep(3)
date = df["*Bill Date"].values[0]
bill_date.send_keys(date)
bill_date.send_keys(Keys.RETURN)

time.sleep(3)


#OPEN UP ITEM DETAILS TAB
#open up item details rows, it starts closed 
try:
	item_details = driver.find_element_by_xpath("//*[text()='Item details']")
	item_details.click()
	print("success")
except:
	driver.quit()


#BILL NUMBER DETAILS
#enter bill number bu tabbing and grabbing element 
bill_date.send_keys(Keys.TAB * 2)
bill_number = driver.switch_to.active_element

date = df["Bill No."].values[0]
str_date = str(date)
bill_number.send_keys(str_date)
bill_number.send_keys(Keys.RETURN)


#TAB TO PRODUCT/SERVICE ROW1 COLUMN1 IN ITEM DETAILS 
bill_number.send_keys(Keys.TAB * 10)

#TRAVERSE THE CSV DATA by lopping through each row 
for index,row in df.iterrows():

	if str_date == str(df["Bill No."].values[index]): #how to check if you need to start a new bill

		#PRODUCT/SERVICE
		product = driver.switch_to.active_element
		product.click()
		prod = df["Product/Service"].values[index]
		prodInt = (str(prod.item()))
		product.send_keys(prodInt)

		time.sleep(1) 

		product.send_keys(Keys.ARROW_DOWN) 
		product.send_keys(Keys.ARROW_DOWN)
		product.send_keys(Keys.RETURN)

		time.sleep(1) 

		#QUANTITY
		product.send_keys(Keys.TAB * 2)
		qty = driver.switch_to.active_element
		quant = df["Qty"].values[index]
		quantInt = (str(quant.item()))
		qty.send_keys(quantInt)

		#RATE
		qty.send_keys(Keys.TAB)
		rate = driver.switch_to.active_element
		rater = date = df["Rate"].values[index]
		rateInt = (str(rater.item()))
		rate.send_keys(rateInt)
		rate.send_keys(Keys.TAB*4)


	else: #if bill #'s are different save the bill and do it again 
		#GO TO MEMO
		product.send_keys(Keys.TAB * 9)
		memo = driver.switch_to.active_element
		memo_grab = str(df["Memo"].values[index-1])
		memo.send_keys(memo_grab)

		time.sleep(2)
		memo.send_keys(Keys.TAB*217)

		#SAVE
		save = driver.switch_to.active_element
		save.click()

		time.sleep(10)

		#CLOSE THE BILL
		save.send_keys(Keys.TAB*6)
		x_button = driver.switch_to.active_element
		x_button.click()
		
		time.sleep(10)

		#GO TO MENU AND CLICK BILL
		menu_button = driver.find_element_by_class_name("left-nav-default-create")
		menu_button.click()
		bill_link = driver.find_element_by_id("bill")
		bill_link.click()

		time.sleep(10)

		#VENDOR DETAILS
		vendor = driver.switch_to.active_element
		vendor.click()
		vendor.send_keys(Keys.ARROW_DOWN) 
		time.sleep(5)
		vendor.send_keys(Keys.ARROW_DOWN*3) 
		vendor.send_keys(Keys.RETURN)
		vendor.send_keys(Keys.TAB * 3)

		
		#BILL DATE DETAILS
		bill_date = driver.switch_to.active_element
		bill_date.click()
		bill_date.send_keys(Keys.BACKSPACE*10)
		time.sleep(3)

		date = df["*Bill Date"].values[index]
		bill_date.send_keys(date)
		bill_date.send_keys(Keys.RETURN)

		time.sleep(3)

		''
		#OPEN UP ITEM DETAILS TAB
		try:
			item_details = driver.find_element_by_xpath("//*[text()='Item details']")
			item_details.click()
			print("success")
		except Exception as e:
		 	print(e)
		 	driver.quit()
			
		time.sleep(2)

		
		#BILL NUMBER DETAILS
		bill_date.send_keys(Keys.TAB * 2)
		bill_number = driver.switch_to.active_element

		date = df["Bill No."].values[index]
		str_date = str(date)
		bill_number.send_keys(str_date)
		bill_number.send_keys(Keys.RETURN)

		#TAB TO PRODUCT/SERVICE ROW1 COLUMN1 IN ITEM DETAILS 
		bill_number.send_keys(Keys.TAB * 10)

		#PRODUCT/SERVICE
		product = driver.switch_to.active_element
		product.click()
		prod = df["Product/Service"].values[index]
		prodInt = (str(prod.item()))
		product.send_keys(prodInt)

		time.sleep(1) #WAS 2

		product.send_keys(Keys.ARROW_DOWN) 
		product.send_keys(Keys.ARROW_DOWN)
		product.send_keys(Keys.RETURN)

		time.sleep(1) #WAS 3


		#QUANTITY
		product.send_keys(Keys.TAB * 2)
		qty = driver.switch_to.active_element
		quant = df["Qty"].values[index]
		quantInt = (str(quant.item()))
		qty.send_keys(quantInt)
			

		#RATE
		qty.send_keys(Keys.TAB)
		rate = driver.switch_to.active_element
		rater = df["Rate"].values[index]
		rateInt = (str(rater.item()))
		rate.send_keys(rateInt)
		rate.send_keys(Keys.TAB*4)


time.sleep(10)

#MEMO
product.send_keys(Keys.TAB * 9)
memo = driver.switch_to.active_element
memo_grab = str(df["Memo"].values[index])
memo.send_keys(memo_grab)


#SAVE BUTTON 
memo.send_keys(Keys.TAB*217)
save = driver.switch_to.active_element
save.click()

time.sleep(10)

#CLOSE THE BILL
save.send_keys(Keys.TAB*6)
x_button = driver.switch_to.active_element
x_button.click()

time.sleep(2)
driver.quit()









