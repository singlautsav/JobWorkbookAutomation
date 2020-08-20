from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import time

chrome_options = Options()
chrome_options.add_argument('no-sandbox')
# chrome_options.add_argument('--headless')
chrome_options.add_argument("--user-data-dir=chrome-data")
driver = webdriver.Chrome(r'D:\chromedriver.exe',options=chrome_options)
driver.get('https://web.whatsapp.com')
phoneNum = "91"+str(input("Enter Phone Number: "))
link = "https://web.whatsapp.com/send?phone={}&text&source&data&app_absent".format(i)

time.sleep(30)