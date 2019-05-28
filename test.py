#steps:
#1. Open browser
#2. Open the selected website
#3. take screenshot
#4. paste in word file.

from selenium import webdriver
import docx
import time
from datetime import datetime

links = ['https://duckduckgo.com','https://google.com'];

now=datetime.now()
filename=now.strftime("%m%d%Y%H%M%S")
print(filename)
#print(now.split('.')[0].replace(':','').strip())

doc=docx.Document()
filepath = 'C:\\Users\\Hemachandar\\Downloads\\tdr'+filename+'.docx' 
doc.save(filepath)

browser=webdriver.Chrome(executable_path='C:/Users/Hemachandar/Downloads/chromedriver.exe');
browser.maximize_window()

for x in links:
    
    
    browser.get(x)
    filename='C:\\Users\\Hemachandar\\Downloads\\' + x.split('//')[1].split('.')[0] + '_screenshot.png'
    browser.save_screenshot("C:\\Users\\Hemachandar\\Downloads\\" + x.split('//')[1].split('.')[0] + "_screenshot.png")
    #screenshot = browser.get_screenshot_as_png();
    doc.add_picture(filename,width=docx.shared.Inches(7),height=docx.shared.Inches(11))
    doc.save(filepath)
    
    time.sleep(2)
browser.close()
