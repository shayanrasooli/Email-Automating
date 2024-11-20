from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
import yagmail
import time
import win32com.client as win32
import base64
from PIL import Image
import os
import time
# Configure Selenium
options = Options()
options.headless = True  # Run browser in headless mode
driver = webdriver.Chrome(options=options)  # Make sure to download and use the appropriate WebDriver

# URL of the Power BI Dashboard
dashboard_url_1 = "http://t3vw077:8090/DigitalBIReports/powerbi/AAT_report?rs:embed=true"
dashboard_url_2 = "http://t3vw077:8090/DigitalBIReports/powerbi/AAT_report?rs:embed=true"

dashboard_url = [dashboard_url_1,dashboard_url_2]
sheets_to_capture = ["Revenue amount", "recovery_labs"]
# Visit the dashboard

# Take screenshots for each sheet
screenshots = []
for i , page in list(enumerate(dashboard_url)):
    driver.get(page)
    time.sleep(5)
    # Capture screenshot
    file_name = f"{sheets_to_capture[i]}_screenshot.png"
    driver.save_screenshot(file_name)
    screenshots.append(file_name)
    time.sleep(2)  # Optional: brief pause between captures
driver.quit()
# Close the driver


# Send Email with Screenshots
olApp = win32.Dispatch("Outlook.Application")
olNS = olApp.GetNameSpace("MAPI")
mail_item = olApp.CreateItem(0)
mail_item.BodyFormat = 1
mail_item.Sender = "shayan.r@mtnirancell.ir"
mail_item.To = "sana.a@mtnirancell.ir"
mail_item.CC = "sana.a@mtnirancell.ir"
mail_item.Subject = "AAT_Report"
mail_item.HTMLBody = "<h4>Dear Larry,<br>This is an Automatic email.<br>You can find AAT report as below:</h4>"

base_path = "E:\\Irancell\\Digital\\Automatic powe BI send email\\"

for i , screenshot in enumerate(screenshots,0):
    img_path = f"{base_path}{screenshot}"
    img = Image.open(img_path)
#   img.save("E:/Irancell/Digital/Automatic powe BI send email/{}".format(screenshot[i])
#   mail_item.Attachments.Add(screenshot)
    mail_item.HTMLBody += f'<p>{screenshot}:</p><img src="{img_path}" width="600"><br>'
    
#cid:{screenshot}
mail_item.Display()
mail_item.save()
#mail_item.send


