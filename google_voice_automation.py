from playwright.sync_api import sync_playwright
from playwright_stealth import stealth_sync
from time import sleep 
import win32com.client
import pandas as pd 
import traceback 
import pickle 


# global variable and initialasation 
excelApp = win32com.client.GetActiveObject('Excel.Application')
excel_file = excelApp.Workbooks(1)
auth_worksheet = excel_file.Worksheet(3)
text_worksheet = excel_file.Worksheet(2)


# helper functions ------------------------------------------------------------------------------------------------------#
def send_text(page,row):
    global text_worksheet 
    phone = text_worksheet.Cells(row,'M')
    text = text_worksheet.Cells(row,'N')
    delay = int(text_worksheet.Cells(row,'O'))

    page.wait_for_selector('//div[contains(text(),"Send new message")]')
    page.click('//div[contains(text(),"Send new message")]')

    page.wait_for_selector('//label[contains(text(),"Type a name or phone number")]/following-sibling::input')
    page.fill('//label[contains(text(),"Type a name or phone number")]/following-sibling::input',phone)
    page.keyboard.press('Enter')
    
    page.wait_for_selector('textarea')
    page.fill('textarea',text)

    page.wait_for_selector('button[aria-label="Send message"]')
    page.click('button[aria-label="Send message"]')
    sleep(delay)

def login(auth_sheet,page):
    page.goto('https://voice.google.com/',timeout=60000)

    page.wait_for_selector('//a[contains(text(),"Sign in")]')
    page.click('//a[contains(text(),"Sign in")]')

    page.wait_for_selector('input[type="email"]')
    page.fill('input[type="email"]',auth_worksheet.Cells(2,1))

    page.wait_for_selector('//span[contains(text(),"Next") or contains(text(),"Suivant")]')
    page.click('//span[contains(text(),"Next") or contains(text(),"Suivant")]')

    page.wait_for_selector('input[type="password"]') 
    page.fill('input[type="password"]',auth_worksheet.Cells(2,2))

    page.wait_for_selector('//span[contains(text(),"Next") or contains(text(),"Suivant")]')
    page.click('//span[contains(text(),"Next") or contains(text(),"Suivant")]')

    page.wait_for_timeout(3000)
    page.goto('https://voice.google.com/u/0/messages',timeout=60000)

if __name__ == '__main__':
    with sync_playwright() as p :
        browser = p.chromium.launch(headless=False)
        context = browser.new_context()
        page = context.new_page()
        stealth_sync(page)
        login(auth_worksheet,page)
        row = 2
        while not text_worksheet.Cells(row,'O') : 
            try :
                send_text(page,row)
            except :
                print(traceback.format_exc())
