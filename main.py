#Google Spreadsheet API Connection
#Armatör kodu ve booking numarası
#Armatör kodu ile uygun siteye giriş ve booking numarası ile aratma
#Alınan verileri google spreadsheete yazma



from __future__ import print_function

import os.path
import selenium
import time
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from selenium.webdriver.common.by import By
from selenium.webdriver.remote.webelement import WebElement
from selenium.common.exceptions import NoSuchElementException

driver = selenium.webdriver.Safari()

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

# The ID and range of a sample spreadsheet.
SAMPLE_SPREADSHEET_ID = '12iBFv282e2Psgq_BAm1igvhIvc1IZAZ7OJnoJt6VnL4'
SAMPLE_RANGE_NAME = 'Booking!A:B'


def slow_type(element: WebElement, text: str, delay: float=0.1):
    """Send a text to an element one character at a time with a delay."""
    for character in text:
        element.send_keys(character)
        time.sleep(delay)

def API_Connection():
    """Shows basic usage of the Sheets API.
    Prints values from a sample spreadsheet.
    """
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    try:
        service = build('sheets', 'v4', credentials=creds)

        # Call the Sheets API
        sheet = service.spreadsheets()
        result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                    range=SAMPLE_RANGE_NAME).execute()
        values = result.get('values', [])
        SelectShipOwner(values)

        if not values:
            print('No data found.')
            return
    except HttpError as err:
        print(err)


#//*[@id="select2-select-carrier-results"]/li
#//*[@id="select2-select-carrier-results"]/li
def SelectShipOwner(values):
    shipsgo_url = "https://shipsgo.com/"
    for row in values:
        driver.get(shipsgo_url)
        driver.maximize_window()
        spanclass = driver.find_element(By.XPATH,'//*[@id="select2-select-carrier-container"]')
        spanclass.click()
        textbox = driver.find_element(By.XPATH,'/html/body/span/span/span[1]/input')
        slow_type(textbox,row[0])
        time.sleep(0.5)
        highlitedbutton = driver.find_element(By.XPATH,'//*[@id="select2-select-carrier-results"]/li')
        time.sleep(0.5)
        highlitedbutton.click()
        bookingnotextbox = driver.find_element(By.XPATH,'//*[@id="input-number"]')
        slow_type(bookingnotextbox,row[1])
        searchbutton = driver.find_element(By.XPATH,'//*[@id="home-tracking"]/form/div/div[3]/div/button')
        searchbutton.click()
        time.sleep(5)

if __name__ == '__main__':
    API_Connection()