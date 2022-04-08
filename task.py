import pandas as pd

from RPA.Browser.Selenium import Selenium
import requests
from RPA.HTTP import HTTP
import glob
import pyautogui
import os

from bs4 import BeautifulSoup

from RPA.Desktop.Windows import Windows

browser = Selenium(auto_close=False)
down = HTTP()
lib = Windows()

def login():
    browser.open_chrome_browser(url='https://botsdna.com/school/')
    browser.maximize_browser_window()
    
def download_data():
    down.download(url="https://botsdna.com/school/Master%20Template.xlsx",target_file="data/")

def close_tab():
    pyautogui.hotkey('ctrl', 'w')

def del_files(filenames):
    filenames=filenames

    for x in filenames:

        os.remove(x)

def enter_data(text):
    browser.input_text(locator='//*[@id="SchoolCode"]', text=text, clear=True)
    browser.click_button(locator='//*[@id="SearchSchool"]')
    browser.wait_until_element_is_visible(locator="css:body", timeout="6")
    


def load_data():
    df1 = pd.read_excel("data/Master%20Template.xlsx")
    mypath = "data/"
    df2 = pd.DataFrame(columns=['School Code','School Name:','School Address :',
    'School Phonenumber :','Student\'s Strenth :','Prncipal Name :',
    'Number of TeachingStaff :','Number of Non-TeachingStaff :',
    'Number of School buses :','School Playground :',
    'Facilities :', 'School Accrediation :','School Hostel :', 
    'School Canteen :', 'School Stationary :','School Teaching method\'s :', 
    'School Timing :', 'School Achivements :','School Awards :', 
    'School Uniform :'])
    
    for x in df1['School Code']:
        print(x)
        enter_data(x)
        school_data = get_data(x)
        pd.concat([df2, school_data],ignore_index=True )
        school_data.to_csv("data/"+str(x)+".csv", index=False)
        close_tab()

        # break
    
    filenames = glob.glob(mypath+"*.csv")
    fi = []
    for filename in filenames:
        df = pd.read_csv(filename, index_col=None, header=0)
        fi.append(df)

    df = pd.concat(fi, axis=0, ignore_index=True)
    del_files(filenames)
    return df.to_excel("data/school_details.xlsx")

def get_data(school_id):
    r = requests.get(f"https://botsdna.com/school/{school_id}.html")
    soup = BeautifulSoup(r.text, "lxml")
    df = pd.read_html(f"https://botsdna.com/school/{school_id}.html")
    df = pd.DataFrame(df[0])
    df = df.T
    df.columns = df.iloc[0]
    df = df[1:]
    df['School Name:'] = soup.find('h1').text
    df['School Code'] = school_id
    df = df[['School Code','School Name:','School Address :',
    'School Phonenumber :','Student\'s Strenth :','Prncipal Name :',
    'Number of TeachingStaff :','Number of Non-TeachingStaff :',
    'Number of School buses :','School Playground :',
    'Facilities :', 'School Accrediation :','School Hostel :', 
    'School Canteen :', 'School Stationary :','School Teaching method\'s :', 
    'School Timing :', 'School Achivements :','School Awards :', 
    'School Uniform :']]
    # df = df[:].values
    return df



def main():
    try:
        download_data()
        login()
        load_data()
    finally:
        close_tab()



if __name__ == "__main__":
    main()