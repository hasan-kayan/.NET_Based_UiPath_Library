from bs4 import BeautifulSoup
import requests
import csv
import undetected_chromedriver as uc
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import time
from bs4.element import Comment
from datetime import date, timedelta

def printer(URL,Type,Note):
    with open("./output.csv", 'a') as result:
        fieldnames = ["URL", "Type", "Note"]
        #TODO Append
        writer = csv.writer(result)
        writer.writerow(fieldnames)
        writer.writerow([URL, Type, Note])
        print(URL+" - "+Type+" - "+Note)

def tag_visible(element):
    if element.parent.name in ['style', 'script', 'head', 'title', 'meta', '[document]']:
        return False
    if isinstance(element, Comment):
        return False
    return True

def checkImageQuality(soup,url):

    driver.execute_script("window.scrollTo({ top: 10000, behavior: 'smooth' })")
    time.sleep(5)
    images = soup.find_all('img',class_ ='inline-block')
    for image in images:
        if  image.get('src').find('blur=') > -1:
            printer(url,"FAIL","Error : Low Resolution Images - "+url+" - "+image.get('src'))
            return False
    return True

def checkTranslate(soup,url):
    texts = soup.findAll(string=True)
    visible_texts = filter(tag_visible, texts)
    visible_texts = list(visible_texts)
    for text in visible_texts:
        if text.isascii():
            printer(url,"FAIL","Error : Inner pages not translated - "+url+" - "+text)
            return False
    return True


def mainPage(url):

    driver.get(url)

    try:
        WebDriverWait(driver, 30).until(EC.presence_of_element_located
        ((By.CLASS_NAME, "contain-page ")))
        time.sleep(2)
    except TimeoutException:
        print("\n\nPage TimeOut, please check your Internet Connection.")
        driver.quit()

    
    soup = BeautifulSoup(driver.page_source, "html.parser")
    checkTrueSite = soup.find('meta',attrs={'name':'application-name'}).get('content')
    if checkTrueSite != "Class Central":
        printer(url,"FAIL","Wrong Page")
        return
    


    checkDropDown= True
    children = soup.find('a',class_ ='symbol-report').findChildren("font")
    for child in children:
        print(child)
        checkDropDown=False
    if not checkDropDown:
        printer(url,"FAIL","Javascript dropdown not working properly")
        return
    
    if checkImageQuality(soup,url):
        if checkTranslate(soup,url):
            subUrls = soup.find_all('a')
            for subUrl in subUrls:
                if subUrl.get('href').find('https://www.classcentral.com/course/') > -1:
                    subSoup=BeautifulSoup(driver.page_source, "html.parser")
                    if checkImageQuality(subSoup,subUrl.get('href')):
                        if checkTranslate(subSoup,subUrl.get('href')):
                            continue
            printer(url,"PASS","")

params = {
    "count": 10
}
driver = uc.Chrome()
links_str = input("Enter a comma-separated list of links: ")
links = links_str.split(", ")






for url in links:
    mainPage(url)
driver.minimize_window()
driver.quit()