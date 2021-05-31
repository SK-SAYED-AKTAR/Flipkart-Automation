from selenium import webdriver
import time
import xlsxwriter
from selenium.webdriver.common.keys import Keys

move=1

def sendRequest(siteUrl):
    #Navigate to flipkart site
    driver.get(siteUrl)

    #Ignore the Login Part
    try:
        crossButton = driver.find_element_by_xpath("/html/body/div[2]/div/div/button")
        time.sleep(4)
        crossButton.click()
    except:
        print("An exception occured !!!")

def searchYourItem():
    #Locate and click on the search bar
    searchBar = driver.find_element_by_xpath("//*[@id='container']/div/div[1]/div[1]/div[2]/div[2]/form/div/div/input")
    searchBar.click()

    #Search Your item
    searchBar.send_keys("Mobiles")
    searchBar.send_keys(Keys.ENTER)
    time.sleep(2)



def grabInformation():
    global move

    allTitleClass = driver.find_elements_by_class_name("_4rR01T")
    allOfferPrice = driver.find_elements_by_class_name("_30jeq3")
    allOriginalPrice = driver.find_elements_by_class_name("_3I9_wc")
    allUrl = driver.find_elements_by_class_name("_1fQZEK")


    for i in allTitleClass:
        title.append(i.text)
        # print(i.text)

    for i in allOfferPrice:
        offer.append(i.text)
        # print(i.text)

    for i in allOriginalPrice:
        originalPrice.append(i.text)
        # print(i.text)

    for i in allUrl:
        myurl=i.get_attribute('href')
        productUrl.append(myurl)
        # print(i.text)


    if move<=2:
        move=move+1
        move_to_next_page()


def move_to_next_page():
    next_page = driver.find_element_by_xpath("//*[@id='container']/div/div[3]/div[1]/div[2]/div[26]/div/div/nav/a[11]")
    next_page.click()

    time.sleep(3)
    grabInformation()


def insertIntoExcel(title, offer, original):
    row=1
    with xlsxwriter.Workbook("ItemList.xlsx") as workbook:
        workshit = workbook.add_worksheet()
        # workshit = workbook.get_worksheet_by_name("ItemList.xlsx")
        #Initialize the title
        workshit.write("A1", "Item Name")
        workshit.write("B1", "Offer Price")
        workshit.write("C1", "Orginal Price")
        workshit.write("D1", "Product URL")

        #Write Down the Information
        for n, m, a, u in zip(title, offer, original, productUrl):
            row=row+1
            workshit.write("A"+str(row), n)
            workshit.write("B"+str(row), m)
            workshit.write("C"+str(row), a)
            workshit.write("D"+str(row), u)

if __name__=="__main__":

    # Take the site URL
    siteUrl = input("Please enter the url:")

    driver = webdriver.Chrome()
    driver.maximize_window()


    #Initialize Information Category
    title=[]
    offer=[]
    originalPrice=[]
    productUrl=[]


    sendRequest(siteUrl)

    searchYourItem()

    grabInformation()

    insertIntoExcel(title, offer, originalPrice)

