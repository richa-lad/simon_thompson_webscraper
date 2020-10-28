#! python3
# opens up simon thompson websites
# opens articles advising on what to buy
# searches article for "strong buy"
# if it finds phrase, copy article link to excel sheet
# if not, move to next article


from selenium import webdriver  # get web function from selenium module
import openpyxl
from openpyxl import workbook
from datetime import date

url = "https://www.investorschronicle.co.uk/simon-thompson/"  # webpage where articles are listed

myworkbook = openpyxl.load_workbook(
    "C:\\Users\\richa\\PycharmProjects\\anils_webscraper\\simon_thompson.xlsx")  # location of file we save to
worksheet = myworkbook.worksheets[0]  # select the first sheet in the excel doc

# register firefox as a browsing application
browser = webdriver.Firefox()
type(browser)  # reveal data type of browser
browser.get(url)  # open url

# click cookie message button
cookie = browser.find_element_by_class_name("o-cookie-message__button")
cookie.click()

# get href links
elements = browser.find_elements_by_class_name("card__link")  # find list of articles to click on
links = [element.get_attribute("href") for element in elements]  # array with article links

browser.quit()  # close browser


# function to scour doc for phrase "value buy"
def save_value_buys(links, workbook, worksheet):
    r = 1  # row of spreadsheet
    c = 1  # column of spreadsheet
    max_articles = 3  # max number of articles you want to save
    for link in links:  # loop through each link
        browser = webdriver.Firefox()
        # paywall avoider add-on
        extension_dir = 'C:\\Users\\richa\\AppData\\Roaming\\Mozilla\\Firefox\\Profiles\\sn8o6vnj.default-release\\extensions\\'
        # remember to include .xpi at the end of your file names
        extension = 'bypasspaywalls@bypasspaywalls.weebly.com.xpi'
        browser.install_addon(extension_dir + extension, temporary=True)

        browser.get(link)  # open article
        article = browser.find_element_by_class_name("article__content").text  # get article text
        article = article.lower()  # make article lower case

        if "strong buy" in article:  # search article for phrase and execute code if present
            browser.quit()
            worksheet.cell(row=r, column=c, value=link)  # save article link in spreadsheet
            workbook.template = False  # internet told me to do this, not sure why
            workbook.save(
                "C:\\Users\\richa\\PycharmProjects\\anils_webscraper\\simon_thompson.xlsx")  # save the spreadsheet
            r += 1  # this will save the next link on the next row in the spreadsheet

            if r == max_articles:
                break  # breaks the loop if we reach the max number of articles

        else:
            # if strong buy is not found, close the browser and go to start of loop to open next link
            browser.quit()


save_value_buys(links, myworkbook, worksheet)  # calls function
