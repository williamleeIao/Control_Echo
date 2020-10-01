from selenium import webdriver
from time import sleep, time
from bs4 import BeautifulSoup
import string
from bs4.element import NavigableString, Tag
from datetime import datetime, date
import os


class webpage_Extract:
    # load the library from website
    def webpage_refresh(self, driver):
        sleep(3)
        page_source = driver.page_source
        sleep(3)
        soup = BeautifulSoup(page_source, 'lxml')
        self.soup = soup
        self.driver = driver
        return soup

    def check_all(self, *column_to_check):
        twoD_array = []
        oneD_array = []
        print(column_to_check)
        # convert to list
        list_column_to_check = list(column_to_check)
        items = self.soup.find_all('div', id='contenttablejqxGrid')
        for item in range(len(items[0].contents)):  # <--- This is the 25 elements
            print(items[0].contents[item])  # <--- This is 1 of the 25 element
            # go into contents
            # check if  text is empty if yes, break the loop
            if items[0].contents[item].contents[0].text == "":
                break
            else:
                for individual in list_column_to_check:
                    print(items[0].contents[item].contents[individual].text)
                    oneD_array.append(items[0].contents[item].contents[individual].text)
                twoD_array.append(oneD_array)
            # reset back 1D array
            oneD_array=[]
        return twoD_array


class Control_Echo:

    def __init__(self):
        self.driver = webdriver.Firefox()
        print('init')

    def launch_address(self, address):
        self.driver.get(address)
        print("open " + address)

    def webpage_control(self, id_name, action, value=""):

        var = self.driver.find_element_by_id(id_name)
        if action == "click":
            var.click()
            return True

        if action == "write":
            var.send_keys(value)
            return True

        else:
            print("match nothing")
            return False

    def buffer_time(self, value):
        sleep(value)

    @property
    def get_driver(self):
        print("get _driver ")
        return self.driver


def main():
    successfully = False
    webpage_ext = webpage_Extract()
    echo = Control_Echo()  # init webdriver
    echo.launch_address('https://echo.natinst.com/')
    echo.buffer_time(10)
    print('opened Echo')
    successfully = echo.webpage_control(id_name='i0116', action='write',
                                        value='william.lee@ni.com')
    if successfully:
        print("Email Id Entered")
        successfully = False
    echo.buffer_time(5)

    successfully = echo.webpage_control(id_name='idSIButton9', action='click')
    if successfully:
        print("Button Pressed")
        successfully = False
    echo.buffer_time(5)

    print("User need to enter email and password id at the message box due to no way to control")
    input("Once Enter , press any key to continue")
    echo.buffer_time(10)

    successfully = echo.webpage_control(id_name='inputPart', action='write',
                                        value='159572*-000L')
    if successfully:
        print("Part Number Entered")
        successfully = False
    echo.buffer_time(5)

    successfully = echo.webpage_control(id_name='searchButton', action='click')
    if successfully:
        print("Search button Clicked")
        successfully = False

    echo.buffer_time(5)

    # driver = echo.get_driver()
    soup = webpage_ext.webpage_refresh(echo.driver)
    webpage_ext.check_all(0, 5)


if __name__ == "__main__":
    main()
