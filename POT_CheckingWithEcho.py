from selenium import webdriver
from time import sleep, time
from bs4 import BeautifulSoup
import openpyxl
import re
from bs4.element import NavigableString, Tag
from datetime import datetime, date
import functools


file_name = r"C:\Users\willlee\Downloads\LIST POT 2020.xlsx"


class Excel_Operation:

    def replace_alphabet(self, text):
        return re.sub(r'[a-zA-Z]', r'*', text)

    def __init__(self, file_name, sheet_name):
        self.file_name = file_name
        self.sheet_name = sheet_name

    def load_workbook_with_sheet_name(self):
        self.book = openpyxl.load_workbook(self.file_name)
        self.work_sheet = self.book[self.sheet_name]

    def cell_to_read(self, column_to_read, row_to_start):
        # This function is assumed both read and write having the same row thus just have one row_to_start variable
        read_position = column_to_read + row_to_start  # C2

        text = self.work_sheet[read_position].value
        new_string = self.replace_alphabet(text)
        # replace last character back
        new_string = new_string[:-1] + text[-1]
        print(new_string)
        return new_string, text

    def cell_to_write(self, column_to_write, row_to_start, value_to_write_in):
        write_position = column_to_write + row_to_start  # D2
        self.work_sheet[write_position].value = value_to_write_in

    def save_file(self, file_name):
        self.book.save(file_name)


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
            oneD_array = []
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

    def teardown(self):
        self.driver.close()
        self.driver.quit()
        print("Finished")

# put an decorator for time recording
def logging_time(func):
    @functools.wraps(func)
    def wrapper_time(*args):
        start_time = time()
        func(*args)
        end_time = time()
        run_time = end_time - start_time
        print("--- %s seconds ---" % (time() - start_time))

    return wrapper_time

@logging_time
def main():
    successfully = False
    # ----------Webpage Operation-------------------------------
    webpage_ext = webpage_Extract()
    # ---------Excel Operation---------------------------------
    column_to_read = "C"
    column_to_write = "D"
    row_to_start = "339"
    row_to_end =""
    excel_operation = Excel_Operation(file_name, r"LIST POT 2020")
    excel_operation.load_workbook_with_sheet_name()
    # ---------Webpage Operation-------------------------------
    echo = Control_Echo()  # init webdriver
    # ---------------------------------------------------------
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

    while True:
        print (row_to_start)
        wild_char_text, text = excel_operation.cell_to_read(column_to_read, row_to_start)
        if text == "":
            break
        else:

            successfully = echo.webpage_control(id_name='inputPart', action='write',
                                                value=wild_char_text)
            if successfully:
                print("Part Number Entered")
                successfully = False
            echo.buffer_time(3)

            successfully = echo.webpage_control(id_name='searchButton', action='click')
            if successfully:
                print("Search button Clicked")
                successfully = False

            echo.buffer_time(4)

            # driver = echo.get_driver()
            soup = webpage_ext.webpage_refresh(echo.driver)
            two_d_list = webpage_ext.check_all(0, 5)

            # do what? if it look for most active Check if active same with the text need to take part number to check whether it is active
            for i in two_d_list:
                if i[0] == text:  # word might not correct
                    # take out the active one and do comparison with the excel if same , then write active if not same then write latest one
                    if i[1] == "Active":
                        excel_operation.cell_to_write(column_to_write, row_to_start, "Active")

                        break
                else:  # this is not active one
                    # look for the active
                    if i[1] == "Active" or i[1] == "Standard Support":
                        if i[1] == "Active":
                            statement = "Active part number is " + i[0]
                        if i[1] == "Standard Support":
                            statement = "Standard Support part number is " + i[0]
                        excel_operation.cell_to_write(column_to_write, row_to_start, statement)

            successfully = echo.webpage_control(id_name='clearSearchButton', action='click')
            if successfully:
                print("Clear button Clicked")
                row_to_start = int(row_to_start) + 1
                row_to_start = str(row_to_start)
                excel_operation.save_file(r"C:\Users\willlee\Downloads\LIST POT 2020_2.xlsx") #Save content
                successfully = False
    echo.teardown()


if __name__ == "__main__":
    main()
