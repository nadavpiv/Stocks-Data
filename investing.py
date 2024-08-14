import helpers
import os
import xlsxwriter
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# take input from the user
file_name = helpers.file_name()
my_stock = helpers.stock_names()
number_years = helpers.number_years()

# open the workbook with the name that the user request
workbook = xlsxwriter.Workbook(file_name + '.xlsx')
workSheet = workbook.add_worksheet()

title_list = ['Name', 'Number years', 'Low yield', 'High yield', 'Years of decline', 'Years of increase',
              'Average']
# change format for the title list
format_title = workbook.add_format({'bold': True, 'align': 'center', 'bg_color': 'green'})
workSheet.write_row(0, 1, title_list, format_title)

# selenium environment preparation
os.environ['PATH'] += r'C:\Program Files (x86)\selenium'

stock_counter = 1
result_list = []
# loop all the stocks
for stock in my_stock:
    print('Start working on ' + stock)
    driver = webdriver.Chrome()
    driver.get('https://www.google.com/')
    driver.set_window_position(-10000, 0)
    helpers.google_click(stock, driver)
    helpers.investing_website_click(driver)
    helpers.selecting_sorting_method(driver)
    helpers.close_ad(driver, stock)

    # variables for retrieving the data on the stock
    counter_years = 0
    count_first_data = 0
    sum_stock = 0
    first = ''
    second = ''
    high_yield = 0
    low_yield = 0
    real_year = int(number_years)
    years_of_decline = 0
    years_of_increase = 0

    while counter_years <= int(number_years):
        helpers.change_dates_click(driver)
        # change the start and the end year for the request of the user
        curr_year = helpers.curr_year(number_years, counter_years)
        helpers.change_start_year(driver, curr_year)
        helpers.change_end_year(driver, curr_year)
        helpers.choose_button_click(driver)

        # get the table data
        table_check = WebDriverWait(driver, 10).until(EC.visibility_of_element_located
                                                      ((By.ID, 'curr_table')))
        table = driver.find_element_by_id('curr_table')
        body_check = WebDriverWait(driver, 10).until(EC.presence_of_element_located
                                                     ((By.TAG_NAME, 'tbody')))
        body = table.find_element_by_tag_name('tbody')
        col_check = WebDriverWait(driver, 10).until(EC.presence_of_element_located
                                                    ((By.TAG_NAME, 'td')))
        col = body.find_elements_by_tag_name('td')

        # extract the data that we need for the average calculate
        table_list = []
        helpers.get_table_data(col, table_list)
        # if the dates are illegal
        check_legal = helpers.check_legal_dates(table_list, stock)
        if check_legal == -1:
            real_year -= 1
            counter_years += 1
            continue

        # get the data for the calculation of the yield percent of the year
        # the first condition happened if the first year of the data is with just one month
        if count_first_data == 0 and len(table_list) < 8:
            second = table_list[len(table_list) - helpers.DIFF_FOR_STOCK_FOR_THE_FIRST_DATE]
            first = table_list[1]
        elif len(table_list) < 8:
            real_year -= 1
            counter_years += 1
            continue
        elif count_first_data == 0:
            second = table_list[len(table_list) - helpers.DIFF_FOR_STOCK_FOR_THE_FIRST_DATE]
            first = table_list[helpers.CURR_VALUE_STOCK]
        elif counter_years == int(number_years):
            second = first
            first = table_list[1]
        else:
            second = first
            first = table_list[helpers.CURR_VALUE_STOCK]
        counter_years += 1
        count_first_data += 1

        # change the data for float mode
        first_num = float(first.replace(",", ""))
        second_num = float(second.replace(",", ""))

        # calculate the result yield percent of the year
        result_year = helpers.calculate_year_yield_percent(first_num, second_num)
        if result_year > 0:
            years_of_increase += 1
        elif result_year < 0:
            years_of_decline += 1

        # check if there is a change in the max or min yield
        high_yield = max(high_yield, result_year)
        low_yield = min(low_yield, result_year)
        sum_stock += result_year

    # format the floats for 2 numbers after the point
    format_average = "{:.2f}".format(sum_stock / (int(real_year) + 1))
    format_high = "{:.2f}".format(high_yield)
    format_low = "{:.2f}".format(low_yield)

    # create the option list for the workbook
    # change the format of the workbook for the data
    cell_format1 = workbook.add_format()
    cell_format1.set_align('center')
    new_list = [stock.upper(), real_year + 1, format_low + '%', format_high + '%', years_of_decline,
                years_of_increase, float(format_average)]
    workSheet.write_row(counter_years, 1, new_list, cell_format1)
    stock_counter += 1
    driver.quit()


# write title number for every stock, from 1 to the last stock
format_title_yellow = workbook.add_format({'bold': True, 'align': 'center', 'bg_color': 'yellow'})
number = 1
for num in range(len(result_list)):
    workSheet.write(number, 0, number, format_title_yellow)
    number += 1

workbook.close()
