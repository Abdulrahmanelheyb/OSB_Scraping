import os
import time
import _bases

from selenium.webdriver import Edge
from selenium.webdriver.common.by import By
from xlsxwriter import Workbook


excel_file_name = _bases.getfilepath('bursaderiosb')


def get_componies():
    """
    Get Companies from web and write to excel file
    :return:
    """
    _bases.kill_web_driver_edge()
    driver = Edge()
    componies = []

    driver.get('https://bursaderiosb.com/uyeler')

    componies_data = driver.find_elements(By.XPATH, '//*[@id="main"]/div/div/div[2]/div')
    compscount = None

    if len(componies_data) > 0:
        compscount = len(componies_data)

    for i in range(1, compscount):
        company_anchor = driver.find_element(By.XPATH, f'//*[@id="main"]/div/div/div[2]/div[{i}]/div[2]/a')
        company_anchor.click()
        company = {
            'Name': driver.find_element(By.XPATH, '//*[@id="main"]/div/div/div[1]').text,
        }
        company_data = str(driver.find_element(By.XPATH, '//*[@id="main"]/div/div/div[2]/div').text).split('\n')
        for compdata in company_data:

            mailstr = 'Mail:'
            if compdata.startswith(mailstr):
                company['Mail'] = compdata.split(mailstr)[1]

            phonestr = 'Telefon:'
            if compdata.startswith(phonestr):
                company['Tel'] = compdata.split(phonestr)[1]

        componies.append(company)

        driver.back()

    # region > Write to excel
    row = 0
    workbook = Workbook(excel_file_name)
    worksheet = workbook.add_worksheet('Bursa Deri OSB')

    hformat = workbook.add_format()
    hformat.set_bold()
    hformat.set_align('center')
    hformat.set_align('vcenter')

    worksheet.write(row, 0, 'Firma Adi', hformat)
    worksheet.write(row, 1, 'Mail Adresi', hformat)
    worksheet.write(row, 2, 'Telefon', hformat)

    worksheet.set_column('A:A', 100)
    worksheet.set_column('B:B', 80)
    worksheet.set_column('C:C', 50)
    row += 1

    for cmpy in componies:
        if 'Name' in cmpy:
            worksheet.write(row, 0, str(cmpy['Name']))

        if 'Mail' in cmpy:
            worksheet.write(row, 1, str(cmpy['Mail']))

        if 'Tel' in cmpy:
            worksheet.write(row, 2, str(cmpy['Tel']))

        row += 1

    if os.path.exists(excel_file_name):
        os.remove(excel_file_name)

    workbook.close()
    # endregion

    time.sleep(3)
    driver.quit()
