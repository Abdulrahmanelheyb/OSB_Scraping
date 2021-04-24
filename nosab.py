import os
import time
import _bases
from selenium.webdriver import Edge
from selenium.webdriver.common.by import By
from xlsxwriter import Workbook


excel_file_name = _bases.getfilepath('nosab')


def getcomponies():
    """
    Get Companies from web and write to excel file
    :return:
    """
    _bases.kill_web_driver_edge()
    driver = Edge()
    componies = []

    driver.get('https://www.nosab.org.tr/firmalar/tr')
    alphabetslinks = []

    for links in driver.find_elements(By.XPATH, '//*[@id="accordion-2"]/li/a'):
        link = {
            'Sector': links.text,
            'Name': links.get_attribute('href')
        }
        alphabetslinks.append(link)

    for anchor in alphabetslinks:
        driver.get(anchor['Name'])
        companies_sector = {
            'Sector': anchor['Sector'],
            'comps': []
        }

        componies_count = len(driver.find_elements(By.XPATH, '/html/body/div[7]/div/div[2]/div[3]/ul/li/a'))

        for indx in range(1, componies_count + 1):
            comp = driver.find_element(By.XPATH, f'/html/body/div[7]/div/div[2]/div[3]/ul/li[{indx}]/a')
            comp.click()
            companies_sector['Sector'] = anchor['Sector']
            company = {
                'Name': driver.find_element(By.XPATH, '/html/body/div[7]/div/div[2]/div[1]/div').text,
                'Data': str(driver.find_element(By.XPATH, '/html/body/div[7]/div/div[2]/div[4]').text)
            }

            companies_sector['comps'].append(company)
            driver.back()

        componies.append(companies_sector)

    row = 0
    workbook = Workbook(excel_file_name)
    worksheet = workbook.add_worksheet('nosab')

    hformat = workbook.add_format()
    hformat.set_bold()
    hformat.set_align('center')
    hformat.set_align('vcenter')

    worksheet.write(row, 0, 'Firma Adi', hformat)
    worksheet.set_column('A:A', 100)

    worksheet.write(row, 1, 'Bilgileri', hformat)
    worksheet.set_column('B:B', 120)

    row += 1

    fwarp = workbook.add_format()
    fwarp.set_text_wrap()

    fname_centralize = workbook.add_format()
    fname_centralize.set_align('center')

    for company in componies:

        if 'Sector' in company:
            worksheet.write(row, 0, company['Sector'], hformat)
            row += 1

        if 'comps' in company:
            for comp in company['comps']:

                if 'Name' in comp:
                    worksheet.write(row, 0, comp['Name'], fname_centralize)

                if 'Data' in comp:
                    worksheet.write(row, 1, comp['Data'], fwarp)

                row += 1

    if os.path.exists(excel_file_name):
        os.remove(excel_file_name)

    time.sleep(_bases.timeout)
    workbook.close()
    driver.close()
