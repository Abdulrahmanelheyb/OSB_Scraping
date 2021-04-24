import os
import time
from selenium.webdriver import Edge
from selenium.webdriver.common.by import By
from xlsxwriter import Workbook
import _bases

excel_file_name = _bases.getfilepath('kosab')


def get_componies():
    _bases.kill_web_driver_edge()
    driver = Edge()
    componies = []
    # region > info type headers variables
    sector = 'Faaliyet Alanı '
    mail = 'Epostalar '
    phone = 'Telefonlar '
    # endregion

    driver.get('http://www.kosab.org.tr/FIRMALAR/')

    # Get table rows count
    pagination_count = len(driver.find_elements(By.XPATH, '/html/body/div[2]/div/div/span[2]/table/tbody/tr'))
    pagination_links = []

    anchors = driver.find_elements(By.XPATH, f'/html/body/div[2]/div/div/span[2]/table/tbody/tr[{pagination_count}]/td/a')
    for pagelink in anchors:
        pagination_links.append(pagelink.get_attribute('href'))

    for page in pagination_links:
        driver.get(page)

        def get_comps_row():
            compoinesrows = driver.find_elements(By.XPATH, '/html/body/div[2]/div/div/span[2]/table/tbody/tr')
            return compoinesrows

        for row in range(1, len(get_comps_row()) - 1):
            company = {}
            compsrow = get_comps_row()
            compsrow[row].click()

            datatable = driver.find_elements(By.XPATH, '/html/body/div[2]/div/div/span[2]/table/tbody/tr')

            for cell in range(0, len(datatable) - 1):

                datastr = str(datatable[cell].text)

                if cell == 0:
                    company['Name'] = datatable[cell].text

                if datastr.startswith(sector):
                    company['Sector'] = str(datatable[cell].text).split(sector)[1]

                if datastr.startswith(mail):
                    company['Mail'] = str(datatable[cell].text).split(mail)[1]

                if datastr.startswith(phone):
                    company['Tel'] = str(datatable[cell].text).split(phone)[1]

            componies.append(company)
            driver.back()

    # region > Write excel
    row = 0
    workbook = Workbook(excel_file_name)
    worksheet = workbook.add_worksheet('Kosab')

    hformat = workbook.add_format()
    hformat.set_bold()
    hformat.set_align('center')
    hformat.set_align('vcenter')
    hformat.set_font_color('white')
    hformat.set_bg_color('blue')

    worksheet.write(0, 0, 'Firma Adi', hformat)
    worksheet.write(0, 1, 'Mail Adresi', hformat)
    worksheet.write(0, 2, 'Sektor', hformat)
    worksheet.write(0, 3, 'Telefon', hformat)

    worksheet.set_column('A:A', 70)
    worksheet.set_column('B:B', 40)
    worksheet.set_column('C:C', 40)
    worksheet.set_column('D:D', 30)
    row += 1

    for cmpy in componies:

        if 'Name' in cmpy:
            worksheet.write(row, 0, str(cmpy['Name']))

        if 'Mail' in cmpy:
            worksheet.write(row, 1, str(cmpy['Mail']))

        if 'Sector' in cmpy:
            worksheet.write(row, 2, str(cmpy['Sector']))

        if 'Tel' in cmpy:
            worksheet.write(row, 3, str(cmpy['Tel']))

        row += 1

    if os.path.exists(excel_file_name):
        os.remove(excel_file_name)

    time.sleep(2)
    workbook.close()
    # endregion
