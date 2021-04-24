import os
import time
import wget
import cv2
import _bases
from selenium.webdriver import Edge
from selenium.webdriver.common.by import By
from xlsxwriter import Workbook
import pytesseract


excel_file_name = _bases.getfilepath('dosab')


def getcomponies():
    """
    Get Companies from web and write to excel file
    :return:
    """
    _bases.kill_web_driver_edge()
    driver = Edge()
    componies = []
    driver.get('https://www.dosab.org.tr/Alfabetik-Firmalar-Listesi')

    # Get links
    # links = []
    # datalinks = driver.find_elements(By.XPATH, '/html/body/div[2]/div/ul/li/div/a')
    # for link in datalinks:
    #     linkobj = {
    #         'link': link.get_attribute('href'),
    #         'name': link.text
    #     }
    #     links.append(linkobj)

    # Downlaod Mail Images
    # for complink in componies:
    #     parsedlink = str(complink['link']).split('/')
    #     mailimg = f'https://www.dosab.org.tr/dosyalar/emailler/{parsedlink[4]}_EMail.jpg'
    #     wget.download(mailimg, "imgs")

    # OCR Image to text
    pytesseract.pytesseract.tesseract_cmd = r'C:\Users\abdul\AppData\Local\Tesseract-OCR\tesseract.exe'
    imgfiles = os.listdir('imgs')
    imgfiles.sort()

    for imgfile in imgfiles:
        compid = imgfile.split('_EMail.jpg')[0]
        driver.get(f'https://www.dosab.org.tr/Firma/{compid}')
        compname = driver.find_element(By.XPATH, '/html/body/div[2]/div/div[2]/h4').text
        img = cv2.imread(f'imgs/{imgfile}')
        emailtext = str(pytesseract.image_to_string(img, lang='eng')).replace('\n\f', '')

        if '@' not in emailtext:
            emailtext = ''

        company = {
            'mail': emailtext,
            'name': compname
        }
        componies.append(company)

    workbook = Workbook(excel_file_name)
    worksheet = workbook.add_worksheet('dosab')
    row = 0
    hformat = workbook.add_format()
    hformat.set_bold()
    worksheet.write(row, 0, "Firma Adi", hformat)
    worksheet.write(row, 1, 'Mailler', hformat)
    row += 1

    for comp in componies:
        worksheet.write(row, 0, comp["name"])

        if '@' in comp['mail']:
            worksheet.write(row, 1, comp['mail'])
        row += 1

    workbook.close()

    driver.close()


