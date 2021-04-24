import os


timeout = 3


def getfilepath(filename) -> str:
    outputfilename = f'data/{filename}.xlsx'
    return outputfilename


def kill_web_driver_edge():
    try:
        os.system('taskkill /f /im MicrosoftWebDriver.exe')
    except Exception as ex:
        print(ex)
