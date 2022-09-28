import cv2 as cv
from PIL import Image
import pyautogui
import subprocess
import win32gui
import os
from urllib.parse import urlparse
from urllib.request import urlopen
import validators
import time
from glob import glob1
#Global Variables
X = 0
Y = 0
XYcoordinates = {}

# pyautogui.FAILSAFE= True
def callApp():
    credit = '=====================================================================\n' \
             '[*][*]     Welcome to File Downloading Application Manager     [*][*]\n' \
             '=====================================================================\n\n' \
             '                    Owner - Sasindu Lakshan\n' \
             '              GitHub - https://github.com/Sasindu0\n'

    with open('log.txt', 'w') as file:
        file.write(credit)
    print(credit)

    chromeFiles = []
    fdmFiles = []
    printLog('[+] Checking connection.. ')
    for i in range(3):
        if not connection():
            checkConnection()
            printLog('[+] Connecting..')
            mouseMove(XYcoordinates['connect.png'][0],XYcoordinates['connect.png'][1])
            time.sleep(5)
            if not connection():
                time.sleep(10)
                continue
            printLog('[+] Connected ')
        else:
            checkConnection()
            printLog('[+] Already connected ')
            break
        if i == 9:
            printLog('[+] Connection can not be establish\n'
                     '[+] Exiting.. ')
            exit()

    isAdd = False
    yesNo = input('[+] Do you want to add Chromium downloads?(y/any) ')
    if yesNo == 'y' or yesNo == 'Y':
        printLog('\n-----------------------------------'
                 '\n[+] Input Chromium download URLs [+]'
                 '\n-----------------------------------')
        chromeFiles = inputUrl()
        isAdd = True

    yesNo = input('[+] Do you want to add fdm downloads?(y/any) ')
    if yesNo == 'y' or yesNo == 'Y':
        printLog('\n-------------------------------'
                 '\n[+] Input FDM download URLs [+]'
                 '\n-------------------------------')
        fdmFiles = inputUrl()
    else:
        if not isAdd:
            printLog('\n[*][*]-- Have a nice day --[*][*]')
            exit()
        printLog('[*] URL adding completed')

    isShutDown = False
    yesNo = input('[+] Do you want shut down the pc after downloads?(y/any) ')
    if yesNo == 'y' or yesNo == 'Y':
        isShutDown = True

    if len(chromeFiles) > 0:
        printLog('\n[+] Chromium file downloading started')
        while True:
            value = isFileExist(chromeFiles)
            if value == 'Disconnect':
                for attempt in range(10):
                    printLog('[+] Connecting..')
                    checkConnection()
                    mouseMove(XYcoordinates['connect.png'][0], XYcoordinates['connect.png'][1])
                    time.sleep(5)
                    if not connection():
                        printLog('[+] Waiting for 10 minutes.. ')
                        time.sleep(60*10)
                        printLog(f'\n[+] Attempt - {attempt+2}')
                        continue
                    else:
                        printLog('[+] Connected')
                        break
            elif value == 'Complete':
                printLog('[*] Chromium file downloading completed\n')
                break

    if len(fdmFiles) > 0:
        progress = checkFdm()
        if progress == 'allPause':
            mouseMove(XYcoordinates['activePlay.png'][0],XYcoordinates['activePlay.png'][1])
            printLog('[+] FDM file downloading started')

            while True:
                value = isFileExist(fdmFiles)
                if value == 'Disconnect':
                    for attempt in range(10):
                        printLog('[+] Connecting..')
                        checkConnection()
                        mouseMove(XYcoordinates['connect.png'][0], XYcoordinates['connect.png'][1])
                        time.sleep(5)
                        if not connection():
                            printLog('[+] Waiting for 10 minutes.. ')
                            time.sleep(60 * 10)
                            printLog(f'\n[+] Attempt - {attempt + 2}')
                            continue
                        else:
                            printLog('[+] Connected')
                            break
                elif value == 'Complete':
                    printLog('[*] FDM file downloading completed\n\n'
                             '============================================\n'
                             '[*]     All files download completed     [*]\n'
                             '============================================\n')
                    break

    if isShutDown:
        isConnect = connection()
        if isConnect:
            checkConnection()
            printLog('[+] Disconnecting..')
            mouseMove(XYcoordinates['disconnect.png'][0],XYcoordinates['disconnect.png'][1])
            time.sleep(5)
            printLog('[+] Disconnected ')
        printLog('[*][*]-- Shutting Down --[*][*]')
        # os.system('shutdown /s /t 10')

def printLog(text):
    print(text)
    writeLog(text)

def writeLog(text):
    with open('log.txt', 'a') as file:
        file.write('\n'+text)

def connection():
    try:
        urlopen('https://www.google.com/')
        return True
    except:
        printLog('[-] (error)Disconnected')
        return False

def screenShot():
    myScreenshot = pyautogui.screenshot()
    myScreenshot.save(r'tests\scrnShot.png')
    return 'tests\scrnShot.png'

def checkFdm():
    subprocess.call('D:\\Program Files\\Softdeluxe\\Free Download Manager\\fdm.exe')
    time.sleep(1)
    scrnShot = screenShot()

    imgList = ['templates\\activePlay.png', 'templates\\inactivePlay.png','templates\\activePause.png', 'templates\\inactivePause.png',
               'templates\\hostNotFound.png']

    resultList = []
    for template in imgList[:4]:
        result = checkTemplate(scrnShot, template)
        print('Result: ',result)
        resultList.append(result)
        templateName = os.path.basename(template)
        XYcoordinates[templateName] = X,Y

    if resultList[0]>resultList[1] and resultList[3]>resultList[2]:
        return 'allPause'
    elif resultList[1]>resultList[0] and resultList[2]>resultList[3]:
        #print('FDM downloading play')
        return 'allPlay'
    elif resultList[1]>resultList[0] and resultList[3]>resultList[2]:
        #print('FDM downloads completed')
        return 'allComplete'
    else:
        result = checkTemplate(scrnShot, imgList[4])
        print('Result: ',result)
        if result == 1.0:
            return 'error'

def checkConnection():
    subprocess.call('C:\\Program Files (x86)\\Mobile Partner\\Mobile Partner.exe')
    time.sleep(1)
    mobilePartner = win32gui.FindWindow(None, 'Mobile Partner')
    win32gui.SetForegroundWindow(mobilePartner)
    scrnShot = screenShot()

    imgList = ['templates\\online.png', 'templates\\offline.png',
               'templates\\connect.png', 'templates\\disconnect.png', 'templates\\cancel.png']
    resultList = []
    for template in imgList:
        result = checkTemplate(scrnShot, template)
        print('Result: ',result)
        resultList.append(result)
        templateName = os.path.basename(template)
        XYcoordinates[templateName] = X, Y

    if resultList[4] == 1.0:
        mouseMove(XYcoordinates['cancel.png'][0], XYcoordinates['cancel.png'][1])


def checkTemplate(scrnShot,template):

    method = cv.TM_SQDIFF_NORMED

    small_image = cv.imread(template)
    large_image = cv.imread(scrnShot)

    result = cv.matchTemplate(small_image, large_image, method)

    mn, _, mnLoc, _ = cv.minMaxLoc(result)
    topx, topy = mnLoc
    trows, tcols = small_image.shape[:2]
    result = cropImg(scrnShot,template,topx,topy,topx+tcols,topy+trows)

    return result

def cropImg(scrnShot,template,topx,topy,bottomx,bottomy):
    global X,Y

    im = Image.open(scrnShot, mode='r')
    image = im.crop((topx, topy, bottomx, bottomy))

    X = topx + (bottomx-topx)/2
    Y = topy + (bottomy - topy) / 2

    image.save('tests\cropImg.png')
    result = checkHisto('tests\cropImg.png',template)

    return result

def checkHisto(image,template):

    base = cv.imread(template)
    test = cv.imread(image)

    hsv_base = cv.cvtColor(base, cv.COLOR_BGR2HSV)
    hsv_test = cv.cvtColor(test, cv.COLOR_BGR2HSV)

    h_bins = 50
    s_bins = 60
    histSize = [h_bins, s_bins]
    h_ranges = [0, 180]
    s_ranges = [0, 256]
    ranges = h_ranges + s_ranges
    channels = [0, 1]

    hist_base = cv.calcHist([hsv_base], channels, None, histSize, ranges, accumulate=False)
    cv.normalize(hist_base, hist_base, alpha=0, beta=1, norm_type=cv.NORM_MINMAX)
    hist_test = cv.calcHist([hsv_test], channels, None, histSize, ranges, accumulate=False)
    cv.normalize(hist_test, hist_test, alpha=0, beta=1, norm_type=cv.NORM_MINMAX)

    compare_method = cv.HISTCMP_CORREL
    base_test = cv.compareHist(hist_base, hist_test, compare_method)

    return base_test

def inputUrl():
    fileNames = []

    while True:
        url = input('URL: ')
        if url != '':
            url = url.replace(' ','%20')
            validate = validators.url(url)
            if validate:
                filePath = urlparse(url)
                fileName = os.path.basename(filePath.path)
                fileName = fileName.replace('%20', ' ')
                fileNames.append(fileName)
                printLog(f'[+] File name \"{fileName}\" listed..')
            else:
                printLog('[*] File Name listing completed')
                break
        else:
            printLog('[*] File Name listing completed')
            break
    return fileNames

def isFileExist(fileNames):
    downloadPath = 'D:\\Disk_4\\Downloads'
    notfile = len(glob1(downloadPath, '*.crdownload')+glob1(downloadPath, '*.fdmdownload'))
    filelen = len(os.listdir(downloadPath)) - notfile
    while True:
        notfile = len(glob1(downloadPath, '*.crdownload') + glob1(downloadPath, '*.fdmdownload'))
        currentlen = len(os.listdir(downloadPath)) - notfile
        isConnect = connection()
        if isConnect:
            if filelen != currentlen:
                for fileName in fileNames:
                    localPath = downloadPath + '\\' + fileName
                    if os.path.exists(localPath):
                        printLog(f'[+] \"{fileName}\" downloading completed.')
                        fileNames.remove(fileName)
                if len(fileNames) == 0:
                    return 'Complete'
            else:
                time.sleep(60)
        else:
            return 'Disconnect'

def mouseMove(X,Y):
    pyautogui.click(x=X, y=Y)
    pyautogui.FAILSAFE = True
    pyautogui.moveTo(0,0)
    pyautogui.FAILSAFE = False

if __name__ == '__main__':
    callApp()