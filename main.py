import pandas as pd  # pip install pandas
import pyautogui as pg  # pip install pyautogui
import time
import keyboard  # pip install keyboard
import clipboard  # pip install clipboard
from PIL import ImageGrab #pip install pillow
#import pytesseract #pip install pytesseract
#pip install openpyxl

#pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

mode = "기록" #"정보", "기록", "소견서"

idmx = []


def toexcel(idhos) :
    global datafm

    pg.click(785, 248, clicks=2)
    clipboard.copy(idhos)
    pg.hotkey('ctrl', 'v')
    time.sleep(1)

    pg.click(897,248, clicks=2)
    time.sleep(0.01)
    pg.hotkey('ctrl', 'c')
    name = clipboard.paste()
    time.sleep(0.05)

    pg.click(812,529)
    pg.dragTo(908,529, duration=0.3)
    time.sleep(0.01)
    pg.hotkey('ctrl','c')
    idjumin = clipboard.paste()
    time.sleep(0.05)

    pg.click(873,562)
    pg.dragTo(1158,559, duration=0.3)
    time.sleep(0.01)
    pg.hotkey('ctrl','c')
    juso = clipboard.paste()
    time.sleep(0.05)

    pg.click(812,337)
    pg.dragTo(899,337, duration=0.3)
    time.sleep(0.01)
    pg.hotkey('ctrl','c')
    phone = clipboard.paste()
    time.sleep(0.05)

    savelist = [[idhos, name, idjumin, juso, phone]]

    appdf = pd.DataFrame(data=savelist, columns=['환자번호', '성명', '주민등록번호', '주소', '전화'])
    datafm = datafm.append(appdf)
    print(name, "완료")


def clkpatient(num) :
    pg.click(767, 268+24*(num-1), clicks=2)
    time.sleep(5)

def clkseosik() :
    pg.click(2451,141)
    pg.click(2483,268, clicks=2)

    pg.click(2451, 116)
    pg.click(2483, 242, clicks=2)

    pg.click(2451,88)
    pg.click(2483,214, clicks=2)

    time.sleep(5)

def jdcode() :
    pg.click(3418,849)
    time.sleep(3)
    pg.press('z')
    time.sleep(0.1)
    pg.press('down')
    pg.press('enter')
    time.sleep(1)
    pg.press('enter')
    time.sleep(1)
    pg.click(3407,929)
    time.sleep(2)
    pg.press('enter')
    time.sleep(0.3)
    pg.press('enter')
    time.sleep(2)

def jdcodeOM() :
    pg.click(3418,863)
    time.sleep(3)
    pg.press('d')
    time.sleep(0.1)
    pg.press('down')
    pg.press('enter')
    time.sleep(1)
    pg.press('enter')
    time.sleep(1)

    pg.press('d')
    time.sleep(0.1)
    pg.press('down')
    pg.press('down')
    if ptct == "손목" :
        pg.press('down')
    pg.press('enter')
    time.sleep(1)
    pg.press('enter')
    time.sleep(1)

    pg.click(3407,929)
    time.sleep(2)
    pg.press('enter')
    time.sleep(0.3)
    pg.press('enter')
    time.sleep(2)

def writeEHR() :
    pg.click(3349, 207)
    if ptct == "공무원" :
        clipboard.copy('공무원 채용검진')
    elif ptct == "교원자격증" :
        clipboard.copy('국가면허신청용 검진(교원자격증)')
    pg.hotkey('ctrl', 'v')
    time.sleep(0.3)

    pg.click(3258, 268)
    if ptct == "공무원":
        clipboard.copy('공무원 채용검진')
    elif ptct == "교원자격증":
        clipboard.copy('국가면허신청용 검진(교원자격증)')
    pg.hotkey('ctrl', 'v')
    time.sleep(0.3)

    pg.click(3188, 332)
    clipboard.copy('ns')
    pg.hotkey('ctrl', 'v')
    time.sleep(0.3)

    pg.click(3251, 400)

    for i in range(15):
        pg.scroll(-50)

    pg.click(3199, 619)
    clipboard.copy('vs stable')
    pg.hotkey('ctrl', 'v')
    time.sleep(0.3)

    pg.click(3280,717)
    time.sleep(0.1)
    pg.click(3494, 739)

    jdcode()

    pg.click(3185, 974)
    if ptct == "공무원" :
        clipboard.copy('x-ray, lab')
    elif ptct == "교원자격증" :
        clipboard.copy('TBPE')
    pg.hotkey('ctrl', 'v')
    time.sleep(0.3)


def writeEHROM() :
    pg.click(3348, 173)
    if ptct == "손목" :
        clipboard.copy('직원용 손목 보호대 소견서')
    else :
        clipboard.copy('직원용 스타킹 소견서')
    pg.hotkey('ctrl', 'v')
    time.sleep(0.3)

    pg.click(3267, 221)
    if ptct == "손목" :
        clipboard.copy('손목통증(예방적)')
    else :
        clipboard.copy('하지통증(예방적)')

    pg.hotkey('ctrl', 'v')
    time.sleep(0.3)

    pg.click(3186, 370)
    clipboard.copy('ns')
    pg.hotkey('ctrl', 'v')
    time.sleep(0.3)

    pg.click(3185, 458)
    clipboard.copy('vs stable')
    pg.hotkey('ctrl', 'v')
    time.sleep(0.3)

    pg.click(3254, 515)
    time.sleep(0.5)

    pg.click(3234, 560)
    if ptct == "손목" :
        clipboard.copy('손목')
    else :
        clipboard.copy('하지')
    pg.hotkey('ctrl', 'v')
    time.sleep(0.3)

    for i in range(20):
        pg.scroll(-50)

    pg.click(3538, 332)
    pg.click(3538, 419)
    pg.click(3538, 522)
    pg.click(3273, 666)
    pg.click(3273, 708)
    pg.click(3273, 746)

    jdcodeOM()

    pg.click(3189, 991)
    clipboard.copy('직원용 소견서')
    pg.hotkey('ctrl', 'v')
    time.sleep(0.3)

    for i in range(20):
        pg.scroll(50)

    pg.click(3192, 314)

def clksave() :
    pg.click(3391, 132)

    pg.click(371, 95)

if mode == "정보" :
    print('로딩완료')
    datafm = pd.DataFrame(columns=['환자번호', '성명', '주민등록번호', '주소', '전화'])


    for id in idmx :
        toexcel(id)
    print(datafm)
    datafm.to_excel('C:\\Users\\DELL3040\\Desktop\\datafm.xlsx')


elif mode == "기록" :
    print('alt+s : 교원자격증, alt+d : 공무원')
    i = 0
    '''
    for i in range(5,15) :
        ptct = "공무원"
        clkpatient(i)
        clkseosik()
        writeEHR()
        clksave()
        i = i + 1
    '''
    while True :
        key = keyboard.read_hotkey(suppress=False)

        if key == 'alt+a' :
            print('다음사람')
            i = i+1
            print(i)
            clkpatient(i)
        if key == 'alt+s':
            print('교원자격증')
            ptct = "교원자격증"
            clkseosik()
            writeEHR()
            clksave()
        if key == 'alt+d':
            print('공무원')
            ptct = "공무원"
            clkseosik()
            writeEHR()
            clksave()

        else :
            time.sleep(0.1)

elif mode == "소견서" :
    print('alt+s : 스타킹, alt+d : 손목')
    i = 0

    while True :
        key = keyboard.read_hotkey(suppress=False)

        if key == 'alt+a' :
            i = i+1
            print(i)
            clkpatient(i)

        if key == 'alt+s':
            ptct = "스타킹"
            #clkseosik()
            writeEHROM()

        if key == 'alt+d':
            ptct = "손목"
            #clkseosik()
            writeEHROM()

        if key == 'alt+f':
            clksave()

        else :
            time.sleep(0.1)


