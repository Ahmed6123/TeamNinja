import requests
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.firefox.service import Service as FirefoxService

from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as when
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys

import pause; import os; import re
import time; from datetime import datetime
import colorama; from termcolor import colored

from webdriver_manager.chrome import ChromeDriverManager
from webdriver_manager.firefox import GeckoDriverManager

colorama.init()

###################################################################
#                        Teams Meet             HH:MM:SS DD/MM/YYYY
MEETS = {
    "1https://teams.microsoft.com/l/meetup-join/19:1mBIxfF6WIZ_supXSas7J9NBhl_UF4qETXbO3s-FDpU1@thread.tacv2/1717520576099?context=%7b%22Tid%22%3a%2275df096c-8b72-48e4-9b91-cbf79d87ee3a%22%2c%22Oid%22%3a%221504c83d-fde9-407a-8724-871d52b494a7%22%7d": "23:59:59 10/06/2023",
         "2 https://meet.google.com/meetURL2": "23:59:59 10/06/2024",
         "3 https://meet.google.com/meetURL3": "23:59:59 10/06/2024",
         "4 https://meet.google.com/meetURL4": "23:59:59 10/06/2024",
         # Add more Teams URLs (if any) using the same format as above
         }

DURATION = 60 # Duration of each Meetng in minutes
USERNAME = "01-131222-008@student.bahria.edu.pk"
PASSWORD = "tasNas32$"

#                   Google Chrome
#           Linux: "ChromeDrivers/linux64/chromedriver"
#             Mac: "ChromeDrivers/mac64/chromedriver"
#        Mac (M1): "ChromeDrivers/mac64_m1/chromedriver"
#         Windows: "ChromeDrivers/win32/chromedriver.exe"

#                   Mozilla Firefox
#     Linux (x32): "FirefoxDrivers/linux32/geckodriver"
#     Linux (x64): "FirefoxDrivers/linux64/geckodriver"
#             Mac: "FirefoxDrivers/mac64/geckodriver"
#   Windows (x32): "FirefoxDrivers/win32/geckodriver.exe"
#   Windows (x64): "FirefoxDrivers/win64/geckodriver.exe"
##################################################################

# All required interactive elements' locators (text fields, buttons, etc.)
usernameFieldPath = "i0116"
nextButtonPath="idSIButton9"
passwordFieldPath = "i0118"
switchToNewTeamsButtonPath=".app-switcher-install-by-policy-dialog--btn.ts-btn.ts-btn-fluent.ts-btn-fluent-primary"
joinButton1Path = "[aria-label='Join now']"
joinButton2Path = "//span[contains(text(), 'Ask to join')]"
endButtonPath = "[aria-label='Leave (Ctrl+Shift+H)']"
ToggleMuteButtonPath = "[data-tid='toggle-mute']"
ContinueinBrowserButtonPath = "[aria-label='Join meeting from this browser']"
JoinWithoutAudioVideoButtonPath = ".fui-Button.r1alrhcs"

BANNER1 = colored('''
  ▄▄▄█████▓▓▓█████ ▀▄▄▄       ███▄ ▄███ ███▄    █  ██▓ ███▄    █  ▄▄▄██▀▀▀▄▄▄
  ▓  ██▒ ▓▒▒▓█   ▀ ▒████▄    ▓██▒▀█▀ ██ ██ ▀█   █ ▓██▒ ██ ▀█   █    ▒██  ▒████▄
  ▒ ▓██░ ▒░░▒███   ▒██  ▀█▄  ▓██    ▓██▓██  ▀█ ██▒▒██▒▓██  ▀█ ██▒   ░██  ▒██  ▀█▄
  ░ ▓██▓ ░ ▒▓█   ▄ ░██▄▄▄▄██ ▒██    ▒██▓██▒  ▐▌██▒░██░▓██▒  ▐▌██▒▓██▄██▓ ░██▄▄▄▄██
  ▒ ▒██▒ ░ ▒░▒████▒ ▓█   ▓██▒▒██▒   ░██▒██░   ▓██░░██░▒██░   ▓██░ ▓███▒   ▓█   ▓██▒
  ░ ▒ ░░   ░░░ ▒░ ░ ▒▒   ▓▒█░░ ▒░   ░  ▒░   ▒ ▒ ░▓  ░ ▒░   ▒ ▒  ▒▓▒▒░   ▒▒   ▓▒█░
  ░   ░    ░ ░ ░  ░  ▒   ▒▒ ░░  ░      ░░   ░ ▒░ ▒ ░░ ░░   ░ ▒░ ▒ ░▒░    ▒   ▒▒ ░
    ░          ░     ░   ▒   ░      ░   ░   ░ ░  ▒ ░   ░   ░ ░  ░ ░ ░    ░          
  ░            ░  ░      ░  ░       ░   ░  ░           ░  ░   ░        ░  ░''', 'blue')
BANNER2 = colored('''                    ------------------------------------''', 'blue')
BANNER3 = colored('''                    || TeamNinja: The Microsoft Teams Bot ||''', 'red')
BANNER4 = colored('''                    ------------------------------------''', 'blue')

def printBanner():
    print(BANNER1), print(BANNER2), print(BANNER3), print(BANNER4)


def fixTimeFormat(rawTime):
    rawTime = list(rawTime.split())
    times = list(map(int, rawTime[0].split(":")))
    dates = list(map(int, reversed(rawTime[1].split("/"))))
    startTime = dates + times
    return startTime


def timeStamp():
    timeNow = str(datetime.now())
    timeRegEx = re.findall(r"([0-9]+:[0-9]+:[0-9]+)", timeNow)
    return(timeRegEx[0])


def initBrowser():
    BrowserChoice=input("Enter Browser of Your Choice:\n1. Google Chrome\n2. Mozilla Firefox\n\n")
    print("\nInitializing browser...", end="")
    match BrowserChoice:
        case "1":
            chromeOptions = webdriver.ChromeOptions()
            chromeOptions.add_argument("--disable-infobars")
            chromeOptions.add_argument("--disable-gpu")
            chromeOptions.add_argument("--disable-extensions")
            chromeOptions.add_argument("--window-size=800,800")
            chromeOptions.add_argument("--incognito")
            chromeOptions.add_experimental_option('excludeSwitches', ['enable-logging'])
            chromeOptions.add_experimental_option("prefs", {"profile.default_content_setting_values.media_stream_mic": 2,
                                                            "profile.default_content_setting_values.media_stream_camera": 2,
                                                            "profile.default_content_setting_values.notifications": 2,
                                                            "hardware.audio_capture_allowed_urls": ["https://teams.microsoft.com"],
                                                            "hardware.video_capture_allowed_urls": ["https://teams.microsoft.com"]
                                                            })
            chrome_service = ChromeService(ChromeDriverManager().install())
            driver = webdriver.Chrome(service=chrome_service, options=chromeOptions)
            # else:
            #     servicePath = Service(BROWSER_DRIVER)
            #     driver = webdriver.Chrome(service=servicePath, options=chromeOptions)
        case "2":
            firefoxOptions = webdriver.FirefoxOptions()
            firefoxOptions.add_argument("--width=800"), firefoxOptions.add_argument("--height=800")
            firefoxOptions.set_preference("network.cookie.cookieBehavior" , 0)
            #firefoxOptions.set_preference("network.cookie.alwaysAccept" , True)
            firefoxOptions.headless = True
            firefoxOptions.set_preference("layers.acceleration.disabled", True)
            firefoxOptions.set_preference("browser.privatebrowsing.autostart", True)
            firefoxOptions.set_preference("permissions.default.microphone", 2)
            firefoxOptions.set_preference("permissions.default.camera", 2)
            firefox_service = FirefoxService(GeckoDriverManager().install())
            sdriver = webdriver.Firefox(service=firefox_service, options=firefoxOptions)
            # else:
            #     servicePath = Service(BROWSER_DRIVER)
            #     driver = webdriver.Firefox(service=servicePath, options=firefoxOptions)
        case _:
            print("Invalid Choice")
    print(colored(" Success!", "green"))
    return(driver)


def login():
    print("Logging into Microsoft account...", end="")
    driver.get('https://teams.microsoft.com/v2/')

    time.sleep(5)
    usernameField = wait.until(when.element_to_be_clickable((By.ID, usernameFieldPath)))
    time.sleep(1)
    usernameField.send_keys(USERNAME)
    usernameNextButton = wait.until(when.element_to_be_clickable((By.ID, nextButtonPath)))
    usernameNextButton.click()

    passwordField = wait.until(when.element_to_be_clickable((By.ID, passwordFieldPath)))
    time.sleep(3)
    passwordField.send_keys(PASSWORD)
    passwordNextButton = wait.until(when.element_to_be_clickable((By.ID, nextButtonPath)))
    passwordNextButton.click()

    StaySignInButton = wait.until(when.element_to_be_clickable((By.ID, nextButtonPath)))
    time.sleep(1)
    StaySignInButton.click()
    time.sleep(15)
    driver.get('https://teams.microsoft.com')
    print("\nWaiting to Switch to New Teams v2")
    time.sleep(20)    #Just Teams Things
    try:
        switchToNewTeamsButton = wait.until(when.element_to_be_clickable((By.CSS_SELECTOR,switchToNewTeamsButtonPath)))
        switchToNewTeamsButton.click()
        print("\nSwitching To New Meets v2")
    except:
        pass
    print(colored(" Success!", "green"))


def attendMeet():
    time.sleep(10)
    print(f"\n\nNavigating to Google Meet #{meetIndex}...", end="")
    driver.get(URL[1:])
    print(colored(" Success!", "green"))
    print(f"Entering Google Meet #{meetIndex}...", end="")
    time.sleep(5)
    # action.send_keys(Keys.ESCAPE).perform()
    try:
        ContinueinBrowserButton=wait.until(when.element_to_be_clickable((By.CSS_SELECTOR, ContinueinBrowserButtonPath)))
        ContinueinBrowserButton.click()
    except:
        pass
    time.sleep(10)
    # try:
    #     JoinWithoutAudioVideoButton=wait.until(when.element_to_be_clickable((By.CLASS_NAME, JoinWithoutAudioVideoButtonPath)))
    #     print(JoinWithoutAudioVideoButton)
    #     MyButton = JoinWithoutAudioVideoButton[11]
    #     print(MyButton)
    #     MyButton.click()
    # except:
    #     print(e)

    try:
        ToggleMuteButton = wait.until(when.element_to_be_clickable((By.CSS_SELECTOR, ToggleMuteButtonPath)))
        time.sleep(1)
        ToggleMuteButton.click()
    except:
        pass

    joinButton = wait.until(when.element_to_be_clickable((By.CSS_SELECTOR, joinButton1Path)))
    time.sleep(1)
    
    joinButton.click()

    print(colored(" Success!", "green"))
    time.sleep(1)
    print(colored(f"Now attending Google Meet #{meetIndex} @{timeStamp()}", "green"), end="")


def endMeet():
    endButton = wait.until(when.element_to_be_clickable((By.CSS_SELECTOR, endButtonPath)))  
    endButton.click()
    print(colored(f"\nSuccessfully ended Google Meet #{meetIndex} @{timeStamp()}\n", "red"), end="")
    time.sleep(5)
    try:
        driver.quit()
    except:
        pass


def clrscr():
    if os.name == 'posix':
        _ = os.system('clear')
    else:
        _ = os.system('cls')
    printBanner()


def hibernate():
    print("\nHibernating in 10 seconds. Press Ctrl + C to abort.")
    time.sleep(13)
    _ = os.system('shutdown /h /f')


############### Main ###############

if __name__ == "__main__":

    printBanner()
    try:
        DURATION *= 1
        driver = initBrowser()
        wait = webdriver.support.ui.WebDriverWait(driver, 7)
        action = ActionChains(driver)
        for meetIndex, (URL, rawTime) in enumerate(MEETS.items(), start=1):
            startTime = fixTimeFormat(rawTime)
            if (meetIndex <= 1):
                print(colored(f"Waiting until first Meet start time [{rawTime}]...", "yellow"), end="")
            else:
                print(colored(f"\n\nWaiting until next Meet start time [{rawTime}]...", "yellow"), end="")
            pause.until(datetime(*startTime))
            print(colored(" Started!", "green"))
            if (meetIndex <= 1):
                login()
            attendMeet()
            time.sleep(DURATION)
            endMeet()
        print("\n\nAll Meets completed successfully.")
        # hibernate()
        # Uncomment above to hibernate after a 10 second countdown upon completion of all Meets (Ctrl + C to abort hibernation)
        print("Press Enter to exit.")
        input()
        print("\nCleaning up and exiting...", end="")
        driver.quit()

    except KeyboardInterrupt:
        # clrscr()
        print("\n\nCTRL ^C\n\nThrew a wrench in the works.")
        print("Press Enter to exit.")
        input()
        print("\nCleaning up and exiting...", end="")
        driver.quit()
    except:
        print(e)
        # Uncomment above to display error traceback (use when reporting issues)