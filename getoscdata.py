import webbrowser
import pyautogui
from selenium import webdriver
pyautogui.PAUSE = 1
pyautogui.FAILSAFE = True
browser = webdriver.Firefox()
browser.get('http://osc.orioautoparts.com/')
userElem = browser.find_element_by_id('topname')
userElem.clear()
userElem.send_keys('patricia')
pwElem = browser.find_element_by_id('toppassword')
pwElem.clear()
pwElem.send_keys('SittikulOAP2016')
pwElem.submit()
browser.find_element_by_partial_link_text('View as Company').click()
allOSC = browser.find_elements_by_css_selector('[data-site="OSC"]')
for osc in allOSC:
    
