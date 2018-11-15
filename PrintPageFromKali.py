from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import win32com.client
autoit = win32com.client.Dispatch("AutoItX3.Control")

driver = webdriver.Chrome(executable_path=r"C:\Selenium\chromedriver_win32\chromedriver.exe")
driver.get('https://tools.kali.org/tools-listing')
element = WebDriverWait(driver, 10).until(EC.title_contains("Penetration Testing Tools"))
all_options = driver.find_elements(By.XPATH, './/div[@class="wpb_wrapper"]//ul[@class="lcp_catlist"]//a')
for option in all_options:
    WebDriverWait(driver, 3)
    title = option.get_attribute("title")
    href = option.get_attribute("href")
    print("title is: %s" % option.get_attribute("title"))
    driver.execute_script("window.open('');")
    # Switch to the new window
    driver.switch_to.window(driver.window_handles[1])
    driver.get(href)
    # close the active tab
    # Switch back to the first tab
    # Close the only tab, will also close the browser.
    #option.click()
    
    if autoit.WinExists("[CLASS:Chrome_WidgetWin_1]"):
        autoit.WinActive("[CLASS:Chrome_WidgetWin_1]")
        #autoit.Send("^!p",0)
        #autoit.ShellExecuteA(window_before,NULL,NULL,NULL,"print")
        #autoit.ControlSend(window_before,"","","^!P")
        #autoit.ControlSend(window_before,"","","{CTRLDOWN}{SHIFTDOWN}p{CTRLUP}{SHIFTUP}")
        autoit.Send("{CTRLDOWN}{SHIFTDOWN}p{CTRLUP}{SHIFTUP}")
        #autoit.Send("{RCTRL}{RSHIFT}p")
        autoit.sleep(2000)
        autoit.Send("{Enter}")
        autoit.sleep(2000)
        autoit.Send(title)
        autoit.Send("{Enter}")
        autoit.sleep(2000)
        driver.close()
        if autoit.WinExists("Confirm Save As"):
            autoit.ControlClick("Confirm Save As", "", "[CLASS:Button; INSTANCE:2]")
            autoit.ControlClick("Save Print Output As","","[CLASS:Button; INSTANCE:3]")
    driver.switch_to.window(driver.window_handles[0])
driver.close()    