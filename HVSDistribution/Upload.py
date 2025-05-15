
""" Files """
import Import

""" Libraries """
import time
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
# Chrome
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.remote.webelement import WebElement

# Other
import dotenv
import os

#Other
import sys
sys.path.append(os.getcwd())
sys.path.append(f"{os.getcwd()}/main_automation_programs")
from main_automation_programs import tools

def get_date():
    import datetime
    today = datetime.date.today().strftime("%d %b %Y")
    return today

def setDistribution(width, height): 
    # Import.distribution = {12345: [10369938501, 10484075401], 54321: [10725847701]}

    try: 
        chrome_options = webdriver.ChromeOptions()
        chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
        chrome_options.add_experimental_option("detach", True)
        chrome_options.add_argument(rf"user-data-dir={tools.get_userlogindatapath()}") #helps remember recent logins, eliminates need to reverify okta

        driver_path: str = ''
        try:
            driver_path = ChromeDriverManager().install()
            driver: webdriver.Chrome = webdriver.Chrome(service=ChromeService(driver_path), options=chrome_options)
            print(f"automatic chromedriver path: '{driver_path}'")
        except:
            print("Failed to get chromedriver, initiating automatic install")
            driver_path: str = tools.get_chromedriver() #calls automatic chromedriver fetch
            driver: webdriver.Chrome = webdriver.Chrome(service=ChromeService(driver_path))
        driver.set_window_size(width, height)

        #Maximizing browser window
        maximized = False
        while not maximized:
            try:
                driver.maximize_window()
            except:
                continue
            else:
                maximized = True

        driver.get("https://ccbs-portal-ui-prod.app.clecc3-az1.paas.fedex.com/login") #getting main ccbs page

        # Start CCBS Login 
        ccbsLogin: WebElement = WebDriverWait(driver, float('inf')).until(
                EC.presence_of_element_located((By.XPATH, """//*[@id="bg-container"]/div/div[2]/div[2]/div/a/button"""))
        )
        # time.sleep(3)
        ccbsLogin.click()

        #Automated okta login
        dotenv.load_dotenv(dotenv_path=dotenv.find_dotenv(), override=True)
        if os.environ['okta_automation'] == '1': #only automate when user selected so
            try:
                dotenv.load_dotenv(override=True) #Loading credentials
                tools.okta_login(driver=driver, target_xpath="""//*[@id="applications-list-container"]/div/div[2]/a/mat-card/mat-card-content""",
                                    username=os.environ['okta_username'], pssword=os.environ['okta_password'])
            except Exception as e:
                print(f'Failed to automate okta login, please set login credentials. Exception: {e}')
        
    
        # Press on Classify btn 
        classifyBtn: WebElement = WebDriverWait(driver, float('inf')).until(
                EC.presence_of_element_located((By.XPATH, """//*[@id="applications-list-container"]/div/div[2]/a/mat-card/mat-card-content"""))
        )
        time.sleep(3)
        classifyBtn.click()
        
        home_dnt_handle: str = driver.current_window_handle #saving current window handle

        WebDriverWait(driver, float('inf')).until(EC.number_of_windows_to_be(2))  # Wait for the new window to open
        for window_handle in driver.window_handles:
            if window_handle != home_dnt_handle:
                #Will first look for elements, to make sure page has finished loading before removing new window
                driver.switch_to.window(window_handle)
                WebDriverWait(driver, float('inf')).until(#Waiting for elements to load in new Classify window
                    EC.presence_of_element_located((By.ID, """classifyAdvancedFilter"""))
                )

                #Once loaded, safe to close old window
                driver.switch_to.window(home_dnt_handle) #switching to original window handle
                driver.close() #closing original dnt window
                driver.switch_to.window(window_handle) #switching to new window handle

        # Advanced Filter Dropdown
        filterDropdown = WebDriverWait(driver, float('inf')).until(
            EC.presence_of_element_located((By.ID, """classifyAdvancedFilter"""))
        )
        select = Select(filterDropdown)
        select.select_by_visible_text("AWB Search")

        # time.sleep(5)

        n_brokers = len(Import.distribution)
        for broker, awbs in Import.distribution.items(): #iterating brokers to process
            n_brokers -= 1
            print("Broker: ", broker)
            print("Awbs: ", awbs)

            # Search Button
            searchBtn = WebDriverWait(driver, float('inf')).until(
                EC.presence_of_element_located((By.XPATH, """/html/body/div[1]/div[2]/div[1]/div/form/div/div[2]/div/span[1]/input"""))
            )
            searchBtn.click()

            awb_str = "\n".join(map(str, awbs))
            awbText = WebDriverWait(driver, float('inf')).until(
                EC.presence_of_element_located((By.XPATH, """/html/body/div[3]/div[2]/form/div/b/b/fieldset/div/span[3]/textarea"""))
            )
            awbText.clear()
            awbText.send_keys(awb_str) 

            submitBtn = WebDriverWait(driver, float('inf')).until(
                EC.presence_of_element_located((By.XPATH, """/html/body/div[3]/div[2]/form/b/b/div[1]/span/button[2]"""))
            )
            submitBtn.click()

            for i in range(len(awbs)):

                # Press Next Button
                if i != 0:  
                    nextBtn = WebDriverWait(driver, float('inf')).until(
                        EC.presence_of_element_located((By.ID, """shipmentListNextButton"""))
                    )
                    nextBtn.click()
                
                time.sleep(0.5)

                # Action dropdown
                actionDropdown = WebDriverWait(driver, float('inf')).until(
                    EC.presence_of_element_located((By.ID, """navBarActionSelect"""))
                )
                select2 = Select(actionDropdown)
                select2.select_by_visible_text("Add Comment")

                time.sleep(0.5)

                # Comment Option dropdown
                commentDropdown = WebDriverWait(driver, float('inf')).until(
                    EC.presence_of_element_located((By.ID, """classifyAddComment"""))
                )
                select2 = Select(commentDropdown)
                select2.select_by_visible_text("FREE FORM")

                time.sleep(0.5)

                # Comment input
                assignment_str = "AWB ASSIGNED TO " + str(broker)
                awbText = WebDriverWait(driver, float('inf')).until(
                    EC.presence_of_element_located((By.ID, """addCommentDesc"""))
                )
                awbText.clear()
                awbText.send_keys(assignment_str) 

                time.sleep(0.5)

                # Add Comment button
                addCommentBtn = WebDriverWait(driver, float('inf')).until(
                    EC.presence_of_element_located((By.XPATH, """/html/body/div[3]/div[2]/div[2]/button[2]"""))
                )
                addCommentBtn.click()

                time.sleep(1)

                # acknowledge Comment button
                okBtn = WebDriverWait(driver, float('inf')).until(
                    EC.presence_of_element_located((By.XPATH, """/html/body/div[3]/div[2]/div[2]/button"""))
                )
                okBtn.click()

                time.sleep(0.5)
            print(f"Brokers left to process: {n_brokers}\n")


    except Exception as e:
        print('Exception: ', e)
        user = ""
        #trying to find user for exception, when user provided their ID
        try:
            dotenv.load_dotenv(dotenv_path=dotenv.find_dotenv(), override=True)
            user = os.environ['okta_username']
        except:
            user = "unknown"