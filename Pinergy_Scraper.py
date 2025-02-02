"""
Author: Mike
Desc:
- Scrape the MA MLS Database for Realtor Emails
- Scrape by Office not Agent because some cities have...
... too many agents and the results max at 1,000
- Do NOT have driver-browser in use when running this as MLS...
... seems to choose between opening as New Tab, New Window or Self
- Script only grabs first person under office-admin because...
... the HTML code changes depending on one person or multiple
... and, with them being 'admin', I do not feel it to be overly
... important to collect their info
"""

# Imports
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
import time
import sys
import re
import csv
import pandas as pd
import os

# Function: Read in the Zip-code CSV
def load_csv_to_list(filename):
    with open(filename, 'r') as file:
        csv_reader = csv.reader(file)
        data = list(csv_reader)
    return data

# Function: Sign-out, quit the selenium driver and exit the terminal
def end_session(active_driver):
    # Sign out of MLSPIN
    sign_out_button = active_driver.find_element(By.CLASS_NAME, "signout-link")
    sign_out_button.click()

    # Kill the driver
    active_driver.quit()

    # End the terminal
    sys.exit()

# STAGE ONE: Pull zip-codes into memory
zip_codes = load_csv_to_list("MA-Zip-Codes.csv")

# STAGE TWO: Initiate browser and initiate first search
# Open the browser to the URL (in full-screen)
print("Stage 01 initiating...")
driver = webdriver.Firefox()
driver.maximize_window()
web_address = "https://pinergy.mlspin.com/"
driver.get(web_address)


# STEP 01
# Wait-for/Click-on Cookie Consent Button
CLASS_cookie_modal_button = "mls-js-cookie-consent-action"
try:
    el = WebDriverWait(driver, 15).until(
        EC.visibility_of_element_located((By.CLASS_NAME, CLASS_cookie_modal_button))
        )
    el.click()
except:
    print("Error: Could not find Cookie Consent Button.")
    driver.quit()
    sys.exit()

print("Step 01... complete.")
time.sleep(3)

# STEP 02
# Complete MLSPIN login form
username_field = driver.find_element(By.NAME, "user_name")
username_field.send_keys("UPDATE_ME")

password_field = driver.find_element(By.NAME, "pass")
password_field.send_keys("UPDATE_ME")

signin_button = driver.find_element(By.CLASS_NAME, "mls-js-submit-btn")
signin_button.click()

print("Step 02... complete.")
time.sleep(3)

# STEP 03A
# Wait-for subsequent-session warning
CLASS_subsequent_session_button = "btn-primary"
try:
    el = WebDriverWait(driver, 15).until(
        EC.visibility_of_element_located((By.CLASS_NAME, CLASS_subsequent_session_button))
        )
    el.click()
except:
    print("Warning: Could not find Subsequent Session Prompt.")

# STEP 03B
# Prepare to switch windows (MLS appears to do a change-over w/o me knowing)
main_window = driver.window_handles[0]

# STEP 03C
# Wait-for/Click-on Agent/Office Rosters (i.e., the search form)
CSS_office_roster_link = "div.quicklink-item:nth-child(1) > a:nth-child(2)"
try:
    el = WebDriverWait(driver, 15).until(
        EC.visibility_of_element_located((By.CSS_SELECTOR, CSS_office_roster_link))
        )
    el.click()
except:
    print("Error: Could not find Agent/Office Rosters hyperlink.")
    end_session(driver)

# STEP 03D
# Switch to new window opened as master
try:
    assert len(driver.window_handles) == 1
    print("One window confirmed. Proceeding...")
except:
    slave_window = driver.window_handles[1]
    driver.switch_to.window(slave_window)

print("Step 03... complete.")
time.sleep(3)

# STEP 04
# Wait-for/fill-out search form (dud attempt)
NAME_state_field = "state"
try:
    el = WebDriverWait(driver, 15).until(
        EC.visibility_of_element_located((By.NAME, NAME_state_field))
        )
except:
    print("Error: Could not find State search-field.")
    end_session(driver)

state_dropdown = driver.find_element(By.NAME, NAME_state_field)
state_dropdown.send_keys("Massachusetts")
state_dropdown.send_keys(Keys.ENTER)

zip_field = driver.find_element(By.NAME, "zip")
zip_field.send_keys("00000")

print("Step 04... complete.")
time.sleep(3)

# STEP 05
# Wait-for/Click-on Find Offices button
ID_find_offices_button = "btnFindOffices"
try:
    el = WebDriverWait(driver, 15).until(
        EC.visibility_of_element_located((By.ID, ID_find_offices_button))
        )
    el.click()
except:
    print("Error: Could not find Find Offices button.")
    end_session(driver)

print("Step 05... complete.")
time.sleep(3)

# STEP 06
# Wait-for/Click on Edit Search
ID_edit_search_button = "btnEditSearch"
try:
    el = WebDriverWait(driver, 15).until(
        EC.visibility_of_element_located((By.ID, ID_edit_search_button))
        )
    el.click()
except:
    print("Error: Could not find Edit Search button.")
    end_session(driver)

print("Step 06... complete.")
time.sleep(3)

# STEP 07
# Initiate the DataFrame
realtor_db = []
realtor_db_headers = ["Office-Name", "Office-Info", "Type", "Agent-Name", "Agent-Phone", "Agent-Email"]

print("Step 07... complete.")
print("Stage 01... complete.")
time.sleep(3)

# STAGE THREE: Loop through zip-codes, scrape data and output to CSV
# Loop through by remaining steps by zip-code
print("Stage 02 initiating...")

for zip_code in zip_codes:
    # STEP 01
    # Re-Initialize an empty DataFrame every new zip code
    realtor_df = pd.DataFrame(realtor_db, columns = realtor_db_headers)

    # Clear the console
    os.system("cls")

    # Wait-for/fill-out search form
    NAME_state_field = "state"
    try:
        el = WebDriverWait(driver, 15).until(
            EC.visibility_of_element_located((By.NAME, NAME_state_field))
            )
    except:
        print("Error: Could not find State search-field.")
        end_session(driver)

    state_dropdown = driver.find_element(By.NAME, NAME_state_field)
    state_dropdown.send_keys("Massachusetts")
    state_dropdown.send_keys(Keys.ENTER)

    zip_field = driver.find_element(By.NAME, "zip")
    zip_field.clear()
    zip_field.send_keys(zip_code[0])

    print("Step 01... complete.")
    time.sleep(3)

    # STEP 02
    # Wait-for/Click-on Find Offices button
    ID_find_offices_button = "btnFindOffices"
    try:
        el = WebDriverWait(driver, 15).until(
            EC.visibility_of_element_located((By.ID, ID_find_offices_button))
            )
        el.click()
    except:
        print("Error: Could not find Find Offices button.")
        end_session(driver)

    print("Step 02... complete.")
    time.sleep(3)

    # STEP 03
    # Check if there are any results
    CLASS_no_results_label = "mls-ros-no-results"
    try:
        el = WebDriverWait(driver, 15).until(
            EC.visibility_of_element_located((By.CLASS_NAME, CLASS_no_results_label))
            )
    except:
        results_counter = driver.find_element(By.CLASS_NAME, "mls-sr-count-disp")
        print("There is/are {} office(s) for {}.".format(results_counter.get_attribute("innerText"), zip_code[0]))
    else:
        print("There are no valid search results for zip-code {}.".format(zip_code[0]))
        print("Proceeding to next zip code...")

        # Wait-for/Click on Edit Search
        ID_edit_search_button = "btnEditSearch"
        try:
            el = WebDriverWait(driver, 15).until(
                EC.visibility_of_element_located((By.ID, ID_edit_search_button))
                )
            el.click()
        except:
            print("Error: Could not find Edit Search button.")
            end_session(driver)
        continue

    print("Step 03... complete.")
    time.sleep(3)

    # STEP 04
    # Locate all results on page
    CLASS_office_name_links = "mls-js-ros-dtl-link"
    try:
        el = WebDriverWait(driver, 15).until(
            EC.visibility_of_element_located((By.CLASS_NAME, CLASS_office_name_links))
            )
    except:
        print("Error: Could not find Office links.")
        end_session(driver)

    office_names = driver.find_elements(By.CLASS_NAME, CLASS_office_name_links)

    print("Step 04... complete.")
    time.sleep(3)

    # STEP 05
    # Loop through page results
    for office in office_names:
        # Create empty list for data-set

        # Get office-name from search results (then click on it)
        print(office.get_attribute("innerText"))
        office.click()

        # Confirm office-info modal has successfully opened
        CLASS_modal_close_button = "mls-js-close-icon"
        try:
            el = WebDriverWait(driver, 15).until(
                EC.visibility_of_element_located((By.CLASS_NAME, CLASS_modal_close_button))
                )
        except:
            print("Error: Could not find Modal Close button.")
            end_session(driver)

        close_button = driver.find_element(By.CLASS_NAME, CLASS_modal_close_button)

        iframes = driver.find_elements(By.TAG_NAME, "iframe")

        if len(iframes) == 1:
            print("Error: Inadequat iFrames located.")
            close_button.click()
            end_session(driver)
        
        print("Step 05... complete.")
        time.sleep(3)

        # STEP 06A
        # Get pop-up box results (Selenium needs to 'switch' to the iframe)
        driver.switch_to.frame(1)

        print("Step 06A... complete.")

        # STEP 06B
        # Ensure elements are present
        # Done in order of how they appear
        # Locate the Office Information div
        CLASS_office_info_specs = "mls-ros-dtl-office-info-div"
        try:
            el = WebDriverWait(driver, 15).until(
                EC.visibility_of_element_located((By.CLASS_NAME, CLASS_office_info_specs))
                )
        except:
            print("Error: Could not find Office Information fields.")
            end_session(driver)
        office_info_specs = driver.find_elements(By.CLASS_NAME, CLASS_office_info_specs)

        
        # Locate the Office Contacts names
        CLASS_office_admin_names = "mls-ros-dtl-office-contact-names"
        try:
            el = WebDriverWait(driver, 15).until(
                EC.visibility_of_element_located((By.CLASS_NAME, CLASS_office_admin_names))
                )
        except:
            print("Warning: Could not find Office Admin names.")
            print("Proceeding to the next office...")
            driver.switch_to.default_content()
            close_button.click()
            # Skip to the next office (no Admin of prior office)
            continue

        office_admin_names = driver.find_elements(By.CLASS_NAME, CLASS_office_admin_names)

        # Locate the Agent Names
        CLASS_office_agent_names = "mls-ros-dtl-office-names"
        try:
            el = WebDriverWait(driver, 15).until(
                EC.visibility_of_element_located((By.CLASS_NAME, CLASS_office_agent_names))
                )
        except:
            print("Warning: Could not find Office Agent names.")
        # Selenium will return an empty list in the event that there are no agents associated with an office
        office_agent_names = driver.find_elements(By.CLASS_NAME, CLASS_office_agent_names)


        # Locate the Phone Numbers
        CLASS_agent_phone_numbers = "mls-phone-non-mobile"
        try:
            el = WebDriverWait(driver, 15).until(
                EC.visibility_of_element_located((By.CLASS_NAME, CLASS_agent_phone_numbers))
                )
        except:
            print("Error: Could not find Agent numbers.")
            end_session(driver)
        agent_phone_numbers = driver.find_elements(By.CLASS_NAME, CLASS_agent_phone_numbers)


        # Locate the Email addresses
        CLASS_agent_email_addresses = "mls-js-dtl-mail-to"
        try:
            el = WebDriverWait(driver, 15).until(
                EC.visibility_of_element_located((By.CLASS_NAME, CLASS_agent_email_addresses))
                )
        except:
            print("Error: Could not find Email addresses.")
            end_session(driver)
        agent_email_addresses = driver.find_elements(By.CLASS_NAME, CLASS_agent_email_addresses)

        print("Step 06B... complete.")
        
        # STEP 06C
        # Get office information
        office_information = ""
        office_info_specs_len = len(office_info_specs)
        j = 0
        for office_info in office_info_specs:
            if office_info_specs_len - 1 == j:
                office_information += office_info.get_attribute("innerText")
            else:
                office_info_text = office_info.get_attribute("innerText")
                office_info_text += "~~~"
                office_information += office_info_text
            j += 1

        office_information = re.sub(r"\s+", " ", office_information)

        print("Step 06C... complete.")

        # STEP 06D
        # Get office admin information
        number_of_admins = len(office_admin_names)
        admin_table = driver.find_element(By.CSS_SELECTOR, "#mlsOfficeContactsPanel > div > table")
        admin_name = ""
        admin_number = ""
        admin_email = ""
        if number_of_admins == 1:
            try:
                admin_name_obj = admin_table.find_element(By.CSS_SELECTOR, "tbody > tr.mls-ros-dtl-agent-alt-row.mls-js-ros-dtl-row > td.mls-ros-dtl-office-contacts-titlename-col.text-nowrap > div > div.col-12.col-sm-6 > a > span")
                admin_name = admin_name_obj.get_attribute("innerText")
            except NoSuchElementException:
                admin_name = "Not Found"
            
            try:
                admin_number_obj = admin_table.find_element(By.CSS_SELECTOR, "tbody > tr.mls-ros-dtl-agent-alt-row.mls-js-ros-dtl-row > td.mls-ros-dtl-office-phone-col.mls-hidden-xs-down > span.mls-phone-non-mobile.text-nowrap")
                admin_number = admin_number_obj.get_attribute("innerText")
            except NoSuchElementException:
                admin_number = "Not Found"
            
            try:
                admin_email_obj = admin_table.find_element(By.CSS_SELECTOR, "tbody > tr.mls-ros-dtl-agent-alt-row.mls-js-ros-dtl-row > td.float-right.text-nowrap > button.btn.btn-sm.mls-btn-transparent.mls-js-dtl-mail-to.btn.mls-dtl-mail-to.ml-1")
                admin_email = admin_email_obj.get_attribute("data-mailto")
            except NoSuchElementException:
                admin_email = "Not Found"

        else:
            try:
                admin_name_obj = admin_table.find_element(By.CSS_SELECTOR, "tbody > tr:nth-child(2) > td.mls-ros-dtl-office-contacts-titlename-col.text-nowrap > div > div.col-12.col-sm-6 > a > span")
                admin_name = admin_name_obj.get_attribute("innerText")
            except NoSuchElementException:
                admin_name = "Not Found"
            
            try:
                admin_number_obj = admin_table.find_element(By.CSS_SELECTOR, "tbody > tr:nth-child(2) > td.mls-ros-dtl-office-phone-col.mls-hidden-xs-down > span.mls-phone-non-mobile.text-nowrap")
                admin_number = admin_number_obj.get_attribute("innerText")
            except NoSuchElementException:
                admin_number = "Not Found"
            
            try:
                admin_email_obj = admin_table.find_element(By.CSS_SELECTOR, "tbody > tr:nth-child(2) > td.float-right.text-nowrap > button.btn.btn-sm.mls-btn-transparent.mls-js-dtl-mail-to.btn.mls-dtl-mail-to.ml-1")
                admin_email = admin_email_obj.get_attribute("data-mailto")
            except NoSuchElementException:
                admin_email = "Not Found"

        # Add admin information to a dataframe
        admin_data = pd.DataFrame(
            {
            "Office-Name": [office.get_attribute("innerText")],
            "Office-Info": [office_information],
            "Type": ["Admin"],
            "Agent-Name": [admin_name],
            "Agent-Phone": [admin_number],
            "Agent-Email": [admin_email]
            }
        )

        # Append admin dataframe to main dataframe
        realtor_df = pd.concat([realtor_df, admin_data], ignore_index=True)

        print("Step 06D... complete.")

        # STEP 06E
        # Get office agent information
        number_of_agents = len(office_agent_names)
        agent_tables = driver.find_elements(By.CSS_SELECTOR, "#mlsOfficeSubscriberPanel > div > table")
        if agent_tables and number_of_agents != 0:
            # If must verify that agent count is greater-than 0 AND that the subscriber-panel table exists
            agent_table = driver.find_element(By.CSS_SELECTOR, "#mlsOfficeSubscriberPanel > div > table")
            i = 1
            for agent in range(number_of_agents):
                agent_name = ""
                try:
                    agent_name_obj = agent_table.find_element(By.CSS_SELECTOR, "tbody > tr:nth-child(" + str(i) + ") > td.mls-ros-dtl-office-agentname-col > a > span")
                    agent_name = agent_name_obj.get_attribute("innerText")
                except NoSuchElementException:
                    agent_name = "Not Found"

                agent_number = ""
                try:
                    agent_number_obj = agent_table.find_element(By.CSS_SELECTOR, "tbody > tr:nth-child(" + str(i) + ") > td.mls-hidden-xs-down.mls-ros-dtl-office-phone-col > span.mls-phone-non-mobile.text-nowrap")
                    agent_number = agent_number_obj.get_attribute("innerText")
                except NoSuchElementException:
                    agent_number = "Not Found"

                agent_email = ""
                try:
                    agent_email_obj = agent_table.find_element(By.CSS_SELECTOR, "tbody > tr:nth-child(" + str(i) + ") > td.float-right.text-nowrap > button.btn.btn-sm.mls-btn-transparent.mls-js-dtl-mail-to.btn.mls-dtl-mail-to.ml-1")
                    agent_email = agent_email_obj.get_attribute("data-mailto")
                except NoSuchElementException:
                    agent_email = "Not Found"

                # Add agent information to a dataframe
                agent_data = pd.DataFrame(
                    {
                    "Office-Name": [office.get_attribute("innerText")],
                    "Office-Info": [office_information],
                    "Type": ["Agent"],
                    "Agent-Name": [agent_name],
                    "Agent-Phone": [agent_number],
                    "Agent-Email": [agent_email]
                    }
                )

                # Append agent dataframe to main dataframe
                realtor_df = pd.concat([realtor_df, agent_data], ignore_index=True)

                # Increase incrementor
                i+=1

        # Switch back to the main window
        driver.switch_to.default_content()
        close_button.click()

        print("Step 06E... complete.")
        time.sleep(3)

    # STEP 07
    # Dump data and proceed to next search

    # Dump data to the CSV
    realtor_df.to_csv('Raw-Realtor-Data.csv', mode = "a", index = False, header = False)

    # Wait-for/Click on Edit Search
    ID_edit_search_button = "btnEditSearch"
    try:
        el = WebDriverWait(driver, 15).until(
            EC.visibility_of_element_located((By.ID, ID_edit_search_button))
            )
        el.click()
    except:
        print("Error: Could not find Edit Search button.")
        end_session(driver)
   
    print("Step 07... complete.")
    time.sleep(5)

    # Next Step: Proceed to next search

    # TODO -> TEMPORARY ... DELETE ME
    # end_session(driver)

print("~~~~~~~~~~~~~~~~~~~~~~")
print("~~~~~~~~~~~~~~~~~~~~~~")
print("All zip codes have successfully been scraped! :)")
print("~~~~~~~~~~~~~~~~~~~~~~")
print("~~~~~~~~~~~~~~~~~~~~~~")
