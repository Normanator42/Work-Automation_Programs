# This program utilises web automation to transfer all inspection data for all completed jobs onto a Sydney Water Site to be reviewed

from datetime import datetime
import os
import shutil
import pandas as pd
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import ElementClickInterceptedException, TimeoutException, NoSuchElementException, StaleElementReferenceException, ElementNotInteractableException
from webdriver_manager.chrome import ChromeDriverManager

# Path to the reviewed and edited Excel file
def copy_with_date_time(src_path, dest_dir, base_filename):
    # Split the filename into name and extension
    name, ext = os.path.splitext(base_filename)
    
    # Get the current date and time
    current_date_time = datetime.now().strftime('%d%m%Y_%H%M')
    
    # Construct the new filename with date and time
    new_filename = f"{name}_{current_date_time}{ext}"
    dest_path = os.path.join(dest_dir, new_filename)
    
    shutil.copy2(src_path, dest_path)

# Define paths
dest_dir = "F:\\NR CCI UPLOADS\\PENDING\\EXCEL COPIES"
base_filename = "COPY.xlsx"
output_file_path = r"F:\NR CCI UPLOADS\PENDING\compiled_data.xlsx"

# Copy the file with date and time appended to the filename
copy_with_date_time(output_file_path, dest_dir, base_filename)


# Load the reviewed data
data = pd.read_excel(output_file_path)

# Set up Selenium WebDriver
options = Options()
options.add_argument("--start-maximized")
options.add_argument("--disable-web-security")
options.add_argument("--disable-gpu")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")
options.add_argument("--window-size=1920x1080")  # Set window size to a common resolution
options.add_argument("--headless")
options.add_argument("--disable-extensions")
options.add_argument("--log-level=3")  # Suppress warnings and informational messages
#options.add_argument("--incognito")

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)


def wait_for_element_and_click(xpath, max_attempts=3, wait_time=30):
    attempts = 0
    while attempts < max_attempts:
        try:
            WebDriverWait(driver, wait_time).until(
                EC.presence_of_element_located((By.XPATH, xpath))
            )
            WebDriverWait(driver, wait_time).until(
                EC.element_to_be_clickable((By.XPATH, xpath))
            )
            element = driver.find_element(By.XPATH, xpath)
            driver.execute_script("arguments[0].click();", element)
            return True
        except TimeoutException:
            attempts += 1
        except ElementClickInterceptedException:
            attempts += 1
            print(f"ElementClickInterceptedException: Attempt {attempts} for {xpath}")
        except Exception as e:
            attempts += 1
            print(f"Exception: {e}. Attempt {attempts} for {xpath}")
    return False


# Define the login function
def login():
    driver.get("https://media.sydneywater.com.au/")
    try:
        print("Attempting to find 'Sydney Water Staff' button...")
        driver.implicitly_wait(10)
        sydney_water_staff_button = WebDriverWait(driver, 300).until(
            EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Sydney Water Staff')]"))
        )
        driver.execute_script("arguments[0].scrollIntoView(true);", sydney_water_staff_button)
        driver.execute_script("arguments[0].click();", sydney_water_staff_button)
        print("'Sydney Water Staff' button clicked.")
        
        print("Waiting for the organizational login page to load...")
        driver.implicitly_wait(10)
        username_field = WebDriverWait(driver, 300).until(
            EC.presence_of_element_located((By.ID, "userNameInput"))
        )
        print("Username field found.")
        username_field.send_keys("***@sydneywater.com.au")
        print("Username entered.")
        
        print("Waiting for the password field...")
        driver.implicitly_wait(10)
        password_field = WebDriverWait(driver, 300).until(
            EC.presence_of_element_located((By.ID, "passwordInput"))
        )
        print("Password field found.")
        password_field.send_keys("******")
        print("Password entered.")
        
        driver.implicitly_wait(10)
        sign_in_button = WebDriverWait(driver, 300).until(
            EC.element_to_be_clickable((By.ID, "submitButton"))
        )
        driver.execute_script("arguments[0].click();", sign_in_button)
        print("Sign in button clicked.")
        
    except Exception as e:
        print(f"Login failed: {e}")
        driver.save_screenshot("login_failure.png")
        print("Screenshot of the error saved as 'login_failure.png'.")
        driver.quit()
        raise e
    

# Define the function to fill out the form
def fill_out_form(row):
    while True:
        try:
            print("Waiting for the 'Add New' button to become clickable...")
            driver.implicitly_wait(10)
            add_new_button = WebDriverWait(driver, 300).until(
                EC.element_to_be_clickable((By.XPATH, "//button[@id='a11y-addNewDropDown']"))
            )
            driver.execute_script("arguments[0].scrollIntoView(true);", add_new_button)
            driver.execute_script("arguments[0].click();", add_new_button)
            print("'Add New' button clicked.")
            
            driver.implicitly_wait(10)
            media_upload_option = WebDriverWait(driver, 300).until(
                EC.element_to_be_clickable((By.XPATH, "//a[@href='/upload/media']"))
            )
            driver.execute_script("arguments[0].click();", media_upload_option)
            print("'Media Upload' option clicked.")
            
            driver.implicitly_wait(10)
            file_input = WebDriverWait(driver, 300).until(
                EC.presence_of_element_located((By.XPATH, "//input[@type='file']"))
            )
            file_input.send_keys(str(row['Inspection Video(s)']))
            print(f"Video file uploaded.")

            driver.implicitly_wait(10)
            name_field = WebDriverWait(driver, 300).until(
                EC.presence_of_element_located((By.ID, "Entry-name"))
            )
            asset_numbers = str(row['Pipe Asset ID']).replace(', ', '_')

            title1 = '' if pd.isna(row['Priority Justification']) else str(row['Priority Justification'])

            name_field.send_keys(f"{title1}_{asset_numbers}_{int(row['JSA/WO'])}_{row['Location Scamp']}_{int(row['Attempt #'])}")
            
            # Switch to the iframe for the WYSIWYG editor
            driver.implicitly_wait(10)
            iframe = WebDriverWait(driver, 300).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "iframe.wysihtml5-sandbox"))
            )
            driver.switch_to.frame(iframe)
            
            # Locate the body of the WYSIWYG editor and send keys
            driver.implicitly_wait(10)
            description_body = WebDriverWait(driver, 300).until(
                EC.presence_of_element_located((By.TAG_NAME, "body"))
            )
            description_body.clear()  # Clear any existing content
            description_body.send_keys(str(row['General comment']))

            # Switch back to the main document
            driver.switch_to.default_content()

            # Check the cleaned field for "JJ" or "jj" to set the tags field
            tags_field = driver.find_element(By.ID, "s2id_autogen2")
            tags = "junctionjet" if "JJ" in str(row['Cleaning']) or "jj" in str(row['Cleaning']) else "cctv"
            tags_field.send_keys(tags)
            time.sleep(2)  # Allow time for the dropdown to appear
            tags_field.send_keys(Keys.ENTER)

            parent_wo_field = driver.find_element(By.ID, "customdata-ParentWorkOrderNumber0")
            parent_wo_field.send_keys(str(row['JSA/WO']) if isinstance(row['JSA/WO'], str) else int(row['JSA/WO']))
            
            child_wo_numbers = str(row['Child WO']).split(',')
            for i, child_wo_number in enumerate(child_wo_numbers):
                child_wo_fields = driver.find_elements(By.XPATH, "//input[@name='customdata[ChildWorkOrderNumbers][]']")
                add_button = driver.find_element(By.ID, "customdata-ChildWorkOrderNumbers-addBtn")
                
                if i < len(child_wo_fields):
                    child_wo_fields[i].send_keys(child_wo_number.strip())
                else:
                    driver.execute_script("arguments[0].click();", add_button)
                    child_wo_fields = driver.find_elements(By.XPATH, "//input[@name='customdata[ChildWorkOrderNumbers][]']")
                    child_wo_fields[i].send_keys(child_wo_number.strip())
            
            wo_description_field = driver.find_element(By.ID, "customdata-WorkOrderDescription0")
            wo_description_field.send_keys(str(row['WO description']))
            
            asset_numbers = str(row['Pipe Asset ID']).split(',')
            for i, asset_number in enumerate(asset_numbers):
                asset_fields = driver.find_elements(By.XPATH, "//input[@name='customdata[AssetNumbers][]']")
                add_button = driver.find_element(By.ID, "customdata-AssetNumbers-addBtn")
                
                if i < len(asset_fields):
                    asset_fields[i].send_keys(asset_number.strip())
                else:
                    driver.execute_script("arguments[0].click();", add_button)
                    asset_fields = driver.find_elements(By.XPATH, "//input[@name='customdata[AssetNumbers][]']")
                    asset_fields[i].send_keys(asset_number.strip())

            task_code_field = driver.find_element(By.ID, "customdata-TaskCode")
            task_code_field.send_keys(str(row['Task code']))
            
            suburb_field = driver.find_element(By.ID, "customdata-Suburb0")
            suburb_field.send_keys(str(row['Suburb']))
            
            address_st_field = driver.find_element(By.ID, "customdata-AddressStreet0")
            address_st_field.send_keys(str(row['Address/Location']))
            
            product_field = driver.find_element(By.ID, "customdata-Product0")
            product_field.send_keys("Wastewater")
            
            contractor_field = driver.find_element(By.ID, "customdata-Contractor")
            contractor_field.send_keys("COMDAININF-001")
            
            usmh_field = driver.find_element(By.ID, "customdata-UpstreamMH")
            usmh_field.send_keys(str(row['US MH']) if isinstance(row['US MH'], str) else int(row['US MH']))
            
            dsmh_field = driver.find_element(By.ID, "customdata-DownstreamMH")
            dsmh_field.send_keys(str(row['DS MH']) if isinstance(row['DS MH'], str) else int(row['DS MH']))
            
            direction_of_survey_field = driver.find_element(By.ID, "customdata-DirectionOfSurvey")
            direction_of_survey_field.send_keys(str(row['Inspection Direction']))
            
            date_field = driver.find_element(By.ID, "customdata-DateOfCompletedInspection")
            date_field.send_keys(str(row['Date of inspection']))
            
            time_field = driver.find_element(By.ID, "customdata-TimeOfCompletedInspection")
            time_field.send_keys(str(row['Time of inspection']))
            
            package_name_field = driver.find_element(By.ID, "customdata-PackageName")
            package_name_field.send_keys(str(row['PackageName']))
            
            cleaned_field = driver.find_element(By.ID, "customdata-Cleaned0")
            cleaned_field.send_keys(str(row['Cleaning']))
            
            surveyed_length_field = driver.find_element(By.ID, "customdata-SurveyedLength0")
            surveyed_length_field.send_keys(str(row['Inspected Length [m]']) + "m")
            
            location_scamp_field = driver.find_element(By.ID, "customdata-Location0")
            location_scamp_field.send_keys(str(row['Operational Area'])) 

            priority_justification_values = ["CRIT", "CRCM", "MHIP", "WATW", "WACC", "WAEP", "WAER", "INTS", "ODOR", "ODCC", "MULT", "MULC", "OPER", "OPCC", "OPSC", "OPOH", "REPT", "RECC", "SALT", "SEEP", "SECC", "SUBX", "SUCC", "SUSC", "SUBS", "WETW", "WECC"]
            priority_justification = str(row['Priority Justification']) if any(val in str(row['Priority Justification']) for val in priority_justification_values) else "OTHER"

            priority_justification_field = driver.find_element(By.ID, "customdata-PriorityJustification")
            priority_justification_field.send_keys(priority_justification)
        
            driver.implicitly_wait(10)
            WebDriverWait(driver,  300).until(
                lambda driver: "Upload Completed!" in driver.find_element(By.XPATH, "//div[contains(@class, 'alert-success')]//strong").text
            )
            print("Video upload complete.")
            time.sleep(1) 

            save_button = driver.find_element(By.ID, "Entry-submit")
            driver.execute_script("arguments[0].click();", save_button)
            print("Saved after video uploading")

            driver.implicitly_wait(10)
            go_to_media_button = WebDriverWait(driver, 300).until(
                EC.element_to_be_clickable((By.XPATH, "//a[@id='back']"))
            )
            driver.execute_script("arguments[0].click();", go_to_media_button)
            print("Clicked 'Go to Media'.")
            
            driver.implicitly_wait(10)
            actions_button = WebDriverWait(driver, 300).until(
                EC.element_to_be_clickable((By.XPATH, "//button[@id='entryActionsMenuBtn']"))
            )
            driver.execute_script("arguments[0].click();", actions_button)
            print("Clicked 'Actions' button.")

            driver.implicitly_wait(10)
            edit_option = WebDriverWait(driver, 300).until(
                EC.element_to_be_clickable((By.XPATH, "//a[@id='tab-Edit']"))
            )
            driver.execute_script("arguments[0].click();", edit_option)
            print("Selected 'Edit' option.")

            driver.implicitly_wait(10)
            attachments_tab = WebDriverWait(driver, 300).until(
                EC.element_to_be_clickable((By.XPATH, "//a[@id='attachments-tab-tab']"))
            )
            driver.execute_script("arguments[0].click();", attachments_tab)
            print("Selected 'attachments' tab.")
            
            driver.implicitly_wait(10)
            upload_file_btn = WebDriverWait(driver, 300).until(
                EC.element_to_be_clickable((By.XPATH, "//a[contains(@href, '/attachments/index/add/entryid') and contains(@class, 'btn btn-primary')]"))
            )
            driver.execute_script("arguments[0].click();", upload_file_btn)
            print("Selected 'Upload file' button.")
            
            driver.implicitly_wait(10)
            file_input = WebDriverWait(driver, 300).until(
                EC.presence_of_element_located((By.XPATH, "//input[@id='attachments_fileinput']"))
            )
            file_input.send_keys(str(row['Section PDF Filename']))

            print(f"PDF file uploaded")
            
            driver.implicitly_wait(10)
            save_file_btn = WebDriverWait(driver, 300).until(
                EC.element_to_be_clickable((By.XPATH, "//a[contains(@class, 'btn attachment-save-btn btn-primary')]"))
            )

            # Extract and adjust the date of inspection
            try:
                inspection_time = pd.to_datetime(row['Time of inspection'], format='%H:%M:%S').time()
            except ValueError:
                inspection_time = pd.to_datetime(row['Time of inspection'], format='%H:%M').time()

            inspection_date = pd.to_datetime(row['Date of inspection'], dayfirst=True)

            # Check if the inspection was done between midnight and 5:30 AM
            if inspection_time >= datetime.strptime('00:00', '%H:%M').time() and inspection_time <= datetime.strptime('05:30', '%H:%M').time():
                inspection_date -= pd.Timedelta(days=1)

            # Format the adjusted date
            date_of_inspection = inspection_date.strftime('%d%m%Y')
            pdf_filename = f"{row['JSA/WO']} CCTV REPORT {date_of_inspection}.pdf"
            shutil.copy(str(row['Section PDF Filename']), f"F:\\NR CCI UPLOADS\\PDF REPORTS\\{pdf_filename}")
            time.sleep(1)
            driver.execute_script("arguments[0].click();", save_file_btn)
            print(f"Saved Uploaded PDF as '{pdf_filename}' and saved to 'PDF REPORTS' folder.")

            driver.implicitly_wait(10)
            publish_tab = WebDriverWait(driver, 300).until(
                EC.element_to_be_clickable((By.XPATH, "//a[@id='Publish-tab']"))
            )
            driver.execute_script("arguments[0].click();", publish_tab)
            print("Selected 'publish' tab.")
            
            # Select 'Published' option
            published_xpath = "//input[@id='published_entry'][@value='published']"
            if not wait_for_element_and_click(published_xpath):
                print("\nRETRYING UPLOAD...\n")
                continue  # Retry the form filling process

            # Select 'COMDAININF-001' box
            comdaininf_checkbox = driver.find_element(By.XPATH, "//input[@id='CategoryTree-214'][@value='1']")
            comdaininf_xpath = "//input[@id='CategoryTree-214'][@value='1']"
            if not wait_for_element_and_click(comdaininf_xpath):
                print("\nRETRYING UPLOAD...\n")
                continue  # Retry the form filling process

            # Click final save
            final_save_button = "//button[contains(@class, 'btn btn-primary pblSave')]"
            if not wait_for_element_and_click(final_save_button):
                print("\nRETRYING UPLOAD...\n")
                continue  # Retry the form filling process

            # Remove the processed row from the DataFrame and save the updated DataFrame
            print(f"Processed and removed row from the Excel file.")
            break
            
        except ElementClickInterceptedException as e:
            print(f"ElementClickInterceptedException: {e}")
            driver.execute_script("arguments[0].click();", save_file_btn)
            print("Clicked 'Save' button using JavaScript.")
        except StaleElementReferenceException as e:
            print(f"StaleElementReferenceException: {e}")
            save_file_btn = WebDriverWait(driver, 300).until(
                EC.element_to_be_clickable((By.XPATH, "//a[contains(@class, 'btn attachment-save-btn btn-primary')]"))
            )
            driver.execute_script("arguments[0].click();", save_file_btn)
            print("Retried clicking 'Save' button.")
        except ElementNotInteractableException as e:
            print(f"ElementNotInteractableException: {e}")
            driver.execute_script("arguments[0].click();", go_to_media_button)
            print("Clicked 'Go to Media' using JavaScript.")
        except Exception as e:
            print(f"Form filling failed: {e}")
            driver.save_screenshot("form_filling_failure.png")
            print("Screenshot of the error saved as 'form_filling_failure.png'.")
            driver.quit()
            raise e


# Main script execution
try:
    login()
    folders_to_move = set()
    for index, row in data.iterrows():
        if pd.isna(row['Inspected Length [m]']):
            continue  # Skip rows with NaN in 'Inspected Length [m]'
        fill_out_form(row)
        # Collect the folders to move
        video_path = str(row['Inspection Video(s)'])
        parent_folder = os.path.dirname(os.path.dirname(os.path.dirname(video_path)))
        folders_to_move.add(parent_folder)
        # Remove the processed row from the DataFrame and save the updated DataFrame
        data.drop(index, inplace=True)
        data.to_excel(output_file_path, index=False)
    print("COMPLETED")
finally:
    driver.quit()
    # Create a new folder with the current date and time within "UPLOADED" folder
    current_date_time_folder = datetime.now().strftime('%d%m%Y')
    destination_root_folder = os.path.join("F:\\NR CCI UPLOADS\\UPLOADED", current_date_time_folder)
    os.makedirs(destination_root_folder, exist_ok=True)
    # Move the collected folders to the new generated folder
    for folder in folders_to_move:
        destination_folder = os.path.join(destination_root_folder, os.path.basename(folder))
        if not os.path.exists(destination_folder):
            shutil.move(folder, destination_folder)
        else:
            print(f"Folder {destination_folder} already exists. Skipping move operation.")
