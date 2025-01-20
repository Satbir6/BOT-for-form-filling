import pyautogui
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
import pandas as pd
import random
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException

# Disable PyAutoGUI fail-safe to prevent accidental interruptions
pyautogui.FAILSAFE = False

# File paths
leads_file = r"6336_satbir.xlsx"
states_file = r"US State Names.xlsx"
successful_leads_file = "Successful_Submissions.xlsx"

# Read data from Excel files
leads_data = pd.read_excel(leads_file)
leads_list = leads_data.to_dict(orient="records")


# Initialize global variables
successful_submissions = []


def connect_vpn(state_name):
    pyautogui.hotkey("win")
    time.sleep(1)
    pyautogui.write("Surfshark")
    pyautogui.press("enter")
    time.sleep(5)  # Wait for Surfshark to open

    # Ensure the state is different from the previous one

    location = pyautogui.locateCenterOnScreen(
        "surfshark_img/search_bar.PNG", confidence=0.6
    )
    if location:
        pyautogui.click(location)
        time.sleep(2)
        # Press Ctrl+A to select all, then Backspace to clear
        pyautogui.hotkey("ctrl", "a")
        time.sleep(0.5)  # Small delay to ensure the command is registered
        pyautogui.press("backspace")
        time.sleep(0.5)
        print(f"Connecting to VPN for state: {state_name}")
        pyautogui.write(state_name)
        time.sleep(5)
        state_image = f"surfshark_us_states/{state_name}.PNG"
        vpn_location = pyautogui.locateCenterOnScreen(state_image, confidence=0.6)
        pyautogui.click(vpn_location)
        time.sleep(20)

    print("VPN connection established.")


# Handle terms and conditions checkbox
def check_terms_checkbox(driver):
    checkbox_names = [
        "terms-n-condition",
        "terms-n-condition_1",
    ]  # List of possible checkbox names
    for checkbox_name in checkbox_names:
        try:
            terms_checkbox = driver.find_element(By.NAME, checkbox_name)
            driver.execute_script("arguments[0].scrollIntoView(true);", terms_checkbox)
            time.sleep(1)  # Allow time for scrolling
            terms_checkbox.click()
            print(f"Clicked on checkbox: {checkbox_name}")
            return  # Exit after successfully clicking
        except NoSuchElementException:
            pass
    print("No terms and conditions checkbox was found.")


def fill_form(driver, url, lead_data):
    driver.get(url)
    time.sleep(15)

    # Function to safely fill a form field
    def safe_fill_field(by_method, identifier, value):
        try:
            element = driver.find_element(by_method, identifier)
            element.send_keys(value)
        except Exception as e:
            print(f"Could not fill field {identifier}: {e}")

    # Fill in form fields using the safe_fill_field function
    safe_fill_field(By.NAME, lead_data.get("first_name-code", ""), lead_data.get("first_name", ""))
    safe_fill_field(By.NAME, lead_data.get("last_name-code", ""), lead_data.get("last_name", ""))
    safe_fill_field(By.NAME, lead_data.get("email-code", ""), lead_data.get("email", ""))
    safe_fill_field(By.NAME, lead_data.get("phone-code", ""), lead_data.get("phone", ""))
    safe_fill_field(By.NAME, lead_data.get("company_name-code", ""), lead_data.get("company_name", ""))
    safe_fill_field(By.NAME, lead_data.get("job_title-code", ""), lead_data.get("job_title", ""))
    safe_fill_field(By.NAME, lead_data.get("industry-code", ""), lead_data.get("industry", ""))
    safe_fill_field(By.NAME, lead_data.get("employee_size-code", ""), lead_data.get("employee_size", ""))
    safe_fill_field(By.NAME, lead_data.get("street_1-code", ""), lead_data.get("street_1", ""))
    safe_fill_field(By.NAME, lead_data.get("city-code", ""), lead_data.get("city", ""))
    safe_fill_field(By.NAME, lead_data.get("state-code", ""), lead_data.get("state", ""))
    safe_fill_field(By.NAME, lead_data.get("code-code", ""), lead_data.get("code", ""))
    safe_fill_field(By.NAME, lead_data.get("country-code", ""), lead_data.get("country", ""))
    safe_fill_field(By.NAME, lead_data.get("q1-code", ""), lead_data.get("q1", ""))
    safe_fill_field(By.NAME, lead_data.get("q2-code", ""), lead_data.get("q2", ""))
    safe_fill_field(By.NAME, lead_data.get("q3-code", ""), lead_data.get("q3", ""))

    # Check terms and conditions checkbox
    check_terms_checkbox(driver)

    time.sleep(5)  # Wait for form submission (adjust as needed)HMA VPN

    # Scroll to the submit button before clicking it
    submit_button = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable(
            (
                By.CSS_SELECTOR,
                "button[type='submit'].ff-btn.ff-btn-submit.ff-btn-md.ff_btn_style",
            )
        )
    )
    driver.execute_script("arguments[0].scrollIntoView(true);", submit_button)
    submit_button.click()

    time.sleep(10)
    print("Form filled and submitted")
    successful_submissions.append(lead_data)


def disconnect_vpn():
    print("Disconnecting VPN...")
    pyautogui.hotkey("win")
    time.sleep(1)
    pyautogui.write("Surfshark")
    pyautogui.press("enter")
    time.sleep(5)
    disconnect_button_pos = pyautogui.locateCenterOnScreen(
        "surfshark_img/disconnect_button.PNG", confidence=0.6
    )

    pyautogui.click(disconnect_button_pos)
    time.sleep(3)


if __name__ == "__main__":
    for lead in leads_list:
        state_name = lead.get("state_lead")
        connect_vpn(state_name)
        # Initialize webdriver
        options = webdriver.ChromeOptions()
        options.add_experimental_option("detach", True)
        driver = webdriver.Chrome(options=options)

        url = lead["url"]
        fill_form(driver, url, lead)

        # Save successful submissions to Excel file
        pd.DataFrame(successful_submissions).to_excel(
            successful_leads_file, index=False
        )

        driver.quit()
        disconnect_vpn()
