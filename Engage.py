import pyautogui
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
import pandas as pd
import random
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Disable PyAutoGUI fail-safe to prevent accidental interruptions
pyautogui.FAILSAFE = False

# File paths
leads_file = r"test.xlsx"
states_file = r"US State Names.xlsx"
successful_leads_file = "Successful_Submissions.xlsx"

# Read data from Excel files
leads_data = pd.read_excel(leads_file)
leads_list = leads_data.to_dict(orient="records")

states_data = pd.read_excel(states_file)
state_names = states_data["State"].tolist()

# Initialize global variables
previous_state = None
successful_submissions = []


def connect_vpn():
    global previous_state

    # Open HMA VPN using PyAutoGUI
    pyautogui.hotkey("win")
    time.sleep(1)
    pyautogui.write("HMA VPN")
    pyautogui.press("enter")
    time.sleep(5)

    # Randomly select a state name different from the previous one
    name = random.choice(state_names)
    while name == previous_state:
        name = random.choice(state_names)

    previous_state = name

    print(f"Connecting to VPN server in {name}...")
    pyautogui.click("img/arrow.PNG")
    time.sleep(2)
    pyautogui.write(f"{name}")
    time.sleep(2)
    pyautogui.click(f"US_states_img/{name}.PNG")
    time.sleep(60)

    print("VPN connection established.")


def fill_form(driver, url, lead_data):
    driver.get(url)
    time.sleep(15)

    # Fill in form fields
    driver.find_element(By.NAME, "names[first_name]").send_keys(lead_data["first_name"])
    driver.find_element(By.NAME, "names[last_name]").send_keys(lead_data["last_name"])
    driver.find_element(By.NAME, "email").send_keys(lead_data["email"])
    driver.find_element(By.NAME, "input_text").send_keys(lead_data["phone"])
    driver.find_element(By.NAME, "input_text_1").send_keys(lead_data["company_name"])
    driver.find_element(By.NAME, "input_text_2").send_keys(lead_data["job_title"])
    driver.find_element(By.NAME, "dropdown").send_keys(lead_data["industry"])
    driver.find_element(By.NAME, "dropdown_1").send_keys(lead_data["employee_size"])
    driver.find_element(By.NAME, "input_text_3").send_keys(lead_data["street_1"])
    driver.find_element(By.NAME, "input_text_4").send_keys(lead_data["city"])
    driver.find_element(By.NAME, "input_text_8").send_keys(lead_data["code"])
    # driver.find_element(By.NAME, "input_text_7").send_keys(lead_data["state"])
    driver.find_element(By.NAME, "country-list").send_keys(lead_data["country"])

    # Scroll to the element before clicking
    terms_checkbox = driver.find_element(By.NAME, "terms-n-condition_3")
    driver.execute_script("arguments[0].scrollIntoView(true);", terms_checkbox)
    time.sleep(1)  # Allow time for scrolling
    terms_checkbox.click()

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
    # driver.execute_script("arguments[0].scrollIntoView(true);", submit_button)
    # submit_button.click()

    time.sleep(10)
    print("Form filled and submitted")
    successful_submissions.append(lead_data)


def disconnect_vpn():
    pyautogui.hotkey("win")
    time.sleep(1)
    pyautogui.write("HMA VPN")
    time.sleep(1)
    pyautogui.press("enter")
    time.sleep(5)

    # Click disconnect button (adjust coordinates based on UI)
    disconnect_button_pos = pyautogui.locateCenterOnScreen(
        "img/disconnect_button.PNG", confidence=0.9
    )
    if disconnect_button_pos is not None:
        pyautogui.click(disconnect_button_pos)
        time.sleep(5)

    print("VPN disconnected")


if __name__ == "__main__":
    for lead in leads_list:
        # connect_vpn()

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
        # disconnect_vpn()
