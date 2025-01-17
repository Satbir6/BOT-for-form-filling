import pyautogui
import time
import pandas as pd
from selenium import webdriver

# Read data from Excel file and create a dictionary
file_path = r"US State Names.xlsx"
data = pd.read_excel(file_path)

pyautogui.FAILSAFE = False


def connect_vpn(state_name):
    print(f"Connecting to VPN for {state_name}")
    pyautogui.hotkey("win")
    time.sleep(1)
    pyautogui.write("Surfshark")
    pyautogui.press("enter")
    time.sleep(15)  # Wait for Surfshark to open

    location = pyautogui.locateCenterOnScreen(
        "surfshark_img/search_bar_blank.PNG", confidence=0.6
    )
    if location:
        pyautogui.click(location)
        time.sleep(2)
        # Press Ctrl+A to select all, then Backspace to clear
        pyautogui.hotkey("ctrl", "a")
        time.sleep(0.5)  # Small delay to ensure the command is registered
        pyautogui.press("backspace")
        time.sleep(0.5)
        pyautogui.write(state_name)
        time.sleep(10)
        state_image = f"surfshark_us_states/{state_name}.PNG"
        vpn_location = pyautogui.locateCenterOnScreen(state_image, confidence=0.6)
        pyautogui.click(vpn_location)
        time.sleep(10)

    disconnect_vpn()


def disconnect_vpn():
    print("Disconnecting VPN...")
    pyautogui.hotkey("win")
    time.sleep(1)
    pyautogui.write("Surfshark")
    pyautogui.press("enter")
    time.sleep(10)
    disconnect_button_pos = pyautogui.locateCenterOnScreen(
        "surfshark_img/disconnect_button.PNG", confidence=0.6
    )

    pyautogui.click(disconnect_button_pos)
    time.sleep(5)


if __name__ == "__main__":
    # Loop through each state and connect to VPN
    for state_name in data["State"]:
        connect_vpn(state_name)
