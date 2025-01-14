import pyautogui
import time
import pandas as pd
from selenium import webdriver

# Read data from Excel file and create a dictionary
file_path = r"US State Names.xlsx"
data = pd.read_excel(file_path)

pyautogui.FAILSAFE = False


# Function to connect to HMA VPN and establish a VPN connection
# def connect_vpn(state_name):
#     # Open HMA VPN using PyAutoGUI
#     print(state_name)
#     pyautogui.hotkey("win")
#     time.sleep(1)
#     pyautogui.write("Surfshark")
#     pyautogui.press("enter")
#     time.sleep(5)

#     # Select VPN server location
#     location = pyautogui.locateCenterOnScreen(
#         "surfshark_img/search_bar.PNG", confidence=0.8
#     )
#     if location is not None:
#         pyautogui.click(location)
#     else:
#         print("Image not found.")

#     time.sleep(2)
#     pyautogui.write(f"{state_name}")
#     time.sleep(10)

#     pyautogui.click(f"surfshark_us_states/{state_name}.PNG")
#     time.sleep(2)  # Click on the VPN server option
#     # Wait for the connection to establish (adjust the time as needed)
#     time.sleep(10)

#     # Disconnect VPN after connection
#     disconnect_vpn()


def connect_vpn(state_name):
    print(state_name)
    pyautogui.hotkey("win")
    time.sleep(1)
    pyautogui.write("Surfshark")
    pyautogui.press("enter")
    time.sleep(10)

    try:
        location = pyautogui.locateCenterOnScreen(
            "surfshark_img/search_bar.PNG", confidence=0.7
        )
        if location:
            pyautogui.click(location)
            time.sleep(2)
            pyautogui.write(f"{state_name}")
            time.sleep(10)

            # Adjust the filename dynamically
            state_image = f"surfshark_us_states/{state_name}.PNG"
            vpn_location = pyautogui.locateCenterOnScreen(state_image, confidence=0.7)
            if vpn_location:
                pyautogui.click(vpn_location)
                time.sleep(10)
            else:
                print(f"VPN location for {state_name} not found.")
        else:
            print("Search bar not found.")
    except pyautogui.ImageNotFoundException:
        print("Image not found.")
        pyautogui.screenshot(f"error_{state_name}.png")


# Function to disconnect from HMA VPN
def disconnect_vpn():
    # Open HMA VPN using PyAutoGUI
    pyautogui.hotkey("win")
    time.sleep(1)
    pyautogui.write("HMA VPN")
    time.sleep(1)
    pyautogui.press("enter")
    time.sleep(5)
    # Click disconnect button (you may need to adjust the coordinates based on HMA VPN's UI)
    disconnect_button_pos = pyautogui.locateCenterOnScreen(
        "img/disconnect_button.PNG", confidence=0.9
    )
    if disconnect_button_pos is not None:
        pyautogui.click(disconnect_button_pos)
        time.sleep(5)  # Wait for the disconnection

    # Close the VPN app
    # pyautogui.hotkey("alt", "f4")  # Close the VPN app
    time.sleep(2)


if __name__ == "__main__":
    # Loop through each state and connect to VPN
    for state_name in data["State"]:
        connect_vpn(state_name)
