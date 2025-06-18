import pyautogui
import time
import ctypes
import sys
import subprocess

print("You have 5 seconds to place your cursor in a text field...")
time.sleep(5)  # Gives you time to click a text field

pyautogui.press('capslock')

pyautogui.typewrite('This is a quick test message.', interval=0.1)