import pyperclip
import pyautogui

def wrt(msg):
    pyperclip.copy(msg)
    pyautogui.hotkey('ctrl', 'v')