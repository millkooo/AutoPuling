import win32gui, win32com.client, pythoncom
import sys
import keyboard
import time, random
from pynput.keyboard import Controller

def press_key_pynput(key):
    while win32gui.GetForegroundWindow() != game_nd:
        time.sleep(1)
    keyboard = Controller()
    keyboard.press(key)
    time.sleep(random.random()*0.2+0.2)
    keyboard.release(key)
    
def set_forground(game_nd):
    try:
        pythoncom.CoInitialize()
        shell = win32com.client.Dispatch("WScript.Shell")
        if getattr(sys, 'frozen', False):
            shell.SendKeys(" ")
        else:
            shell.SendKeys("")
        win32gui.SetForegroundWindow(game_nd)
    except:
        pass

def enum_windows_callback(hwnd, hwnds):
    try:
        class_name = win32gui.GetClassName(hwnd)
        window_name = win32gui.GetWindowText(hwnd)
        if '无限暖暖' in window_name and class_name == 'UnrealWindow':
            hwnds.append(hwnd)
    except:
        pass
    return True

def init_window():
    hwnds = []
    win32gui.EnumWindows(lambda a,b:enum_windows_callback(a,b), hwnds)
    if len(hwnds) == 0:
        print('未找到游戏窗口')
        time.sleep(5)
        return None
    game_nd = hwnds[0]
    set_forground(game_nd)
    return game_nd

stop = False
def on_key_press(event):
    if event.name == "f8":
        global stop
        print("F8 已被按下，尝试停止运行")
        stop = True

keyboard.on_press(on_key_press)
game_nd = init_window()
if game_nd is not None:
    time.sleep(1)
    press_key_pynput('f')
    while not stop:
        time.sleep(random.random()*4+2)
        press_key_pynput('f')
    time.sleep(3)
