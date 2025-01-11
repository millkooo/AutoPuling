import win32gui, win32com.client, pythoncom
import sys
import keyboard
import time, random
from pynput.keyboard import Controller
import win32con
import win32api
import json
    
def send_key_to_window(key='f', tm=0.2, keyup=True):
    hwnd = game_nd
    key = 0x41 + ord(key) - ord('a')
    win32api.PostMessage(hwnd, win32con.WM_KEYDOWN, key, 0)
    time.sleep(random.random()*0.2+tm)
    if keyup:
        win32api.PostMessage(hwnd, win32con.WM_KEYUP, key, 0)    

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
        class_name = win32gui.GetClassName(hwnd).strip()
        window_name = win32gui.GetWindowText(hwnd).strip()
        if (window_name == 'InfinityNikki' or window_name == '无限暖暖') and class_name == 'UnrealWindow':
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
    if foreground:
        set_forground(game_nd)
    return game_nd

def press_key_pynput(key, tm, keyup):
    global stop, game_nd
    cnt = 0
    while win32gui.GetForegroundWindow() != game_nd:
        time.sleep(1)
        if stop:
            return
        cnt += 1
        if cnt > 600:
            print("尝试将游戏窗口置于前台...")
            game_nd = init_window()
            if game_nd is None:
                print("未找到游戏窗口，停止运行")
                stop = 1
                return
            time.sleep(1)
            break
    keyboard = Controller()
    keyboard.press(key)
    time.sleep(random.random()*0.2+tm)
    if keyup:
        keyboard.release(key)

stop = False
def on_key_press(event):
    if event.name == "f8":
        global stop
        print("F8 已被按下，尝试停止运行")
        stop = True

def press(key, tm, keyup = True):
    if foreground:
        press_key_pynput(key, tm, keyup)
    else:
        send_key_to_window(key, tm, keyup)

def load_info():
    try:
        with open('info.txt', 'r') as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        return {}

def save_info(info):
    with open('info.txt', 'w') as f:
        json.dump(info, f)

def ask_user_choice():
    print("请选择运行模式：0为前台模式，1为后台模式(测试功能，请保证游戏窗口状态为被覆盖而不是最小化)")
    while True:
        try:
            foreground = int(input("输入您的选择："))
            if foreground in [0, 1]:
                break
        except ValueError:
            pass
        print("无效输入，请输入0或1。")

    print("是否记住您的选择？0为否，1为记住选择，2为否且不再询问")
    while True:
        try:
            remember = int(input("输入您的选择："))
            if remember in [0, 1, 2]:
                break
        except ValueError:
            pass
        print("无效输入，请输入0、1或2。")

    return foreground, remember

info = load_info()
foreground = info.get("foreground")
remember_choice = info.get("remember_choice", 0)

if remember_choice == 0 or remember_choice == 2:
    foreground, remember_choice = ask_user_choice()

    # 更新并保存选择
    info["foreground"] = foreground
    info["remember_choice"] = remember_choice
    save_info(info)

if remember_choice == 2:
    print("已选择不再询问，请注意需手动删除info.txt以重新设置。")
foreground = False
keyboard.on_press(on_key_press)
game_nd = init_window()
if game_nd is not None:
    tm = time.time()
    end_time = tm + 60 * 60 * 4
    print("游戏窗口已找到，开始运行")
    print(f"预计结束时间：{time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(end_time))}")
    time.sleep(1)
    press('f', 0.2)
    recent_a = 0

    notif_tm = time.time()
    while not stop and time.time() < end_time:
        if int((time.time() - tm) % (30 * 60)) <= 5 and time.time() - notif_tm > 60:
            print(f"已运行{int((time.time() - tm) // 60)}分钟，预计结束时间：{time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(end_time))}")
            notif_tm = time.time()
        time.sleep(random.random() * 1 + 2)
        press('f', 0.2)
        if time.time() - recent_a > 10:
            recent_a = time.time()
            press('a', 0.2, False)

    if stop:
        print("运行已停止")
    else:
        print("运行时间已达4.5小时，自动停止")

    time.sleep(1)
