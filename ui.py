import sys
import os
import json
import time
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QLabel, QPushButton, QRadioButton, 
                             QButtonGroup, QSpinBox, QCheckBox, QProgressBar,
                             QTextEdit, QGroupBox, QMessageBox)
from PyQt5.QtCore import Qt, QTimer, pyqtSignal, QThread
from PyQt5.QtGui import QFont, QIcon
import win32gui, win32com.client, pythoncom
import win32con
import win32api
import keyboard
import random
import ctypes
from pynput.keyboard import Controller

# 获取正确的图标路径，适用于开发环境和打包环境
def resource_path(relative_path):
    """ 获取资源的绝对路径，兼容开发环境和PyInstaller打包后的环境 """
    try:
        # PyInstaller创建临时文件夹，将路径存储在_MEIPASS中
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# 图标路径
icon_path = resource_path('puling.ico')

# 设置应用ID，以便Windows在任务栏中显示正确的图标
try:
    myappid = 'puling.autopuling.1.0' 
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
except Exception:
    pass  # 如果失败，静默处理

class AutoPulingThread(QThread):
    update_log = pyqtSignal(str)
    update_progress = pyqtSignal(int)
    finished = pyqtSignal()
    
    def __init__(self, foreground, duration, parent=None):
        super().__init__(parent)
        self.foreground = foreground
        self.duration = duration
        self.stop = False
        self.game_nd = None
        self.keyboard = Controller()
        
    def run(self):
        self.update_log.emit("开始运行自动脚本...")
        self.game_nd = self.init_window()
        
        if self.game_nd is not None:
            tm = time.time()
            end_time = tm + 60 * self.duration
            self.update_log.emit("游戏窗口已找到，开始运行。")
            self.update_log.emit(f"运行时间：{self.duration} 分钟，预计结束时刻：{time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(end_time))}。")
            
            time.sleep(1)
            self.press('f', 0.2)
            recent_a = 0
            notif_tm = time.time()
            
            while not self.stop and time.time() < end_time:
                # 更新进度条
                progress = int(((time.time() - tm) / (end_time - tm)) * 100)
                self.update_progress.emit(progress)
                
                # 剩余时间提醒
                if time.time() - notif_tm > (60 * 5):  # 5分钟提醒一次
                    self.update_log.emit(f"已运行 {int((time.time() - tm) // 60)} 分钟，剩余 {self.duration - int((time.time() - tm) // 60)} 分钟。")
                    notif_tm = time.time()
                
                time.sleep(random.random() * 1 + 2)
                self.press('f', 0.2)
                
                if time.time() - recent_a > 10:
                    recent_a = time.time()
                    self.press('a', 0.2, False)
            
            self.update_progress.emit(100)
            
            if self.stop:
                self.update_log.emit("运行已停止。")
            else:
                self.update_log.emit(f"运行时间已达 {self.duration} 分钟，运行自动停止。")
            
        else:
            self.update_log.emit("未找到游戏窗口，请确保游戏正在运行。")
        
        self.finished.emit()
    
    def stop_thread(self):
        self.stop = True
    
    # 枚举所有窗口
    def enum_windows_callback(self, hwnd, hwnds):
        try:
            class_name = win32gui.GetClassName(hwnd).strip()
            window_name = win32gui.GetWindowText(hwnd).strip()
            if (window_name == 'InfinityNikki' or window_name == '无限暖暖') and class_name == 'UnrealWindow':
                hwnds.append(hwnd)
        except:
            pass
        return True
    
    # 检测窗口是否最小化
    def is_window_minimized(self, hwnd):
        placement = win32gui.GetWindowPlacement(hwnd)
        return placement[1] == win32con.SW_SHOWMINIMIZED
    
    # 强制游戏窗口处于前台
    def set_forground(self, hwnds):
        """将窗口设置为前台"""
        try:
            pythoncom.CoInitialize()
            
            # 检查并恢复窗口状态
            if self.is_window_minimized(hwnds):
                win32gui.ShowWindow(hwnds, win32con.SW_RESTORE)
                
            # 直接设置为前台
            win32gui.SetForegroundWindow(hwnds)
        except Exception as e:
            self.update_log.emit(f"设置前台失败: {e}")
    
    # 检测游戏窗口
    def init_window(self):
        hwnds = []
        win32gui.EnumWindows(lambda a, b: self.enum_windows_callback(a, b), hwnds)
        
        if len(hwnds) == 0:
            self.update_log.emit('未找到游戏窗口。')
            return None
        
        game_nd = hwnds[0]
        
        if self.foreground:
            self.set_forground(game_nd)
            
        return game_nd
    
    # 后台模式按键输入
    def send_key_to_window(self, key='f', tm=0.2, keyup=True):
        hwnd = self.game_nd
        
        # 取消最小化
        if self.is_window_minimized(hwnd):
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            
        key_code = 0x41 + ord(key) - ord('a')
        win32api.PostMessage(hwnd, win32con.WM_KEYDOWN, key_code, 0)
        time.sleep(random.random() * 0.2 + tm)
        if keyup:
            win32api.PostMessage(hwnd, win32con.WM_KEYUP, key_code, 0)
    
    # 前台模式按键输入
    def press_key_pynput(self, key, tm, keyup):
        while win32gui.GetForegroundWindow() != self.game_nd:
            time.sleep(1)
            if self.stop:
                return
                
            self.update_log.emit("尝试将游戏窗口置于前台...")
            self.game_nd = self.init_window()
            if self.game_nd is None:
                self.update_log.emit("未找到游戏窗口，停止运行。")
                self.stop = True
                return
            time.sleep(1)
            break
            
        self.keyboard.press(key)
        time.sleep(random.random() * 0.2 + tm)
        
        if keyup:
            self.keyboard.release(key)
    
    # 模拟按键输入
    def press(self, key, tm, keyup=True):
        if self.foreground:
            self.press_key_pynput(key, tm, keyup)
        else:
            self.send_key_to_window(key, tm, keyup)

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.init_ui()
        self.load_settings()
        
        # 全局热键设置
        keyboard.add_hotkey('f8', self.stop_script)
        
        self.worker_thread = None
    
    def init_ui(self):
        self.setWindowTitle("无限暖暖自动挂机")
        self.setFixedSize(600, 500)
        
        # 设置应用图标
        self.setWindowIcon(QIcon(icon_path))
        
        # 主布局
        main_widget = QWidget()
        main_layout = QVBoxLayout()
        main_widget.setLayout(main_layout)
        self.setCentralWidget(main_widget)
        
        # 运行模式选择
        mode_group = QGroupBox("运行模式")
        mode_layout = QVBoxLayout()
        
        self.foreground_radio = QRadioButton("前台模式")
        self.background_radio = QRadioButton("后台模式")
        self.mode_remember = QCheckBox("记住我的选择")
        
        mode_layout.addWidget(self.foreground_radio)
        mode_layout.addWidget(self.background_radio)
        mode_layout.addWidget(self.mode_remember)
        mode_group.setLayout(mode_layout)
        
        # 运行时间设置
        time_group = QGroupBox("运行时间")
        time_layout = QHBoxLayout()
        
        time_layout.addWidget(QLabel("运行时间 (分钟):"))
        self.duration_spin = QSpinBox()
        self.duration_spin.setRange(1, 1440)  # 1分钟到24小时
        self.duration_spin.setValue(60)
        self.duration_spin.setSingleStep(5)
        
        self.time_remember = QCheckBox("记住时间设置")
        
        time_layout.addWidget(self.duration_spin)
        time_layout.addWidget(self.time_remember)
        time_group.setLayout(time_layout)
        
        # 自动关闭设置
        close_group = QGroupBox("完成后设置")
        close_layout = QVBoxLayout()
        
        self.auto_close = QCheckBox("运行结束后自动关闭游戏")
        self.close_remember = QCheckBox("记住此设置")
        
        close_layout.addWidget(self.auto_close)
        close_layout.addWidget(self.close_remember)
        close_group.setLayout(close_layout)
        
        # 控制按钮
        control_layout = QHBoxLayout()
        
        self.start_button = QPushButton("开始运行")
        self.start_button.clicked.connect(self.start_script)
        
        self.stop_button = QPushButton("停止运行")
        self.stop_button.clicked.connect(self.stop_script)
        self.stop_button.setEnabled(False)
        
        control_layout.addWidget(self.start_button)
        control_layout.addWidget(self.stop_button)
        
        # 状态和日志
        status_group = QGroupBox("运行状态")
        status_layout = QVBoxLayout()
        
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        
        status_layout.addWidget(QLabel("进度:"))
        status_layout.addWidget(self.progress_bar)
        status_layout.addWidget(QLabel("运行日志:"))
        status_layout.addWidget(self.log_text)
        status_group.setLayout(status_layout)
        
        # 将所有组件添加到主布局
        main_layout.addWidget(mode_group)
        main_layout.addWidget(time_group)
        main_layout.addWidget(close_group)
        main_layout.addLayout(control_layout)
        main_layout.addWidget(status_group)
        
        # 初始设置
        self.background_radio.setChecked(True)
        self.add_to_log("程序已启动，请设置参数后点击\"开始运行\"按钮开始。按F8可随时停止运行。")
        
    def add_to_log(self, message):
        current_time = time.strftime("%H:%M:%S", time.localtime())
        self.log_text.append(f"[{current_time}] {message}")
        self.log_text.ensureCursorVisible()
    
    def load_settings(self):
        try:
            if os.path.exists('info.txt'):
                with open('info.txt', 'r') as f:
                    info = json.load(f)
                    
                    # 运行模式
                    foreground = info.get("foreground", 0)
                    if foreground == 1:
                        self.foreground_radio.setChecked(True)
                    else:
                        self.background_radio.setChecked(True)
                    
                    # 记住选择
                    self.mode_remember.setChecked(info.get("remember_foreground", 0) == 1)
                    self.time_remember.setChecked(info.get("remember_duration", 0) == 1)
                    self.close_remember.setChecked(info.get("remember_closure", 0) == 1)
                    
                    # 运行时间
                    self.duration_spin.setValue(info.get("duration", 60))
                    
                    # 自动关闭
                    self.auto_close.setChecked(info.get("closure", 0) == 1)
                    
                self.add_to_log("已加载保存的设置")
        except Exception as e:
            self.add_to_log(f"加载设置出错: {e}")
    
    def save_settings(self):
        info = {}
        
        # 保存运行模式
        info["foreground"] = 1 if self.foreground_radio.isChecked() else 0
        
        # 保存记住选择状态
        info["remember_foreground"] = 1 if self.mode_remember.isChecked() else 0
        info["remember_duration"] = 1 if self.time_remember.isChecked() else 0
        info["remember_closure"] = 1 if self.close_remember.isChecked() else 0
        
        # 保存运行时间
        info["duration"] = self.duration_spin.value()
        
        # 保存自动关闭设置
        info["closure"] = 1 if self.auto_close.isChecked() else 0
        
        try:
            with open('info.txt', 'w') as f:
                json.dump(info, f)
            self.add_to_log("设置已保存")
        except Exception as e:
            self.add_to_log(f"保存设置出错: {e}")
    
    def start_script(self):
        # 保存设置
        self.save_settings()
        
        # 禁用开始按钮，启用停止按钮
        self.start_button.setEnabled(False)
        self.stop_button.setEnabled(True)
        self.progress_bar.setValue(0)
        
        foreground = self.foreground_radio.isChecked()
        duration = self.duration_spin.value()
        
        # 创建并启动工作线程
        self.worker_thread = AutoPulingThread(foreground, duration)
        self.worker_thread.update_log.connect(self.add_to_log)
        self.worker_thread.update_progress.connect(self.progress_bar.setValue)
        self.worker_thread.finished.connect(self.on_script_finished)
        
        self.worker_thread.start()
    
    def stop_script(self):
        if self.worker_thread and self.worker_thread.isRunning():
            self.add_to_log("正在停止运行...")
            self.worker_thread.stop_thread()
    
    def on_script_finished(self):
        self.start_button.setEnabled(True)
        self.stop_button.setEnabled(False)
        
        # 如果需要自动关闭游戏
        if self.auto_close.isChecked() and self.worker_thread.game_nd:
            self.add_to_log("准备自动关闭游戏窗口...")
            try:
                win32gui.PostMessage(self.worker_thread.game_nd, win32con.WM_CLOSE, 0, 0)
                self.add_to_log("游戏窗口已关闭")
            except Exception as e:
                self.add_to_log(f"关闭窗口失败: {e}")
    
    def closeEvent(self, event):
        if self.worker_thread and self.worker_thread.isRunning():
            reply = QMessageBox.question(
                self, '确认退出', 
                "脚本正在运行中，确定要退出吗？",
                QMessageBox.Yes | QMessageBox.No, 
                QMessageBox.No
            )
            
            if reply == QMessageBox.Yes:
                self.worker_thread.stop_thread()
                self.worker_thread.wait(1000)  # 等待线程结束
                event.accept()
            else:
                event.ignore()
        else:
            event.accept()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setWindowIcon(QIcon(icon_path))  # 为整个应用程序设置图标
    window = MainWindow()
    window.show()
    sys.exit(app.exec_()) 