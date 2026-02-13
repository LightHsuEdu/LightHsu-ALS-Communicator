# LightMouseCTRLMENU.py
# 多级菜单 + 眨眼/闭眼/头/嘴控制 + 紧急呼叫 + TTS
# 参数和菜单从 configMENU.ini / menuData.dat 读取，如果没有则生成默认文件

'''
--- 测试版本 ---
Python          : 3.12.10
numpy           : 2.4.1
opencv-python   : 4.13.0
mediapipe       : 0.10.32
keyboard        : 0.13.5
pyttsx3         : 2.99
Pillow          : 12.0.0
pyautogui       : 0.9.54
pygame          : 2.6.1
'''

import warnings
warnings.filterwarnings("ignore", category=UserWarning)

import sys
import os
import time
import json
import configparser
import cv2
import numpy as np
import mediapipe as mp
import ctypes
import tkinter as tk
import threading
import keyboard
import pyttsx3
from PIL import Image, ImageDraw, ImageFont
import webbrowser 
import subprocess
import pygame

try:
    import pyautogui
    pyautogui.FAILSAFE = False
    HAS_PYAUTOGUI = True
except Exception:
    HAS_PYAUTOGUI = False

# 路径
FILE_BASE_DIR = os.path.dirname(os.path.abspath(__file__))

INI_PATH = os.path.join(FILE_BASE_DIR, "configMENU.ini")
MENU_PATH = os.path.join(FILE_BASE_DIR, "menuData.dat")
SND_FILENAME = os.path.join(FILE_BASE_DIR, "001.mp3")
TXT_FILENAME = os.path.join(FILE_BASE_DIR, "textSnd.txt")
TEXT_SAVE_PATH = os.path.join(FILE_BASE_DIR, "myTypeText.txt")

mp_tasks = mp.tasks
BaseOptions = mp_tasks.BaseOptions
FaceLandmarker = mp_tasks.vision.FaceLandmarker
FaceLandmarkerOptions = mp_tasks.vision.FaceLandmarkerOptions
FaceLandmarkerResult = mp_tasks.vision.FaceLandmarkerResult
VisionRunningMode = mp_tasks.vision.RunningMode

MODEL_PATH = "face_landmarker.task"
NOSE_INDEX = 4

# 屏幕信息
user32 = ctypes.windll.user32
screen_w = user32.GetSystemMetrics(0)
screen_h = user32.GetSystemMetrics(1)

# OpenCV 窗口配置
WINDOW_NAME = "LightMouseCTRL MENU Version"
CALIB_WINDOW_W, CALIB_WINDOW_H = 800, 600
RUN_WINDOW_W, RUN_WINDOW_H = 500, 400
WINDOW_W, WINDOW_H = RUN_WINDOW_W, RUN_WINDOW_H

# 默认配置
DEFAULT_INI = {
    "Thresholds": {
        "BlinkThreshold": "0.3",
        "BlinkMinDur":   "0.1",
        "BlinkMaxDur":   "0.8",
        "MouthTrigger":  "0.4",
        "MouthCalib":    "0.5"
    },
    "Timers": {
        "CloseBack":       "1.0",
        "CloseTop":        "2.0",
        "CloseEmergency":  "10.0",
        "MouthTapMinDur":  "0.1",
        "MouthEmergency":  "6.0",
        "FaceMissingEmergency": "6.0"
    },
    "Head": {
        "HeadRangeX":     "0.2",
        "HeadTiltThresX": "0.4",
        "HeadRangeY":     "0.2",
        "HeadTiltThresY": "0.4"
    },
    "Cooldowns": {
        "BlinkDown":  "1.0",
        "EyesBack":   "1.0",
        "EyesTop":    "2.0",
        "MouthTap":   "1.0",
        "HeadLeft":   "1.0",
        "HeadRight":  "1.0",
        "HeadUp":     "2.0",
        "HeadDown":   "2.0"
    }
}

# 默认菜单数据（3级菜单）
DEFAULT_MENU_DATA = {
    "快速要求-呼吸机": "气量\n面罩",
    "快速要求-喉咙": "吸痰\n拍背",
    "★打开抖音（https://www.douyin.com/?recommend=1）": "",
    "身体位置-头部": "脸\n后脑勺\n左边\n右边",
    "身体位置-肩颈": "脖子\n左肩\n右肩",
    "身体位置-腰背": "上\n下\n左\n右",
    "身体位置-四肢": "左胳膊\n右胳膊\n左腿\n右腿",
    "身体位置-胸腹": "上\n下\n左\n右",
    "护理调整-体位床位": "坐\n躺\n翻身\n坐垫床垫",
    "护理调整-按摩力度": "轻一点\n重一点\n慢一点\n停下来",
    "护理调整-位置调整": "上\n下\n左\n右",
    "基本需求-饮食相关": "喝水\n吃饭",
    "基本需求-个人卫生": "去厕所\n擦汗\n洗澡",
    "环境调节-冷热光": "冷\n热\n开灯\n关灯\n太吵了",
    "沟通说话-常用语": "好的\n不行\n再见\n谢谢",
    "其他-！特殊功能": "朗读 textSnd.txt\n打开文字窗口"
}

# INI 读写
def ensure_ini_and_load():
    config = configparser.ConfigParser()
    if not os.path.exists(INI_PATH):
        for sec, kv in DEFAULT_INI.items():
            config[sec] = kv
        try:
            with open(INI_PATH, "w", encoding="utf-8") as f:
                config.write(f)
        except:
            pass
        return config

    config.read(INI_PATH, encoding="utf-8")
    changed = False
    for sec, kv in DEFAULT_INI.items():
        if not config.has_section(sec):
            config.add_section(sec)
            changed = True
        for k, v in kv.items():
            if not config.has_option(sec, k):
                config.set(sec, k, v)
                changed = True
    if changed:
        try:
            with open(INI_PATH, "w", encoding="utf-8") as f:
                config.write(f)
        except:
            pass
    return config

BLINK_THRESHOLD = 0.3
BLINK_MIN_DUR = 0.1
BLINK_MAX_DUR = 0.8

CLOSE_BACK_SEC = 1.0
CLOSE_TOP_SEC = 2.0
CLOSE_EMERGENCY_SEC = 10.0

MOUTH_TRIGGER_THRES = 0.4
CALIB_MOUTH_THRESHOLD = 0.5
MOUTH_TAP_MIN_DUR = 0.1
MOUTH_EMERGENCY_SEC = 6.0
FACE_MISSING_EMERGENCY_SEC = 6.0

HEAD_RANGE_X = 0.2
HEAD_TILT_THRES = 0.4
HEAD_RANGE_Y = 0.2
HEAD_TILT_THRES_Y = 0.4

COOLDOWN_BLINK_DOWN  = 1.0
COOLDOWN_EYES_BACK   = 1.0
COOLDOWN_EYES_TOP    = 2.0
COOLDOWN_MOUTH_TAP   = 1.0
COOLDOWN_HEAD_LEFT   = 1.0
COOLDOWN_HEAD_RIGHT  = 1.0
COOLDOWN_HEAD_UP     = 2.0
COOLDOWN_HEAD_DOWN   = 2.0

# 菜单上下操作模式：'blink' 用眨眼向下，'head' 用点头向下
MENU_UPDOWN_MODE = 'blink'
# MENU_UPDOWN_MODE = 'head' 

def load_parameters_from_ini():
    global BLINK_THRESHOLD, BLINK_MIN_DUR, BLINK_MAX_DUR
    global CLOSE_BACK_SEC, CLOSE_TOP_SEC, CLOSE_EMERGENCY_SEC
    global MOUTH_TRIGGER_THRES, CALIB_MOUTH_THRESHOLD, MOUTH_TAP_MIN_DUR, MOUTH_EMERGENCY_SEC
    global FACE_MISSING_EMERGENCY_SEC
    global HEAD_RANGE_X, HEAD_TILT_THRES, HEAD_RANGE_Y, HEAD_TILT_THRES_Y
    global COOLDOWN_BLINK_DOWN, COOLDOWN_EYES_BACK, COOLDOWN_EYES_TOP
    global COOLDOWN_MOUTH_TAP, COOLDOWN_HEAD_LEFT, COOLDOWN_HEAD_RIGHT
    global COOLDOWN_HEAD_UP, COOLDOWN_HEAD_DOWN

    cfg = ensure_ini_and_load()

    def getf(sec, key, default):
        try:
            return float(cfg.get(sec, key, fallback=str(default)))
        except:
            return float(default)

    BLINK_THRESHOLD = getf("Thresholds", "BlinkThreshold", 0.3)
    BLINK_MIN_DUR   = getf("Thresholds", "BlinkMinDur",   0.1)
    BLINK_MAX_DUR   = getf("Thresholds", "BlinkMaxDur",   0.8)

    MOUTH_TRIGGER_THRES   = getf("Thresholds", "MouthTrigger", 0.4)
    CALIB_MOUTH_THRESHOLD = getf("Thresholds", "MouthCalib",   0.5)

    CLOSE_BACK_SEC       = getf("Timers", "CloseBack",       1.0)
    CLOSE_TOP_SEC        = getf("Timers", "CloseTop",        2.0)
    CLOSE_EMERGENCY_SEC  = getf("Timers", "CloseEmergency", 10.0)
    MOUTH_TAP_MIN_DUR    = getf("Timers", "MouthTapMinDur",  0.1)
    MOUTH_EMERGENCY_SEC  = getf("Timers", "MouthEmergency",  6.0)
    FACE_MISSING_EMERGENCY_SEC = getf("Timers", "FaceMissingEmergency", 6.0)

    HEAD_RANGE_X     = getf("Head", "HeadRangeX",     0.2)
    HEAD_TILT_THRES  = getf("Head", "HeadTiltThresX", 0.4)
    HEAD_RANGE_Y     = getf("Head", "HeadRangeY",     0.2)
    HEAD_TILT_THRES_Y = getf("Head", "HeadTiltThresY", 0.4)

    COOLDOWN_BLINK_DOWN = getf("Cooldowns", "BlinkDown", 1.0)
    COOLDOWN_EYES_BACK  = getf("Cooldowns", "EyesBack",  1.0)
    COOLDOWN_EYES_TOP   = getf("Cooldowns", "EyesTop",   2.0)
    COOLDOWN_MOUTH_TAP  = getf("Cooldowns", "MouthTap",  1.0)
    COOLDOWN_HEAD_LEFT  = getf("Cooldowns", "HeadLeft",  1.0)
    COOLDOWN_HEAD_RIGHT = getf("Cooldowns", "HeadRight", 1.0)
    COOLDOWN_HEAD_UP    = getf("Cooldowns", "HeadUp",    2.0)
    COOLDOWN_HEAD_DOWN  = getf("Cooldowns", "HeadDown",  2.0)

# 菜单数据读写
def load_menu_data():
    if not os.path.exists(MENU_PATH):
        data = dict(DEFAULT_MENU_DATA)
        with open(MENU_PATH, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=4)
        return data
    try:
        with open(MENU_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except:
        try:
            with open(MENU_PATH, "r", encoding="utf-8") as f:
                content = f.read().replace("'", '"')
                data = json.loads(content)
        except:
            data = dict(DEFAULT_MENU_DATA)
        with open(MENU_PATH, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=4)
        return data

class MenuNode:
    def __init__(self, name, children=None, special=None):
        self.name = name
        self.children = children if children is not None else []
        self.special = special

def build_menu_tree_from_data(menu_data: dict) -> MenuNode:
    root = MenuNode("ROOT", children=[])
    top_map = {}

    for key in sorted(menu_data.keys()):
        val = menu_data[key]
        lines = [ln.strip() for ln in val.splitlines() if ln.strip()]

        if "-" in key:
            top_name, sub_name = key.split("-", 1)
        else:
            top_name, sub_name = key, None

        top_stripped = top_name.strip()
        sub_stripped = sub_name.strip() if sub_name is not None else None

        # 菜单文本以 ★ 开头，无下级菜单
        if top_stripped.startswith("★"):
            leaf = MenuNode(top_name, children=[], special=None)
            root.children.append(leaf)
            continue

        if top_name not in top_map:
            top_node = MenuNode(top_name, children=[])
            top_map[top_name] = top_node
            root.children.append(top_node)
        else:
            top_node = top_map[top_name]

        if sub_name is None:
            sub_name = "默认"
            sub_stripped = sub_name

        if sub_stripped.startswith("★"):
            leaf = MenuNode(sub_name, children=[], special=None)
            top_node.children.append(leaf)
            continue

        sub_node = None
        for ch in top_node.children:
            if ch.name == sub_name:
                sub_node = ch
                break
        if sub_node is None:
            sub_node = MenuNode(sub_name, children=[])
            top_node.children.append(sub_node)

        for text in lines:
            leaf = MenuNode(text, children=[], special=None)
            sub_node.children.append(leaf)

    return root
    
# 音频 / TTS
def thread_task(func):
    threading.Thread(target=func, daemon=True).start()

def action_run_tts(text_content: str):
    if not text_content:
        return

    def _run():
        global tts_engine
        engine = None  

        try:
            if tts_engine is not None:
                tts_engine.stop()
        except Exception:
            pass

        try:
            engine = pyttsx3.init()
            engine.setProperty('rate', 150)
            tts_engine = engine
            engine.say(text_content)
            engine.runAndWait()
        except Exception as e:         
            pass
        finally:
            if tts_engine is engine:
                tts_engine = None

    thread_task(_run)

def action_read_file():
    if not os.path.exists(TXT_FILENAME):
        return
    text_content = ""
    try:
        with open(TXT_FILENAME, 'r', encoding='utf-8') as f:
            text_content = f.read()
    except UnicodeDecodeError:
        try:
            with open(TXT_FILENAME, 'r', encoding='gbk') as f:
                text_content = f.read()
        except Exception:
            return
    except Exception:
        return
    if text_content:
        action_run_tts(text_content)

def read_txt_file_and_tts(path: str):
    if not path:
        return
    if not os.path.exists(path):
        return

    text_content = ""
    try:
        with open(path, 'r', encoding='utf-8') as f:
            text_content = f.read()
    except UnicodeDecodeError:
        try:
            with open(path, 'r', encoding='gbk') as f:
                text_content = f.read()
        except Exception:
            return
    except Exception:
        return

    if text_content:
        action_run_tts(text_content)

def action_play_alarm_loop():
    if not os.path.exists(SND_FILENAME):
        return
        
    def _run():
        stop_alarm_event.clear()

        try:
            pygame.mixer.music.load(SND_FILENAME)
            pygame.mixer.music.play(-1)

            vol_step = 0
            last_vol_time = 0

            while not stop_alarm_event.is_set():
                now = time.time()
                # 每 1 秒按一次音量增加键，最多 50 次
                if vol_step < 50 and now - last_vol_time >= 1.0:
                    try:
                        pyautogui.press('volumeup')
                    except Exception:
                        pass
                    vol_step += 1
                    last_vol_time = now

                time.sleep(0.2)

            pygame.mixer.music.stop()
            return

        except Exception:
            pass
                
        while not stop_alarm_event.is_set():
            os.system(f'start "" "{SND_FILENAME}"')
            time.sleep(10)

    thread_task(_run) 

def action_open_website(url: str):
    if not url:
        return

    def _run():
        try:
            os.system(f'start "" "{url}"')
            time.sleep(3)
            if HAS_PYAUTOGUI:
                try:
                    pyautogui.hotkey('win', 'up')
                    time.sleep(0.8)
                    pyautogui.moveTo(1365, 325, duration=0.5)
                except Exception:
                    pass
        except Exception as e:
            print(">>> 打开网站失败:", e)

    thread_task(_run)

def web_scroll_down():
    try:
        if HAS_PYAUTOGUI:
            pyautogui.scroll(-800)
        else:
            keyboard.send('pagedown')
    except Exception as e:
        print(">>> 网页滚动失败:", e)

def web_click_left():
    try:
        if HAS_PYAUTOGUI:
            pyautogui.click()
        else:
            pass
    except Exception as e:
        print(">>> 网页点击失败:", e)
        
def web_close_active_window():
    try:
        if HAS_PYAUTOGUI:
            pyautogui.hotkey('alt', 'f4')
        else:
            keyboard.send('alt+f4')
    except Exception as e:
        print(">>> 关闭浏览器失败:", e)

def start_emergency_action(reason=""):
    global is_emergency_mode, pending_emergency_ui
    global emergency_mouth_armed 
    if is_emergency_mode:
        return
    print(f"\n>>> 紧急模式 原因: {reason}")
    is_emergency_mode = True
    pending_emergency_ui = True
    # 进入紧急模式时，要求再次张嘴才退出
    emergency_mouth_armed = False
    action_play_alarm_loop()
    
def stop_emergency_action():
    global is_emergency_mode
    if not is_emergency_mode:
        return
    is_emergency_mode = False
    stop_alarm_event.set()

def cv2_add_chinese(img, text, position, color=(0, 255, 0), size=32):
    img_pil = Image.fromarray(cv2.cvtColor(img, cv2.COLOR_BGR2RGB))
    draw = ImageDraw.Draw(img_pil)
    font_path = os.path.join(os.path.dirname(FILE_BASE_DIR), "simhei.ttf")
    if not os.path.exists(font_path):
        font_path = "simhei.ttf"
    try:
        font = ImageFont.truetype(font_path, size)
    except Exception:
        font = ImageFont.load_default()
    rgb_color = (color[2], color[1], color[0])
    draw.text(position, text, font=font, fill=rgb_color)
    return cv2.cvtColor(np.asarray(img_pil), cv2.COLOR_RGB2BGR)

# 菜单 UI
class MenuApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("控制菜单")    

        menu_w = max(500, screen_w - WINDOW_W)
        menu_h = screen_h
        self.root.geometry(f"{menu_w}x{menu_h}+{WINDOW_W}+0")  
        self.root.update_idletasks()
       
        self.root.overrideredirect(False)
        self.root.wm_attributes("-topmost", True)
        self.root.withdraw()

        main_frame = tk.Frame(self.root, bg="#222222")
        main_frame.pack(fill=tk.BOTH, expand=True)

        left_frame = tk.Frame(main_frame, bg="#222222")
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)

        right_frame = tk.Frame(main_frame, bg="#222222")
        right_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)

        tk.Label(left_frame, text="当前菜单", fg="white", bg="#222222", font=("Microsoft YaHei", 32, "bold")).pack(anchor="w", padx=5, pady=5)

        self.list_left = tk.Listbox(left_frame, font=("Microsoft YaHei", 32), bg="#333333", fg="white", selectbackground="#4CAF50", activestyle="none")
        self.list_left.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        tk.Label(right_frame, text="下级菜单", fg="white", bg="#222222", font=("Microsoft YaHei", 32, "bold")).pack(anchor="w", padx=5, pady=5)

        self.list_right = tk.Listbox(right_frame, font=("Microsoft YaHei", 32), bg="#333333", fg="white", activestyle="none")
        self.list_right.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        self.info_label = tk.Label(self.root, fg="white", bg="#111111", font=("Microsoft YaHei", 14), justify="left", anchor="w")
        self.info_label.pack(fill=tk.X, padx=5, pady=5)
        self.info_label.config(text="眼睛/嘴巴 状态初始化中...")

        self.status_label = tk.Label(self.root, fg="yellow", bg="#000000", font=("Microsoft YaHei", 9), anchor="w")
        self.status_label.pack(fill=tk.X, padx=5, pady=2)
        
        self.root_node = None
        self.current_node = None
        self.current_index = 0
        self.stack = []

        self.emergency_win = None

    def show(self, root_node: MenuNode):
        self.root_node = root_node
        self.current_node = root_node
        self.current_index = 0
        self.stack = []
        self.refresh_view()
        self.root.deiconify()

    def _display_name(self, node_name: str) -> str:
        s = node_name.strip()

        if "http://" in s or "https://" in s:
            if "（" in s and "）" in s:
                s = s[:s.find("（")].rstrip()
            elif "(" in s and ")" in s:
                s = s[:s.find("(")].rstrip()
            return s or node_name

        return node_name
    
    def refresh_view(self):
        if self.current_node is None:
            return

        children = self.current_node.children

        self.list_left.delete(0, tk.END)
        for ch in children:
            marker = " ▶" if ch.children else ""
            disp = self._display_name(ch.name)
            self.list_left.insert(tk.END, disp + marker)

        if children:
            self.current_index = max(0, min(self.current_index, len(children) - 1))
            self.list_left.select_clear(0, tk.END)
            self.list_left.select_set(self.current_index)
            self.list_left.see(self.current_index)
        else:
            self.current_index = 0
    
        self.list_right.delete(0, tk.END)
        if not children:
            self.list_right.insert(tk.END, "(没有子菜单)")
        else:
            cur = children[self.current_index]
            if cur.children:
                for ch in cur.children:
                    disp = self._display_name(ch.name)
                    self.list_right.insert(tk.END, disp)
            else:
                self.list_right.insert(tk.END, "无下级菜单")
                self.list_right.insert(tk.END, "张嘴 或 头右转 = 朗读此条目")

    def move_down(self):
        if self.current_node is None or not self.current_node.children:
            return
        n = len(self.current_node.children)
        self.current_index = (self.current_index + 1) % n
        self.refresh_view()

    def move_up(self):
        if self.current_node is None or not self.current_node.children:
            return
        n = len(self.current_node.children)
        self.current_index = (self.current_index - 1) % n
        self.refresh_view()
   
    def nav_enter(self):
        if self.current_node is None or not self.current_node.children:
            return

        cur = self.current_node.children[self.current_index]
        name = cur.name
        stripped = name.strip()

        if cur.children:
            # 若是“！xxxx”形式，则只朗读 xxxx；否则不朗读
            if stripped.startswith("！"):
                speak_text = stripped[1:].lstrip()
                if speak_text:
                    action_run_tts(speak_text)

            self.stack.append((self.current_node, self.current_index))
            self.current_node = cur
            self.current_index = 0
            self.refresh_view()
            return

        # 打开文字窗口：特殊功能
        if stripped == "打开文字窗口":
            global pending_open_text_window
            pending_open_text_window = True
            return

        # 如果包含 http/https，则统一当“网站模式”处理
        if ("http://" in name) or ("https://" in name):
            url = ""
            if "（" in stripped and "）" in stripped:
                s = stripped.find("（") + 1
                e = stripped.rfind("）")
                url_candidate = stripped[s:e].strip()
                if url_candidate.startswith("http"):
                    url = url_candidate
            if not url and "(" in stripped and ")" in stripped:
                s = stripped.find("(") + 1
                e = stripped.rfind(")")
                url_candidate = stripped[s:e].strip()
                if url_candidate.startswith("http"):
                    url = url_candidate
            if not url:
                for part in name.split():
                    if part.startswith("http://") or part.startswith("https://"):
                        url = part.strip()
                        break

            if url:
                global website_mode_active, current_website_url, cv_window_shown
                website_mode_active = True
                current_website_url = url
                label = stripped
                if label and label[0] in ("★", "！"):
                    label = label[1:].lstrip()
                if "（" in label:
                    label = label[:label.find("（")].rstrip()
                elif "(" in label:
                    label = label[:label.find("(")].rstrip()
                if label:
                    action_run_tts(label)

                # 隐藏菜单窗口
                try:
                    self.root.withdraw()
                except Exception:
                    pass

                # 关闭 OpenCV 窗口
                if cv_window_shown:
                    try:
                        cv2.destroyWindow(WINDOW_NAME)
                    except Exception as e:
                        print("关闭 OpenCV 窗口失败：", e)
                    cv_window_shown = False

                # 打开网站并尝试全屏 + 移动鼠标
                action_open_website(url)
            else:
                print(">>> 打开网站：未找到 URL")
            return

        # “！xxxx” 形式：朗读 xxxx
        if stripped.startswith("！"):
            speak_text = stripped[1:].lstrip()
            if speak_text:
                action_run_tts(speak_text)
            return

        # “★xxxx” 形式（不含 URL 时）：朗读 xxxx
        if stripped.startswith("★"):
            speak_text = stripped[1:].lstrip()
            if speak_text:
                action_run_tts(speak_text)
            return

        # 是否包含 .txt 文件名
        filename = None
        if ".txt" in name:
            for part in name.split():
                if part.lower().endswith(".txt"):
                    filename = part
                    break

        if filename:
            if not os.path.isabs(filename):
                path = os.path.join(FILE_BASE_DIR, filename)
            else:
                path = filename
            read_txt_file_and_tts(path)
        else:
            action_run_tts(name)
    
    def nav_back_one(self):
        if not self.stack:
            return
        self.current_node, self.current_index = self.stack.pop()
        self.refresh_view()

    def nav_back_top(self):
        if self.root_node is None:
            return
        if not self.stack and self.current_node is self.root_node:
            return
        self.current_node = self.root_node
        self.current_index = 0
        self.stack = []
        self.refresh_view()

    def update_detection_status(self, eye_l, eye_r, mouth_score, eyes_closed, mouth_open):
        eye_state = "闭眼" if eyes_closed else "睁眼"
        mouth_state = "张嘴" if mouth_open else "闭嘴"
        text = (f"眼睛闭合程度: L={eye_l:.2f}  R={eye_r:.2f}  状态: {eye_state}    "
                f"嘴巴开合程度: {mouth_score:.2f}  状态: {mouth_state}")
        self.info_label.config(text=text)

    def show_emergency_window(self):
        if self.emergency_win is not None:
            return

        win = tk.Toplevel(self.root)
        win.title("紧急呼叫")
        w, h = 640, 480
        x = (screen_w - w) // 2
        y = (screen_h - h) // 2
        win.geometry(f"{w}x{h}+{x}+{y}")
        win.overrideredirect(True)

        try:
            self.root.wm_attributes("-topmost", True)
            win.wm_attributes("-topmost", True)
        except Exception:
            pass

        win.configure(bg="red")
        label = tk.Label(
            win,
            text="紧急呼叫模式！\n\n重新闭嘴可停止报警\n\nESC 退出程序",
            fg="white", bg="red",
            font=("Microsoft YaHei", 32, "bold"),
            justify="center"
        )
        label.pack(fill=tk.BOTH, expand=True)

        try:
            win.lift()
            win.focus_force()
        except Exception:
            pass

        self.emergency_win = win

    def update(self, emergency_flag=False):
        if emergency_flag:
            if self.emergency_win is None:
                self.show_emergency_window()
            else:
                try:
                    self.emergency_win.lift()
                except tk.TclError:
                    self.emergency_win = None
        else:
            if self.emergency_win is not None:
                try:
                    self.emergency_win.destroy()
                except Exception:
                    pass
                self.emergency_win = None
                
        try:
            self.root.update_idletasks()
            self.root.update()
        except tk.TclError:
            global should_exit
            should_exit = True

    def destroy(self):
        try:
            self.root.destroy()
        except Exception:
            pass

# 键盘布局
KEY_ROWS = [
    list("QWERTYUIOP"),
    list("ASDFGHJKL"),
    list("ZXCVBNM上下左右"),
    list("1234567890"),
    ["空格", "退格", "回车", "上一页", "下一页", "切换", "退出"]
]

class EyeMouthKeyboardWindow:
    def __init__(self, parent_root):
        self.closed = False
        self.root = parent_root
        self.win = tk.Toplevel(self.root)
        self.win.title("文字输入")

        try:
            self.win.state("zoomed")
        except tk.TclError:
            self.win.geometry(f"{screen_w}x{screen_h}+0+0")

        self.win.overrideredirect(False)
        self.win.wm_attributes("-topmost", True)
        self.win.protocol("WM_DELETE_WINDOW", self.destroy)

        # 布局：文本框 / 状态栏 / 键盘
        self.win.rowconfigure(0, weight=0)
        self.win.rowconfigure(1, weight=0)
        self.win.rowconfigure(2, weight=1)
        self.win.columnconfigure(0, weight=1)

        # 文本框
        TEXT_LINES = 6
        self.info_label = tk.Text(self.win, wrap="word", font=("Microsoft YaHei", 24), bg="#111111", fg="white", insertbackground="white", height=TEXT_LINES)
        self.info_label.grid(row=0, column=0, sticky="nsew", padx=10, pady=(10, 5))

        default_tip = ""
        if os.path.exists(TEXT_SAVE_PATH):
            try:
                with open(TEXT_SAVE_PATH, "r", encoding="utf-8") as f:
                    content = f.read()
                if content.strip():
                    self.info_label.insert("1.0", content)
                else:
                    self.info_label.insert("1.0", default_tip)
            except Exception:
                self.info_label.insert("1.0", default_tip)
        else:
            self.info_label.insert("1.0", default_tip)

        self.info_label.mark_set("insert", "end-1c")

        # Text 焦点保持，保证文字输入可用
        self.info_label.bind("<FocusOut>", self._on_text_focus_out)

        # 状态栏
        self.status_label = tk.Label(self.win, text="文字输入控制方式：眨眼=下一项，张嘴=确定", fg="yellow", bg="#111111", font=("Microsoft YaHei", 16), anchor="center", justify="center")
        self.status_label.grid(row=1, column=0, sticky="ew", padx=10, pady=5)

        # 键盘区
        self.kb_frame = tk.Frame(self.win, bg="#222222")
        self.kb_frame.grid(row=2, column=0, sticky="nsew", padx=10, pady=(5, 10))

        self.mode = "ROW"
        self.row_idx = 0
        self.col_idx = 0

        self.key_labels = []
        for r, row in enumerate(KEY_ROWS):
            row_frame = tk.Frame(self.kb_frame, bg="#222222")
            row_frame.pack(pady=2)
            row_labels = []
            for c, key in enumerate(row):
                bg_color = "#333333"
                if key == "退出":
                    bg_color = "#880000"

                lbl = tk.Label(row_frame, text=key, width=6, height=2, font=("Microsoft YaHei", 24), bg=bg_color, fg="white", bd=2, relief="raised")
                padx = 2
                if r == len(KEY_ROWS) - 1 and key == "退出":
                    padx = (60, 2)

                lbl.pack(side=tk.LEFT, padx=padx, pady=2)
                row_labels.append(lbl)
            self.key_labels.append(row_labels)

        self.refresh_highlight()
        self.win.after(500, self._init_focus)
        
        # 每 5 分钟自动保存
        self.auto_save_interval_ms = 5 * 60 * 1000
        self.win.after(self.auto_save_interval_ms, self._auto_save)

    def _init_focus(self):
        if self.closed:
            return
        try:
            self.win.lift()
            self.win.focus_force()
            self.info_label.see("end-1c")
            self.info_label.mark_set("insert", "end-1c")
            self.info_label.focus_set()
        except tk.TclError:
            pass

    def _on_text_focus_out(self, event):
        if getattr(self, "closed", False):
            return
        self.win.after(50, self._refocus_text)

    def _refocus_text(self):
        if getattr(self, "closed", False):
            return
        try:
            self.info_label.focus_set()
        except tk.TclError:
            pass

    def refresh_highlight(self):
        if getattr(self, "closed", False) or not self.key_labels:
            return

        for row in self.key_labels:
            for lbl in row:
                try:
                    key_text = lbl["text"]
                    base_bg = "#880000" if key_text == "退出" else "#333333"
                    lbl.config(bg=base_bg)
                except Exception:
                    continue

        try:
            if self.mode == "ROW":
                for lbl in self.key_labels[self.row_idx]:
                    try:
                        key_text = lbl["text"]
                        if key_text == "退出":
                            lbl.config(bg="#FF0000")
                        else:
                            lbl.config(bg="#0078D7")
                    except Exception:
                        continue
            else:
                lbl = self.key_labels[self.row_idx][self.col_idx]
                key_text = lbl["text"]
                if key_text == "退出":
                    lbl.config(bg="#FF0000")
                else:
                    lbl.config(bg="#00AA00")
        except Exception:
            return

    def press_key(self, key_label):
        if getattr(self, "closed", False):
            return

        def send(key):
            if HAS_PYAUTOGUI:
                pyautogui.press(key)
            else:
                keyboard.send(key)

        def write(text):
            if HAS_PYAUTOGUI:
                pyautogui.typewrite(text)
            else:
                keyboard.write(text)

        if key_label == "空格":
            send("space")
        elif key_label == "退格":
            send("backspace")
        elif key_label == "回车":
            send("enter")
        elif key_label == "下一页":
            send("]")
        elif key_label == "上一页":
            send("[")
        elif key_label == "退出":
            self.destroy()
            return
        elif key_label in ("上", "下", "左", "右"):
            dir_map = {"上": "up", "下": "down", "左": "left", "右": "right"}
            send(dir_map[key_label])
        elif key_label == "切换":
            if HAS_PYAUTOGUI:
                pyautogui.hotkey("ctrl", "space")
            else:
                keyboard.send("ctrl+space")
        elif key_label in "1234567890":
            send(key_label)
        else:
            write(key_label.lower())

    def do_next(self):
        if getattr(self, "closed", False):
            return
        if self.mode == "ROW":
            self.row_idx = (self.row_idx + 1) % len(KEY_ROWS)
        else:
            self.col_idx = (self.col_idx + 1) % len(KEY_ROWS[self.row_idx])
        self.refresh_highlight()

    def do_confirm(self):
        if getattr(self, "closed", False):
            return

        if self.mode == "ROW":
            self.mode = "COL"
            self.col_idx = 0
            if not getattr(self, "closed", False):
                self.refresh_highlight()
        else:
            key = KEY_ROWS[self.row_idx][self.col_idx]
            self.press_key(key)
            if not getattr(self, "closed", False):
                self.mode = "ROW"
                self.refresh_highlight()

    def update_detection_status(self, eye_l, eye_r, mouth_score, eyes_closed, mouth_open):
        if getattr(self, "closed", False):
            return
        eye_state = "闭眼" if eyes_closed else "睁眼"
        mouth_state = "张嘴" if mouth_open else "闭嘴"
        text = (
            f"眼睛闭合: L={eye_l:.2f} R={eye_r:.2f}（{eye_state}）   "
            f"嘴巴开合: {mouth_score:.2f}（{mouth_state}）   "
            f"（长闭眼/长张嘴=紧急）"
        )
        try:
            self.status_label.config(text=text)
        except Exception:
            pass

    def _auto_save(self):
        if getattr(self, "closed", False):
            return

        try:
            if self.info_label is not None:
                text_content = self.info_label.get("1.0", "end-1c")
                with open(TEXT_SAVE_PATH, "w", encoding="utf-8") as f:
                    f.write(text_content)
                self.status_label.config(text="已自动保存（每5分钟一次）")
        except Exception as e:
            print(">>> 自动保存文字窗口内容失败:", e)

        try:
            self.win.after(self.auto_save_interval_ms, self._auto_save)
        except tk.TclError:
            pass

    def destroy(self):
        if getattr(self, "closed", False):
            return
        self.closed = True

        try:
            if self.info_label is not None:
                text_content = self.info_label.get("1.0", "end-1c")
                with open(TEXT_SAVE_PATH, "w", encoding="utf-8") as f:
                    f.write(text_content)
        except Exception as e:
            print(">>> 保存文字窗口内容失败:", e)

        try:
            if self.win is not None:
                self.win.destroy()
        except Exception:
            pass

        self.key_labels = []
        self.info_label = None
        self.status_label = None
        self.win = None

# 摄像头选择
def select_camera():
    print("\n>>> 正在扫描摄像头...")
    available = []
    for i in range(3):
        cap = cv2.VideoCapture(i, cv2.CAP_DSHOW)
        if cap.isOpened():
            ret, _ = cap.read()
            if ret:
                available.append(i)
            cap.release()

    if not available:
        print(">>> 未找到任何摄像头")
        return None

    if len(available) == 1:
        print(f">>> 自动选择摄像头 ID: {available[0]}")
        return available[0]

    print(f">>> 检测到多个摄像头: {available}")
    sel = input(f"请输入要使用的摄像头 ID (例如 {available[-1]}): ").strip()
    if sel.isdigit() and int(sel) in available:
        return int(sel)
    print(f">>> 输入无效，默认使用 {available[0]}")
    return available[0]

# 全局状态
neutral_x = None
neutral_y = None
calibrating = True
calib_buffer = []
input_block_until = 0.0

face_bbox_norm = None
smooth_mouth_score = 0.0

eyes_closed_prev = False
eyes_closed_start = 0.0
eyes_emergency_triggered = False

mouth_open_prev = False
mouth_open_start = 0.0
mouth_emergency_triggered = False

no_face_start_time = 0.0
no_face_emergency_triggered = False

head_pos = 'center'
head_pos_v = 'center'

last_blink_down_time  = 0.0
last_eye_back_time    = 0.0
last_eye_top_time     = 0.0
last_mouth_tap_time   = 0.0
last_head_left_time   = 0.0
last_head_right_time  = 0.0
last_head_up_time     = 0.0
last_head_down_time   = 0.0

is_emergency_mode = False
stop_alarm_event = threading.Event()
pending_emergency_ui = False

should_exit = False

pending_menu_enter = False
pending_menu_back = False
pending_menu_top = False
pending_menu_down = False
pending_menu_up = False

last_eye_left = 0.0
last_eye_right = 0.0
last_eyes_closed = False
last_mouth_raw = 0.0
last_mouth_open = False

emergency_mouth_armed = False

keyboard_mode_active = False
kb_app = None
pending_open_text_window = False

website_mode_active = False
current_website_url = ""
pending_web_scroll_down = False
pending_web_close = False

pending_web_click = False

cv_window_shown = True

tts_engine = None
pygame.mixer.init()

def get_blendshape_score(blendshapes, name):
    for b in blendshapes:
        if b.category_name.lower() == name.lower():
            return b.score
    return 0.0

def result_callback(result: FaceLandmarkerResult, output_image: mp.Image, timestamp_ms: int):
    global neutral_x, neutral_y, calibrating, calib_buffer, input_block_until
    global face_bbox_norm, smooth_mouth_score
    global eyes_closed_prev, eyes_closed_start, eyes_emergency_triggered
    global mouth_open_prev, mouth_open_start, mouth_emergency_triggered
    global head_pos, head_pos_v
    global pending_menu_enter, pending_menu_back, pending_menu_top, pending_menu_down, pending_menu_up
    global is_emergency_mode, pending_emergency_ui
    global last_eye_left, last_eye_right, last_eyes_closed, last_mouth_raw, last_mouth_open
    global last_blink_down_time, last_eye_back_time, last_eye_top_time
    global last_mouth_tap_time, last_head_left_time, last_head_right_time
    global last_head_up_time, last_head_down_time
    global emergency_mouth_armed   
    global no_face_start_time, no_face_emergency_triggered
    global website_mode_active, pending_web_scroll_down, pending_web_close, pending_web_click
    
    now = time.time() 
          
    if not result.face_landmarks:
        face_bbox_norm = None
        
        if not calibrating:
            if no_face_start_time == 0.0:
                no_face_start_time = now
                no_face_emergency_triggered = False
            else:
                dur = now - no_face_start_time
                if (not is_emergency_mode) and (not no_face_emergency_triggered) \
                        and dur >= FACE_MISSING_EMERGENCY_SEC:
                    start_emergency_action("人脸离开画面超时")
                    no_face_emergency_triggered = True
        return        

    face_landmarks = result.face_landmarks[0]
    nose = face_landmarks[NOSE_INDEX]
    x, y = nose.x, nose.y

    xs = [lm.x for lm in face_landmarks]
    ys = [lm.y for lm in face_landmarks]
    face_bbox_norm = (min(xs), min(ys), max(xs), max(ys))

    no_face_start_time = 0.0
    no_face_emergency_triggered = False

    raw_mouth_score = 0.0
    blink_left = 0.0
    blink_right = 0.0
    if result.face_blendshapes:
        blendshapes = result.face_blendshapes[0]
        s1 = get_blendshape_score(blendshapes, 'jawOpen')
        s2 = get_blendshape_score(blendshapes, 'mouthOpen')
        raw_mouth_score = max(s1, s2)
        blink_left = get_blendshape_score(blendshapes, 'eyeBlinkLeft')
        blink_right = get_blendshape_score(blendshapes, 'eyeBlinkRight')

    smooth_mouth_score_local = 0.7 * smooth_mouth_score + 0.3 * raw_mouth_score

    globals()["smooth_mouth_score"] = smooth_mouth_score_local

    last_eye_left = blink_left
    last_eye_right = blink_right

    # 校准
    if calibrating:
        calib_buffer.append((x, y))
        if len(calib_buffer) > 30:
            calib_buffer.pop(0)
        if smooth_mouth_score_local > CALIB_MOUTH_THRESHOLD and len(calib_buffer) > 10:
            avgs = np.mean(calib_buffer, axis=0)
            neutral_x, neutral_y = float(avgs[0]), float(avgs[1])
            calibrating = False
            print(f"\n>>> 校准完成。中立点: ({neutral_x:.3f}, {neutral_y:.3f})")
            input_block_until = time.time() + 1.5
            eyes_closed_prev = False
            eyes_closed_start = 0.0
            eyes_emergency_triggered = False
            mouth_open_prev = False
            mouth_open_start = 0.0
            mouth_emergency_triggered = False
            head_pos = 'center'
            head_pos_v = 'center'
            t0 = time.time()
            last_blink_down_time = t0
            last_eye_back_time = t0
            last_eye_top_time = t0
            last_mouth_tap_time = t0
            last_head_left_time = t0
            last_head_right_time = t0
            last_head_up_time = t0
            last_head_down_time = t0
        return

    if neutral_x is None:
        return

    if time.time() < input_block_until:
        last_eyes_closed = (blink_left > BLINK_THRESHOLD) or (blink_right > BLINK_THRESHOLD)
        last_mouth_raw = smooth_mouth_score_local
        last_mouth_open = (smooth_mouth_score_local > MOUTH_TRIGGER_THRES)
        return

    # 眼睛逻辑
    eyes_closed = (blink_left > BLINK_THRESHOLD) or (blink_right > BLINK_THRESHOLD)
    last_eyes_closed = eyes_closed

    if eyes_closed:
        if not eyes_closed_prev:
            eyes_closed_start = now
            eyes_emergency_triggered = False
        else:
            dur = now - eyes_closed_start
            if (not is_emergency_mode) and (not eyes_emergency_triggered) and dur >= CLOSE_EMERGENCY_SEC:
                start_emergency_action("闭眼10秒")
                eyes_emergency_triggered = True
   
    else:
        if eyes_closed_prev:
            dur = now - eyes_closed_start
            if not is_emergency_mode:
                if website_mode_active:
                    # 网站模式
                    # 闭眼时长在 CloseBack ~ CloseTop 之间：向下滚动一次
                    if CLOSE_BACK_SEC <= dur < CLOSE_TOP_SEC:
                        pending_web_scroll_down = True
                    # 闭眼时长在 CloseTop+1秒 ~ CloseEmergency 之间：关闭浏览器
                    elif CLOSE_TOP_SEC+1 <= dur < CLOSE_EMERGENCY_SEC:
                        pending_web_close = True
                        
            
                else:
                    # 眨眼：菜单向下（ blink 模式）
                    if BLINK_MIN_DUR <= dur <= BLINK_MAX_DUR:
                        if MENU_UPDOWN_MODE == 'blink':
                            if now - last_blink_down_time >= COOLDOWN_BLINK_DOWN:
                                pending_menu_down = True
                                last_blink_down_time = now

                    # 闭眼时长在 CloseBack ~ CloseTop 之间：返回上级
                    elif CLOSE_BACK_SEC <= dur < CLOSE_TOP_SEC:
                        if now - last_eye_back_time >= COOLDOWN_EYES_BACK:
                            pending_menu_back = True
                            last_eye_back_time = now

                    # 闭眼时长在 CloseTop ~ CloseEmergency 之间：返回最上级
                    elif CLOSE_TOP_SEC <= dur < CLOSE_EMERGENCY_SEC:
                        if now - last_eye_top_time >= COOLDOWN_EYES_TOP:
                            pending_menu_top = True
                            last_eye_top_time = now

        eyes_closed_start = 0.0
        eyes_emergency_triggered = False                
                        
    eyes_closed_prev = eyes_closed

    # 嘴巴逻辑
    mouth_open = (smooth_mouth_score_local > MOUTH_TRIGGER_THRES)
    last_mouth_raw = smooth_mouth_score_local
    last_mouth_open = mouth_open

    if mouth_open:
        if not mouth_open_prev:
            mouth_open_start = now
            mouth_emergency_triggered = False

            if is_emergency_mode:
                emergency_mouth_armed = True
        else:
            dur = now - mouth_open_start
            if (not is_emergency_mode) and (not mouth_emergency_triggered) and dur >= MOUTH_EMERGENCY_SEC:
                start_emergency_action("张嘴6秒")
                mouth_emergency_triggered = True
    else:
        if mouth_open_prev:
            dur = now - mouth_open_start
            if is_emergency_mode:
                if emergency_mouth_armed:
                    stop_emergency_action()
                    emergency_mouth_armed = False
            else:
                if (not mouth_emergency_triggered) and dur >= MOUTH_TAP_MIN_DUR:
                    if now - last_mouth_tap_time >= COOLDOWN_MOUTH_TAP:
                        if website_mode_active:
                            # 网站模式下：短张嘴 = 鼠标左键点击
                            pending_web_click = True
                        else:
                            # 普通模式下：短张嘴 = 菜单/键盘“确定”
                            pending_menu_enter = True
                        last_mouth_tap_time = now

        mouth_open_start = 0.0
        mouth_emergency_triggered = False

    mouth_open_prev = mouth_open

    # 头部左右逻辑
    dx_norm = 0.0
    if HEAD_RANGE_X > 1e-4:
        dx_norm = (x - neutral_x) / HEAD_RANGE_X

    new_head_pos = 'center'
    if dx_norm > HEAD_TILT_THRES:
        new_head_pos = 'right'
    elif dx_norm < -HEAD_TILT_THRES:
        new_head_pos = 'left'

    if new_head_pos != head_pos:
        if head_pos == 'center' and not is_emergency_mode:
            if new_head_pos == 'left':
                if now - last_head_left_time >= COOLDOWN_HEAD_LEFT:
                    pending_menu_back = True
                    last_head_left_time = now
            elif new_head_pos == 'right':
                if now - last_head_right_time >= COOLDOWN_HEAD_RIGHT:
                    pending_menu_enter = True
                    last_head_right_time = now
        head_pos = new_head_pos

    # 头部上下逻辑
    dy_norm = 0.0
    if HEAD_RANGE_Y > 1e-4:
        dy_norm = (y - neutral_y) / HEAD_RANGE_Y

    new_head_pos_v = 'center'
    if dy_norm < -HEAD_TILT_THRES_Y:
        new_head_pos_v = 'up'
    elif dy_norm > HEAD_TILT_THRES_Y:
        new_head_pos_v = 'down'

    if new_head_pos_v != head_pos_v:
        if head_pos_v == 'center' and not is_emergency_mode:
            if MENU_UPDOWN_MODE == 'head':
                if new_head_pos_v == 'up':
                    if now - last_head_up_time >= COOLDOWN_HEAD_UP:
                        pending_menu_up = True
                        last_head_up_time = now
                elif new_head_pos_v == 'down':
                    if now - last_head_down_time >= COOLDOWN_HEAD_DOWN:
                        pending_menu_down = True
                        last_head_down_time = now
        head_pos_v = new_head_pos_v

# 主循环
def main():
    global should_exit, pending_menu_enter, pending_menu_back, pending_menu_top, pending_menu_down, pending_menu_up
    global pending_emergency_ui
    global pending_open_text_window  
    global keyboard_mode_active, kb_app
    global cv_window_shown
    global website_mode_active, current_website_url
    global pending_web_scroll_down, pending_web_close, pending_web_click
    
    # 程序开始：发送 Win+M，最小化所有已有窗口
    try:
        keyboard.send("win+m")
        time.sleep(0.5)
    except Exception as e:
        print(f"发送 Win+M 失败：{e}")

    print("\n>>> 正在启动...")

    # 加载 / 生成配置文件和菜单数据
    load_parameters_from_ini()
    menu_data = load_menu_data()
    menu_root_node = build_menu_tree_from_data(menu_data)

    # 选择摄像头
    cam_id = select_camera()
    if cam_id is None:
        input("按回车键退出...")
        return

    cap = cv2.VideoCapture(cam_id)
    if not cap.isOpened():
        print(">>> 摄像头打开失败")
        return

    # OpenCV 窗口：校准阶段 800x600 居中
    cv2.namedWindow(WINDOW_NAME, cv2.WINDOW_NORMAL)
    cv2.resizeWindow(WINDOW_NAME, CALIB_WINDOW_W, CALIB_WINDOW_H)
    cv2.moveWindow(
        WINDOW_NAME,
        (screen_w - CALIB_WINDOW_W) // 2,
        (screen_h - CALIB_WINDOW_H) // 2
    )
    cv2.setWindowProperty(WINDOW_NAME, cv2.WND_PROP_TOPMOST, 1)
    
    cv_window_shown = True
    
    # Tk 菜单窗口
    menu_app = MenuApp()
    menu_shown = False

    options = FaceLandmarkerOptions(
        base_options=BaseOptions(model_asset_path=MODEL_PATH),
        running_mode=VisionRunningMode.LIVE_STREAM,
        result_callback=result_callback,
        output_face_blendshapes=True,
        num_faces=1
    )

    timestamp_ms = 0
    print("\n>>> 开始校准：将脸移到画面中间，张嘴确认")

    with FaceLandmarker.create_from_options(options) as landmarker:
        while True:
            if keyboard.is_pressed('esc'):
                should_exit = True

            if should_exit:
                break

            if cv_window_shown and not keyboard_mode_active:
                try:
                    cv2.setWindowProperty(
                        WINDOW_NAME,
                        cv2.WND_PROP_TOPMOST,
                        0 if is_emergency_mode else 1
                    )
                except cv2.error:
                    pass

            if menu_app is not None:
                menu_app.update(emergency_flag=is_emergency_mode)

                if not keyboard_mode_active:
                    menu_app.update_detection_status(
                        last_eye_left, last_eye_right,
                        last_mouth_raw, last_eyes_closed, last_mouth_open
                    )

            if keyboard_mode_active and kb_app is not None:
                kb_app.update_detection_status(
                    last_eye_left, last_eye_right,
                    last_mouth_raw, last_eyes_closed, last_mouth_open
                )

                if kb_app.closed:
                    keyboard_mode_active = False
                    kb_app = None

                    try:
                        menu_app.root.deiconify()
                        menu_app.root.wm_attributes("-topmost", True)
                    except Exception:
                        pass

            if not is_emergency_mode and menu_shown and not website_mode_active:
                if keyboard_mode_active and kb_app is not None:
                    if pending_menu_down:
                        kb_app.do_next()
                        pending_menu_down = False
                    if pending_menu_enter:
                        kb_app.do_confirm()
                        pending_menu_enter = False

                    pending_menu_up = False
                    pending_menu_back = False
                    pending_menu_top = False
                else:
                    if pending_menu_down:
                        menu_app.move_down()
                        pending_menu_down = False
                    if pending_menu_up:
                        menu_app.move_up()
                        pending_menu_up = False
                    if pending_menu_back:
                        menu_app.nav_back_one()
                        pending_menu_back = False
                    if pending_menu_top:
                        menu_app.nav_back_top()
                        pending_menu_top = False
                    if pending_menu_enter:
                        menu_app.nav_enter()
                        pending_menu_enter = False

            if website_mode_active and not is_emergency_mode:
                if pending_web_scroll_down:
                    web_scroll_down()
                    pending_web_scroll_down = False

                if pending_web_click:
                    web_click_left()
                    pending_web_click = False

                if pending_web_close:
                    web_close_active_window()
                    pending_web_close = False
                    website_mode_active = False
                    
                    try:
                        menu_app.root.deiconify()
                        menu_app.root.wm_attributes("-topmost", True)
                    except Exception:
                        pass

                    if not cv_window_shown:
                        try:
                            cv2.namedWindow(WINDOW_NAME, cv2.WINDOW_NORMAL)
                            cv2.resizeWindow(WINDOW_NAME, RUN_WINDOW_W, RUN_WINDOW_H)
                            y_pos = (screen_h - RUN_WINDOW_H) // 2
                            cv2.moveWindow(WINDOW_NAME, 0, y_pos)
                            cv_window_shown = True
                        except cv2.error:
                            pass

            # 菜单“打开文字窗口”
            if pending_open_text_window and not keyboard_mode_active:
                pending_open_text_window = False
                
                try:
                    menu_app.root.withdraw()
                except Exception:
                    pass
                if cv_window_shown:
                    try:
                        cv2.destroyWindow(WINDOW_NAME)
                    except cv2.error:
                        pass
                    cv_window_shown = False

                kb_app = EyeMouthKeyboardWindow(menu_app.root)
                keyboard_mode_active = True

            # Mediapipe
            ret, frame = cap.read()
            if not ret:
                print(">>> 摄像头读取失败")
                break
            frame = cv2.flip(frame, 1)
            rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
            mp_image = mp.Image(image_format=mp.ImageFormat.SRGB, data=rgb)

            timestamp_ms += 33
            landmarker.detect_async(mp_image, timestamp_ms)
            if not keyboard_mode_active and not website_mode_active:
                if not cv_window_shown:
                    try:
                        cv2.namedWindow(WINDOW_NAME, cv2.WINDOW_NORMAL)
                        if calibrating:
                            cv2.resizeWindow(WINDOW_NAME, CALIB_WINDOW_W, CALIB_WINDOW_H)
                            cv2.moveWindow(WINDOW_NAME, (screen_w - CALIB_WINDOW_W) // 2, (screen_h - CALIB_WINDOW_H) // 2)
                        else:
                            cv2.resizeWindow(WINDOW_NAME, RUN_WINDOW_W, RUN_WINDOW_H)
                            y_pos = (screen_h - RUN_WINDOW_H) // 2
                            cv2.moveWindow(WINDOW_NAME, 0, y_pos)
                        cv_window_shown = True
                    except cv2.error:
                        cv_window_shown = False

                # 提示信息
                h, w = frame.shape[:2]
                if calibrating:
                    if face_bbox_norm:
                        x1, y1, x2, y2 = face_bbox_norm
                        cv2.rectangle(frame, (int(x1 * w), int(y1 * h)), (int(x2 * w), int(y2 * h)), (0, 255, 0), 2)
                    frame = cv2_add_chinese(frame, "将脸部移至中间，张嘴完成校准", (10, 30), (0, 255, 0), 32)
                    frame = cv2_add_chinese(frame, f"嘴巴开合程度: {smooth_mouth_score:.2f}", (10, 60), (0, 255, 255), 24)
                    frame = cv2_add_chinese(frame, f"眼睛闭合程度: L={last_eye_left:.2f} R={last_eye_right:.2f}", (10, 90), (0, 255, 255), 24)
                    cv2.rectangle(frame, (w // 2 - 60, h // 2 - 80), (w // 2 + 60, h // 2 + 80), (0, 255, 0), 1)
                else:
                    if not menu_shown:
                        try:
                            cv2.resizeWindow(WINDOW_NAME, RUN_WINDOW_W, RUN_WINDOW_H)
                            y_pos = (screen_h - RUN_WINDOW_H) // 2
                            cv2.moveWindow(WINDOW_NAME, 0, y_pos)
                        except cv2.error:
                            pass
                        menu_app.show(menu_root_node)
                        menu_shown = True

                    frame = cv2_add_chinese(frame, "系统运行中 (ESC退出)", (10, 20), (0, 255, 0), 24)
                    frame = cv2_add_chinese(frame, f"嘴巴开合程度: {smooth_mouth_score:.2f}", (10, 50), (0, 255, 255), 24)
                    frame = cv2_add_chinese(frame, f"眼睛闭合程度: L={last_eye_left:.2f} R={last_eye_right:.2f}", (10, 80), (0, 255, 255), 24)
                    if neutral_x is not None:
                        cx = int(neutral_x * w)
                        cv2.line(frame, (cx, 0), (cx, h), (150, 150, 150), 1)
                    if neutral_y is not None:
                        cy = int(neutral_y * h)
                        cv2.line(frame, (0, cy), (w, cy), (150, 150, 150), 1)

                cv2.imshow(WINDOW_NAME, frame)

            else:
                if cv_window_shown:
                    try:
                        cv2.destroyWindow(WINDOW_NAME)
                    except cv2.error:
                        pass
                    cv_window_shown = False

        if cv2.waitKey(1) & 0xFF == 27:
            should_exit = True

    stop_emergency_action()
    if menu_app is not None:
        menu_app.destroy()
    cap.release()
    cv2.destroyAllWindows()
    print("\n>>> 程序已退出。")

if __name__ == "__main__":
    main()