import os
import json
import configparser
import tkinter as tk
from tkinter import ttk, messagebox
from tkinter import filedialog
import ctypes
from datetime import datetime

# 路径定义
FILE_BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INI_PATH = os.path.join(FILE_BASE_DIR, "configMENU.ini")
MENU_PATH = os.path.join(FILE_BASE_DIR, "menuData.dat")
DEFAULT_BACKUP_DIR = r"D:\\" 
DEFAULT_BACKUP_BASE = "menuData_BACKUP"

# 默认配置（匹配 LightMouseCTRLMENU.py 参数）
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
def load_ini():
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

def save_ini_from_vars(config, vars_map):
    for sec, kv in vars_map.items():
        if not config.has_section(sec):
            config.add_section(sec)
        for k, var in kv.items():
            val = var.get()
            if isinstance(val, float):
                val = f"{val:.1f}"
            config.set(sec, k, str(val))

    with open(INI_PATH, "w", encoding="utf-8") as f:
        config.write(f)        
        
# 菜单数据(menuData.dat)读写
def load_menu_data():
    if not os.path.exists(MENU_PATH):
        return dict(DEFAULT_MENU_DATA)
    try:
        with open(MENU_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except:
        try:
            with open(MENU_PATH, "r", encoding="utf-8") as f:
                content = f.read().replace("'", '"')
                return json.loads(content)
        except:
            return dict(DEFAULT_MENU_DATA)
def save_menu_data(data):
    with open(MENU_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=4)
        
# 主界面
class ConfigEditorApp:
    def __init__(self, master):
        self.master = master
        master.title("设置编辑器")
        master.geometry("1200x800")
        master.minsize(900, 600)
        
        style = ttk.Style()
        style.configure("TNotebook.Tab", font=('Microsoft YaHei', 12))
        style.configure("Group.TLabelframe.Label", font=('Microsoft YaHei', 14, 'bold'), foreground="#0056b3")
        
        notebook = ttk.Notebook(master)
        notebook.pack(fill="both", expand=True, padx=10, pady=(10, 0))

        self.frame_ini = ttk.Frame(notebook)
        self.frame_menu = ttk.Frame(notebook)

        notebook.add(self.frame_menu, text="　　菜单数据　　")
        notebook.add(self.frame_ini, text="　　参数设置　　")
        
        self.init_ini_tab()
        self.init_menu_tab()

        btn_reset = tk.Button(master, text="重置参数和菜单为默认", bg="#FF0000", fg="white", font=('Microsoft YaHei', 14, 'bold'), height=2, command=self.reset_all_to_default)
        btn_reset.pack(fill='x', padx=10, pady=10)
        
        master.protocol("WM_DELETE_WINDOW", self.on_close)

    # 参数设置 Tab
    def init_ini_tab(self):
        canvas = tk.Canvas(self.frame_ini)
        scrollbar = ttk.Scrollbar(self.frame_ini, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        win_id = canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")

        def _configure_width(event):
            canvas.itemconfig(win_id, width=event.width)
        canvas.bind("<Configure>", _configure_width)
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        canvas.bind_all("<MouseWheel>", _on_mousewheel)

        config = load_ini()
        self.ini_vars = {"Thresholds": {}, "Timers": {}, "Head": {}, "Cooldowns": {}}

        def get_val(sec, key, type_func, default):
            try:
                return type_func(config.get(sec, key, fallback=default))
            except:
                return type_func(default)

        tv = self.ini_vars["Thresholds"]
        tv["BlinkThreshold"] = tk.DoubleVar(value=get_val("Thresholds", "BlinkThreshold", float, 0.3))
        tv["BlinkMinDur"]   = tk.DoubleVar(value=get_val("Thresholds", "BlinkMinDur",   float, 0.1))
        tv["BlinkMaxDur"]   = tk.DoubleVar(value=get_val("Thresholds", "BlinkMaxDur",   float, 0.8))
        tv["MouthTrigger"]  = tk.DoubleVar(value=get_val("Thresholds", "MouthTrigger",  float, 0.4))
        tv["MouthCalib"]    = tk.DoubleVar(value=get_val("Thresholds", "MouthCalib",    float, 0.5))

        tmv = self.ini_vars["Timers"]
        tmv["CloseBack"]       = tk.DoubleVar(value=get_val("Timers", "CloseBack",       float, 1.0))
        tmv["CloseTop"]        = tk.DoubleVar(value=get_val("Timers", "CloseTop",        float, 2.0))
        tmv["CloseEmergency"]  = tk.DoubleVar(value=get_val("Timers", "CloseEmergency",  float, 10.0))
        tmv["MouthTapMinDur"]  = tk.DoubleVar(value=get_val("Timers", "MouthTapMinDur",  float, 0.1))
        tmv["MouthEmergency"]  = tk.DoubleVar(value=get_val("Timers", "MouthEmergency",  float, 6.0))
        tmv["FaceMissingEmergency"] = tk.DoubleVar(value=get_val("Timers", "FaceMissingEmergency", float, 6.0)
)
        hv = self.ini_vars["Head"]
        hv["HeadRangeX"]     = tk.DoubleVar(value=get_val("Head", "HeadRangeX",     float, 0.2))
        hv["HeadTiltThresX"] = tk.DoubleVar(value=get_val("Head", "HeadTiltThresX", float, 0.4))
        hv["HeadRangeY"]     = tk.DoubleVar(value=get_val("Head", "HeadRangeY",     float, 0.2))
        hv["HeadTiltThresY"] = tk.DoubleVar(value=get_val("Head", "HeadTiltThresY", float, 0.4))

        cv = self.ini_vars["Cooldowns"]
        cv["BlinkDown"] = tk.DoubleVar(value=get_val("Cooldowns", "BlinkDown", float, 1.0))
        cv["EyesBack"]  = tk.DoubleVar(value=get_val("Cooldowns", "EyesBack",  float, 1.0))
        cv["EyesTop"]   = tk.DoubleVar(value=get_val("Cooldowns", "EyesTop",   float, 2.0))
        cv["MouthTap"]  = tk.DoubleVar(value=get_val("Cooldowns", "MouthTap",  float, 1.0))
        cv["HeadLeft"]  = tk.DoubleVar(value=get_val("Cooldowns", "HeadLeft",  float, 1.0))
        cv["HeadRight"] = tk.DoubleVar(value=get_val("Cooldowns", "HeadRight", float, 1.0))
        cv["HeadUp"]    = tk.DoubleVar(value=get_val("Cooldowns", "HeadUp",    float, 2.0))
        cv["HeadDown"]  = tk.DoubleVar(value=get_val("Cooldowns", "HeadDown",  float, 2.0))
        
        self.slider_labels = {}
        
        def add_slider(parent, label, var, from_, to, step, desc):
            f = ttk.Frame(parent)
            f.pack(fill='x', pady=8)
            header = ttk.Frame(f)
            header.pack(fill='x')
            ttk.Label(header, text=label, font=('Microsoft YaHei', 14, 'bold')).pack(side='left')
            ttk.Label(header, text=f"  {desc}", foreground='gray', font=('Microsoft YaHei', 12)).pack(side='left')

            row = ttk.Frame(f)
            row.pack(fill='x', pady=2)
            val_lbl = ttk.Label(row, text=f"{var.get():.1f}", width=8, anchor='e', font=('Arial', 14))
            val_lbl.pack(side='right')

            self.slider_labels[id(var)] = val_lbl

            def update_lbl(v):
                val = float(v)
                val_lbl.config(text=f"{val:.1f}")
                var.set(val)

            scale = ttk.Scale(row, from_=from_, to=to, variable=var, command=update_lbl)
            scale.pack(side='left', fill='x', expand=True, padx=(0, 10))

        grp_th = ttk.LabelFrame(scrollable_frame, text="　阈值设置　", style="Group.TLabelframe", padding=10)
        grp_th.pack(fill='x', padx=10, pady=10)
        add_slider(grp_th, "眨眼阈值", tv["BlinkThreshold"], 0.1, 1.0, 0.1, "越小越容易被判定为闭眼")
        add_slider(grp_th, "眨眼最小时长", tv["BlinkMinDur"], 0.1, 0.5, 0.01, "小于此时间不算有效眨眼")
        add_slider(grp_th, "眨眼最大时长", tv["BlinkMaxDur"], 0.3, 1.5, 0.01, "大于此时间不算短眨眼")
        add_slider(grp_th, "张嘴判定阈值", tv["MouthTrigger"], 0.1, 1.0, 0.01, "张嘴程度超过此值认为是张嘴")
        add_slider(grp_th, "校准时张嘴阈值", tv["MouthCalib"], 0.1, 1.0, 0.01, "校准阶段的张嘴灵敏度")

        grp_tm = ttk.LabelFrame(scrollable_frame, text="　时间参数　", style="Group.TLabelframe", padding=10)
        grp_tm.pack(fill='x', padx=10, pady=10)
        add_slider(grp_tm, "闭眼1秒返回上级", tmv["CloseBack"], 0.5, 3.0, 0.1, "闭眼超过此时间且小于上一级时长 -> 返回上级")
        add_slider(grp_tm, "闭眼2秒返回最上级", tmv["CloseTop"], 1.0, 5.0, 0.1, "闭眼超过此时间且小于紧急时长 -> 回顶层菜单")
        add_slider(grp_tm, "长闭眼触发紧急", tmv["CloseEmergency"], 4.0, 30.0, 0.5, "闭眼超过此时间触发紧急呼叫")
        add_slider(grp_tm, "短张嘴最小时长", tmv["MouthTapMinDur"], 0.05, 1.0, 0.05, "张嘴时长超过此值且未触发紧急 -> 视为一次点击")
        add_slider(grp_tm, "长张嘴触发紧急", tmv["MouthEmergency"], 2.0, 15.0, 0.5, "张嘴超过此时间触发紧急呼叫")
        add_slider(grp_tm, "人脸丢失触发紧急", tmv["FaceMissingEmergency"], 4.0, 30.0, 0.5, "人脸离开画面超过此秒数触发紧急呼叫")

        grp_hd = ttk.LabelFrame(scrollable_frame, text="　头部范围　", style="Group.TLabelframe", padding=10)
        grp_hd.pack(fill='x', padx=10, pady=10)
        add_slider(grp_hd, "左右范围", hv["HeadRangeX"], 0.05, 0.4, 0.01, "脸左右移动归一化范围，越小越敏感")
        add_slider(grp_hd, "左右触发阈值", hv["HeadTiltThresX"], 0.1, 0.8, 0.01, "超过此偏移认为是头左/右转")
        add_slider(grp_hd, "上下范围", hv["HeadRangeY"], 0.05, 0.4, 0.01, "脸上下移动归一化范围，越小越敏感")
        add_slider(grp_hd, "上下触发阈值", hv["HeadTiltThresY"], 0.1, 0.8, 0.01, "超过此偏移认为是头抬头/低头")

        grp_cd = ttk.LabelFrame(scrollable_frame, text="　动作冷却时间　", style="Group.TLabelframe", padding=10)
        grp_cd.pack(fill='x', padx=10, pady=10)
        add_slider(grp_cd, "眨眼向下间隔", cv["BlinkDown"], 0.2, 3.0, 0.1, "两次眨眼向下之间的最小间隔")
        add_slider(grp_cd, "闭眼返回上级间隔", cv["EyesBack"], 0.2, 3.0, 0.1, "两次\"闭眼返回上级\"之间的最小间隔")
        add_slider(grp_cd, "闭眼回顶层间隔", cv["EyesTop"], 0.2, 3.0, 0.1, "两次\"闭眼回顶层\"之间的最小间隔")
        add_slider(grp_cd, "张嘴点击间隔", cv["MouthTap"], 0.2, 3.0, 0.1, "两次短张嘴点击之间的最小间隔")
        add_slider(grp_cd, "头左转返回间隔", cv["HeadLeft"], 0.2, 3.0, 0.1, "两次头左转返回之间的最小间隔")
        add_slider(grp_cd, "头右转进入间隔", cv["HeadRight"], 0.2, 3.0, 0.1, "两次头右转进入之间的最小间隔")
        add_slider(grp_cd, "头抬起向上间隔", cv["HeadUp"], 0.2, 5.0, 0.1, "两次头抬起向上之间的最小间隔")
        add_slider(grp_cd, "头低下向下间隔", cv["HeadDown"], 0.2, 5.0, 0.1, "两次头低下向下之间的最小间隔")

        def on_save():
            try:
                save_ini_from_vars(config, self.ini_vars)
                messagebox.showinfo("成功", "参数已保存到 configMENU.ini")
            except Exception as e:
                messagebox.showerror("错误", str(e))

        tk.Button(scrollable_frame, text="保存所有参数到 configMENU.ini", bg="#4CAF50", fg="white", font=('Microsoft YaHei', 14, 'bold'), height=2, command=on_save).pack(fill='x', padx=20, pady=20)

    # 菜单数据 Tab
    def init_menu_tab(self):
        self.menu_data = load_menu_data()
        paned = tk.PanedWindow(self.frame_menu, orient=tk.HORIZONTAL, sashwidth=5)
        paned.pack(fill='both', expand=True, padx=5, pady=5)

        frame_left = tk.Frame(paned)
        sb = tk.Scrollbar(frame_left)
        sb.pack(side='right', fill='y')
        self.listbox = tk.Listbox(frame_left, width=24, font=('Microsoft YaHei', 14), yscrollcommand=sb.set)
        self.listbox.pack(side='left', fill='both', expand=True)
        sb.config(command=self.listbox.yview)

        for k in sorted(self.menu_data.keys()):
            self.listbox.insert(tk.END, k)
        self.listbox.bind("<<ListboxSelect>>", self.on_menu_select)
        paned.add(frame_left)

        frame_right = tk.Frame(paned)
        tk.Label(frame_right, text="菜单项目包含“-”为菜单分级；　修改现有项目即可新建；　以！或★开头的为特殊朗读/控制；\n包含 http/https 的在主程序中会按“打开网站”处理；　”打开文字窗口“是特殊词不可修改",
                 anchor='w', justify='left', font=('Microsoft YaHei', 14, 'bold')).pack(fill='x', padx=5, pady=5)
        self.entry_key = tk.Entry(frame_right, font=('Microsoft YaHei', 14))
        self.entry_key.pack(fill='x', padx=5)

        tk.Label(frame_right, text="对应内容：可多行，每行一句菜单文本",
                 anchor='w', font=('Microsoft YaHei', 14, 'bold')).pack(fill='x', padx=5, pady=5)
        self.text_val = tk.Text(frame_right, height=12, font=('Microsoft YaHei', 14))
        self.text_val.pack(fill='both', expand=True, padx=5, pady=5)

        btn_frame = tk.Frame(frame_right)
        btn_frame.pack(fill='x', pady=10)

        tk.Button(btn_frame, text="确认修改", bg="#2196F3", fg="white", font=('Microsoft YaHei', 14, 'bold'), command=self.save_menu_item).pack(side='left', padx=5)
        tk.Button(btn_frame, text="删除条目", bg="#F44336", fg="white", font=('Microsoft YaHei', 14, 'bold'), command=self.del_menu_item).pack(side='left', padx=5)
 
        def get_default_backup_filename():
            today = datetime.now().strftime("%Y%m%d")
            return f"{DEFAULT_BACKUP_BASE}_{today}.dat"
 
        def save_backup():
            default_filename = get_default_backup_filename()
            filepath = filedialog.asksaveasfilename(
                initialdir=DEFAULT_BACKUP_DIR,
                initialfile=default_filename,
                defaultextension=".dat",
                filetypes=[("Data files", "*.dat"), ("All files", "*.*")],
                title="备份当前菜单数据"
            )

            if not filepath:
                return

            try:
                with open(filepath, "w", encoding="utf-8") as f:
                    json.dump(self.menu_data, f, ensure_ascii=False, indent=4)

                messagebox.showinfo("备份成功", f"已将当前编辑状态的数据备份到：\n{filepath}")
            except Exception as e:
                messagebox.showerror("备份失败", f"无法备份：\n{str(e)}")

        def load_backup():
            filepath = filedialog.askopenfilename(
                initialdir=DEFAULT_BACKUP_DIR,
                defaultextension=".dat",
                filetypes=[("Data files", "*.dat"), ("All files", "*.*")],
                title="选择要导入的菜单备份文件"
            )

            if not filepath:
                return

            if not messagebox.askyesno("确认覆盖", f"将从以下文件导入并覆盖当前菜单：\n{filepath}\n\n" "未保存的改动会丢失，确定继续？"):
                return

            try:
                with open(filepath, "r", encoding="utf-8") as f:
                    backup_data = json.load(f)

                self.menu_data = backup_data

                self.listbox.delete(0, tk.END)
                for k in sorted(self.menu_data.keys()):
                    self.listbox.insert(tk.END, k)

                self.entry_key.delete(0, tk.END)
                self.text_val.delete("1.0", tk.END)

                messagebox.showinfo("导入成功", "菜单数据已从备份恢复\n记得点击下方「保存所有菜单数据到 menuData.dat」按钮")
            except json.JSONDecodeError:
                messagebox.showerror("导入失败", "备份文件格式错误，手动删除 menuData.dat 后会自动创建默认数据")
            except Exception as e:
                messagebox.showerror("导入失败", f"无法读取备份文件：\n{str(e)}")

        tk.Button(btn_frame, text="备份数据", bg="#4CAF50", fg="white", font=('Microsoft YaHei', 14, 'bold'), command=save_backup).pack(side='right', padx=5)
        tk.Button(btn_frame, text="导入数据", bg="#2196F3", fg="white", font=('Microsoft YaHei', 14, 'bold'), command=load_backup).pack(side='right', padx=5)      
        tk.Button(frame_right, text="保存所有菜单数据到 menuData.dat", bg="#4CAF50", fg="white", height=2, font=('Microsoft YaHei', 14, 'bold'), command=self.save_menu_file).pack(fill='x', padx=10, pady=20)

        paned.add(frame_right)

    def on_menu_select(self, event):
        sel = self.listbox.curselection()
        if not sel:
            return
        key = self.listbox.get(sel[0])
        val = self.menu_data.get(key, "")
        self.entry_key.delete(0, tk.END)
        self.entry_key.insert(0, key)
        self.text_val.delete("1.0", tk.END)
        self.text_val.insert("1.0", val)

    def save_menu_item(self):
        key = self.entry_key.get().strip()
        val = self.text_val.get("1.0", tk.END).strip()
        if not key:
            return
        self.menu_data[key] = val
        keys_sorted = sorted(self.menu_data.keys())
        self.listbox.delete(0, tk.END)
        for k in keys_sorted:
            self.listbox.insert(tk.END, k)
        idx = keys_sorted.index(key)
        self.listbox.select_set(idx)
        self.listbox.see(idx)

    def del_menu_item(self):
        key = self.entry_key.get().strip()
        if key in self.menu_data:
            del self.menu_data[key]
            self.entry_key.delete(0, tk.END)
            self.text_val.delete("1.0", tk.END)
            self.listbox.delete(0, tk.END)
            for k in sorted(self.menu_data.keys()):
                self.listbox.insert(tk.END, k)

    def save_menu_file(self):
        save_menu_data(self.menu_data)
        messagebox.showinfo("成功", "菜单数据已保存到 menuData.dat")

    def reset_parameters_to_default(self):
        for sec, kv in DEFAULT_INI.items():
            if sec in self.ini_vars:
                for k, v_str in kv.items():
                    if k in self.ini_vars[sec]:
                        var = self.ini_vars[sec][k]
                        try:
                            v = float(v_str)
                        except:
                            v = v_str
                        if isinstance(var, tk.DoubleVar):
                            var.set(float(v))
                        else:
                            var.set(float(v))
                        lbl = self.slider_labels.get(id(var))

                        if lbl is not None:
                            val = var.get()
                            if isinstance(var, tk.DoubleVar):
                                lbl.config(text=f"{val:.1f}")
                            else:
                                lbl.config(text=str(int(val)))
        cfg = configparser.ConfigParser()
        for sec, kv in DEFAULT_INI.items():
            cfg[sec] = kv
        with open(INI_PATH, "w", encoding="utf-8") as f:
            cfg.write(f)

    def reset_menu_to_default(self):
        self.menu_data = dict(DEFAULT_MENU_DATA)
        save_menu_data(self.menu_data)
        self.listbox.delete(0, tk.END)
        for k in sorted(self.menu_data.keys()):
            self.listbox.insert(tk.END, k)
        self.entry_key.delete(0, tk.END)
        self.text_val.delete("1.0", tk.END)

    def on_close(self):
        try:
            file_data = load_menu_data()
        except Exception:
            file_data = {}
            
        if file_data != self.menu_data:
            ans = messagebox.askyesnocancel(
                "未保存的更改",
                "有未保存的菜单更改，是否保存到 menuData.dat？\n\n"
                "是：保存并退出\n"
                "否：不保存直接退出\n"
                "取消：返回继续编辑"
            )

            if ans is None:
                return

            if ans is True:
                try:
                    self.save_menu_file()
                except Exception as e:
                    messagebox.showerror("保存失败", f"保存 menuData.dat 失败：\n{e}")
                    return
        self.master.destroy()        
        
    def reset_all_to_default(self):
        if not messagebox.askyesno("确认",
                                   "确定要将所有参数和菜单重置为默认吗？\n此操作会覆盖当前 configMENU.ini 和 menuData.dat。"):
            return
        self.reset_parameters_to_default()
        self.reset_menu_to_default()
        messagebox.showinfo("完成", "已将参数和菜单重置为默认。")

if __name__ == "__main__":
    root = tk.Tk()
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except:
        pass
    app = ConfigEditorApp(root)
    root.mainloop()