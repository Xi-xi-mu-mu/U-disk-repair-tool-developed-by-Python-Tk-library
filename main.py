import os
import shutil
import string
import subprocess
import tkinter as tk
import webbrowser
from datetime import datetime
from tkinter import *
from tkinter import messagebox, Label
from tkinter import ttk
import pystray
import win32api
import win32con
import win32file
from PIL import ImageDraw
from PIL import ImageTk, Image

"""
Python Tk库开发的U盘修复工具 Python3.11.5
U disk repair tool developed by Python Tk library
"""

global tray_icon  # 全局变量


# 最小化提示
def show_info():
    messagebox.showinfo("提示", "已缩小至系统托盘")


def crop_white_background(image_path):
    # 打开图像
    image = Image.open(image_path)
    # 转换为RGBA模式，以支持透明度
    image = image.convert("RGBA")

    # 获取图像的宽度和高度
    width, height = image.size
    # 遍历每个像素点，将白色背景的像素设为透明
    for x in range(width):
        for y in range(height):
            r, g, b, a = image.getpixel((x, y))
            if r == 255 and g == 255 and b == 255:  # 判断是否为白色背景
                image.putpixel((x, y), (255, 255, 255, 0))  # 将白色背景设为透明

    # 裁剪图像，去除透明边缘
    cropped_image = image.crop(image.getbbox())

    # 保存裁剪后的图像
    cropped_image.save("")


# 退出程序时删除这些文件
files_to_remove = [
    "fix.bat",
    "fix2.bat",
    "copy.bat",
    "killer.bat",
]


# 删除本目录文件
def remove_files(files):
    for file_path in files:
        absolute_path = os.path.abspath(file_path)
        if os.path.exists(absolute_path) and os.path.isfile(absolute_path):
            os.remove(absolute_path)
            print(f"已清除：{absolute_path}")


# 开始修复提示
def simulate_repair(drive_letter):
    print(f"开始修复U盘: {drive_letter}")


# 主程序
class Ui_MainWindow:

    def __init__(self):

        self.window = tk.Tk()

        # 定义窗口大小
        self.window_width = 830
        self.window_height = 510

        # 初始化窗口
        self.window.title("U盘修复工具")
        self.window.geometry(f"{self.window_width}x{self.window_height}")
        self.window.iconbitmap('Pictures/ICON.ico')
        self.window.configure(bg='#1B2847')  # 设置窗口背景色（#1B2847）
        self.window.overrideredirect(True) # 无边框
        self.button(self.window)  # 这是定义按钮的方法
        self.setup_combobox()  # 这是定义下拉框的方法
        self.current_time = None  # 时间
        # 在窗口初始化后立即创建并保存Text控件为类属性
        self.text_log = tk.Text(self.window, width=57, height=23, fg="#94C6FF", bg="#141E35", borderwidth=0)
        self.text_log.place(x=390, y=130, anchor=tk.NW)

        # 为了能够移动窗口，可以添加鼠标绑定事件来实现拖动功能
        self.window.x = None
        self.window.y = None

        self.window.bind("<ButtonPress-1>", self.on_mouse_press)
        self.window.bind("<B1-Motion>", self.on_mouse_drag)

        # 窗口居中显示
        self.window.update_idletasks()
        cen_x = (self.window.winfo_screenwidth() - self.window.winfo_width()) // 2
        cen_y = (self.window.winfo_screenheight() - self.window.winfo_height()) // 2
        self.window.geometry(f'+{cen_x}+{cen_y}')


        # 创建圆形图标
        img = Image.new('RGBA', (33, 33))
        draw = ImageDraw.Draw(img)
        draw.ellipse([0, 0, 30, 30], fill="#81ADE3")
        photo_image = ImageTk.PhotoImage(img)

        # 创建按钮，并设置图像、文本及它们的位置关系
        button = tk.Button(self.window,
                           image=photo_image,
                           text="×", font=("黑体", 15),
                           compound="center",  # 或者选择其他如"left", "right", "top", "bottom"
                           bd=0,
                           highlightthickness=0,
                           command=lambda: create_tray_icon(),  # 隐藏主窗口但不退出程序
                           bg='#141E35',
                           activebackground="#141E35")
        # 设置按钮的位置
        button.place(x=780, y=20, anchor=tk.NW)

        # 在窗口初始化后立即获取并打印USB设备信息
        self.update_usb_info()

        # 透明图片
        TRANSCOLOUR = 'gray'  # 'gray'为透明色（我称为心灵的窗户，你也可以用这个方法做一个启动画面）
        self.window.wm_attributes("-transparentcolor", TRANSCOLOUR)  # 设置'gray'为透明色
        photo = ImageTk.PhotoImage(file=r"Pictures/USB flash drive.png")
        label = Label(self.window, image=photo, bg=TRANSCOLOUR)
        label.place(x=70, y=90, anchor=tk.NW)

        def create_tray_icon():
            self.window.withdraw()  # 隐藏主窗口但不退出程序

            remove_files(files_to_remove)  # 退出就删除工作目录下的bat文件

            # 创建系统托盘图标
            show_info()  # 最小化提示
            # 从图片文件中加载图标
            image = Image.open("Pictures/ICON.ico")

            # 创建系统托盘菜单
            menu = (
                pystray.MenuItem("打开U盘修复", lambda: open_app()),
                pystray.MenuItem("退出", lambda: exit_app())
            )

            # 创建系统托盘图标
            global tray_icon
            tray_icon = pystray.Icon("app_name", image, "U盘修复", menu)

            # 设置图标的提示文本
            tray_icon.tooltip = "U盘修复2.0"

            # 显示系统托盘图标
            tray_icon.run()

        # 点击“打开应用”菜单项的事件处理函数
        def open_app():
            global tray_icon  # 引用全局变量

            # 关闭托盘图标
            tray_icon.stop()

            # 显示已隐藏的主窗口
            self.window.deiconify()

        # 点击“退出”菜单项的事件处理函数
        def exit_app():
            global tray_icon  # 引用全局变量

            # 关闭托盘图标并退出程序
            tray_icon.stop()
            os._exit(0)
            # 立即终止整个Python解释器

        # 绑定关闭事件，当点击关闭按钮时隐藏窗口而不是退出程序
        self.window.protocol("WM_DELETE_WINDOW", create_tray_icon)

        self.window.mainloop()

    # 获取USB设备
    def get_usb_drive_info(self):
        drives = win32api.GetLogicalDriveStrings()
        usb_drives = []

        for drive in drives.split('\000')[:-1]:  # 按空字符分割，并排除最后一个空字符串
            if win32file.GetDriveType(drive) == win32con.DRIVE_REMOVABLE:  # 检查是否是可移动驱动器（例如 USB）
                volume_name, max_len, file_system, flags, fs_name = win32api.GetVolumeInformation(drive)
                # volume_name 是卷标，max_len 是卷标的最大长度，file_system 和 fs_name 是文件系统的名称，flags 是卷的属性标志
                info = {}
                info['name'] = volume_name
                info['path'] = drive
                usb_drives.append(info)

        return usb_drives

    # 打印USB设备和插入到日志
    def update_usb_info(self):
        usb_devices = self.get_usb_drive_info()

        for device in usb_devices:
            print(f"U盘名字：{device['name']}")
            print(f"U盘路径：{device['path']}\n")

            # 获取当前时间
            self.current_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            # 然后插入日志文本
            self.text_log.insert(tk.END, "[" + self.current_time + "]" + "检测到U盘:" + device['name'] + device[
                'path'] + " 请选择修复U盘盘符!\n")
            self.text_log.see(tk.END)  # 自动滚动到文本末尾

    # 鼠标事件
    def on_mouse_press(self, event):
        self.window.x = event.x
        self.window.y = event.y

    def on_mouse_drag(self, event):
        dx = event.x - self.window.x
        dy = event.y - self.window.y
        self.window.geometry("+{x}+{y}".format(x=self.window.winfo_x() + dx, y=self.window.winfo_y() + dy))

    def create_button(self, text, command, font_size, padx, pady, x, y):
        """创建按钮并设置属性"""
        return tk.Button(self.window, text=text, command=command,
                         font=("黑体", font_size), fg="#94C6FF", bg="#141E35",
                         activebackground="#141E35", activeforeground="#94C6FF",
                         borderwidth=0, padx=padx, pady=pady, relief="flat").place(x=x, y=y, anchor=tk.NW)

    def button(self, window):
        # 按钮配置
        '''按钮文本: "开始修复"
        按钮命令: self.select_combobox_value
        字体大小: 18 像素
        水平填充: 100 像素
        垂直填充: 12 像素
        x 位置: 40 像素
        y 位置: 365 像素'''
        buttons = [
            ("开始修复", self.select_combobox_value, 18, 100, 12, 40, 365),
            ("修复图标成文件夹", self.select_combobox_value2, 15, 20, 12, 40, 440),
            ("关于", self.select_combobox_value3, 15, 20, 12, 260, 440),
            ("清除修复文件", self.select_combobox_value4, 12, 15, 10, 655, 450)
        ]

        for text, command, font_size, padx, pady, x, y in buttons:
            self.create_button(text, command, font_size, padx, pady, x, y)

        # 分割线
        tk.Frame(self.window, bg="#94C6FF", height=700, width=2).place(x=370, y=0)

        # 横向分割线
        tk.Frame(self.window, bg="#141E35", height=70, width=850).place(x=0, y=0)

        # 标题
        tk.Label(self.window, text="USB Flash Drive", font=("黑体", 15), fg="#81ADE3", bg='#141E35').place(x=30, y=22,
                                                                                                           anchor=tk.NW)

        # 下拉框

        # 下拉框颜色
        combostyle = ttk.Style()
        combostyle.theme_create('combostyle', parent='alt',
                                settings={'TCombobox':
                                    {'configure':
                                        {
                                            'foreground': '#81ADE3',  # 前景色
                                            'selectbackground': '#81ADE3',  # 选择后的背景颜色
                                            'fieldbackground': '#141E35',  # 下拉框内背景色
                                            'background': '#81ADE3',  # 下拉按钮颜色
                                        }}}
                                )
        combostyle.theme_use('combostyle')

    def setup_combobox(self):
        self.var = tk.StringVar()  # 将 var 定义为实例变量
        self.var.trace_add('write', self.on_combobox_selected)
        self.window.option_add("*TCombobox*background", '#141E35')
        self.window.option_add("*TCombobox*Foreground", "#81ADE3")
        comboxlist = ttk.Combobox(self.window, textvariable=self.var, justify=tk.CENTER, width=27, height=13,
                                  font=("黑体", 20))
        comboxlist['values'] = tuple(string.ascii_uppercase)  # 24个字母
        comboxlist.current(0)
        comboxlist.place(x=394, y=90, anchor=tk.NW)

    def on_combobox_selected(self, *args):
        selected_value = self.var.get()  # 现在可以通过 self.var 访问到 StringVar 对象
        print(f"用户选择了：{selected_value}")

    # 当按下快速修复按钮时执行
    def select_combobox_value(self, *args):
        # 获取当前时间
        self.current_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        # 然后插入日志文本
        self.text_log.insert(tk.END, "[" + self.current_time + "]" + '开始快速修复' + self.var.get() + "盘\n")
        self.text_log.see(tk.END)  # 自动滚动到文本末尾
        selected_drive = self.var.get()
        # 定义要检查的路径
        path_to_check = f"{self.var.get()}:/"
        path_exists = tk.BooleanVar(value=False)  # 修改value为False

        if os.path.exists(path_to_check):
            path_exists.set(True)  # 修改value为True
        else:
            messagebox.showerror("错误", f"路径 {path_to_check} 不存在！")  # 错误提示框
            self.text_log.insert(tk.END, "[" + self.current_time + "]" + "错误！请选择正确的盘符！\n")
            self.text_log.see(tk.END)  # 自动滚动到文本末尾
            return  # 不执行接下来的代码

        # 如果还有原来的fix.bat文件就删除原来的文件
        for filename in ["fix.bat"]:
            path = filename
            if os.path.exists(path):
                os.remove(path)
                print(f"已清除{path}")

        simulate_repair(selected_drive)  # 开始修复提示
        self.text_log.insert(tk.END, f"[{self.current_time}] 创建快速修复bat文件\n")
        self.text_log.see(tk.END)  # 自动滚动到文本末尾

        with open('fix.bat', 'w') as file:
            file.write(f"chkdsk {self.var.get()}: /f\n")
            file.write("del %0\n")
        file.close()
        log_message = f"[{self.current_time}] 创建完成\n启动快速修复文件bat\n"
        self.text_log.insert(tk.END, log_message)
        self.text_log.see(tk.END)
        # 启动copy.bat
        os.startfile("fix.bat")

    # 当按下修复图标成文件夹时执行
    def select_combobox_value2(self):
        self.current_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')  # 获取当前时间
        # 然后插入日志文本
        self.text_log.insert(tk.END, "[" + self.current_time + "]" + '开始修复图标成文件夹' + self.var.get() + "盘\n")
        self.text_log.see(tk.END)
        selected_drive = self.var.get()  # 定义
        # 定义要检查的路径
        path_to_check = f"{self.var.get()}:/"
        path_exists = tk.BooleanVar(value=False)  # 修改value为False

        if os.path.exists(path_to_check):
            path_exists.set(True)  # 修改value为True
        else:
            messagebox.showerror("错误", f"路径 {path_to_check} 不存在！")
            self.text_log.insert(tk.END, "[" + self.current_time + "]" + "错误！请选择正确的盘符！\n")
            self.text_log.see(tk.END)
            return  # 不执行接下来的代码

        #  删除修复盘的killer.bat文件
        drive_path = self.var.get() + ":/killer.bat"  # 定义删除路径
        for filename in [drive_path]:
            path = filename
            if os.path.exists(path):
                os.remove(path)  # 删除
                print(f"已清除{path}")
                self.text_log.insert(tk.END, "[" + self.current_time + "]" + "已清除" + path + "\n")
                self.text_log.see(tk.END)

        # 删除修复盘的fix2.bat文件
        drive_path = self.var.get() + ":/fix2.bat"  # 定义删除路径
        for filename in [drive_path]:
            path = filename
            if os.path.exists(path):
                os.remove(path)  # 删除
                print(f"已清除{path}")
                self.text_log.insert(tk.END, "[" + self.current_time + "]" + "已清除" + path + "\n")
                self.text_log.see(tk.END)

        # 如果有就清理目录下的文件
        for filename in ["fix2.bat", "copy.bat", "killer.bat"]:
            path = filename
            if os.path.exists(path):
                os.remove(path)
                print(f"已清除{path}")

        def write_bat_file(filename, lines):
            """将给定的行写入一个批处理文件"""
            with open(filename, 'w') as file:
                file.writelines(lines)

        def log_message(message):
            """向日志中插入一条消息并滚动到末尾"""
            self.text_log.insert(tk.END, f"[{self.current_time}] {message}\n")
            self.text_log.see(tk.END)

        simulate_repair(selected_drive)  # 开始修复提示
        log_message("创建bat文件")

        messagebox.showerror("提示",
                             "修复完成后请点击清除修复文件按钮！\n并删除U盘中的未知文件，如果无法删除请打开安全软件删除！\n删除后拔插U盘！")

        def copy_files_and_run_fix2(drive_letter):
            # 定义文件路径
            source_fix2 = 'fix2.bat'
            source_killer = 'killer.bat'
            destination_fix2 = os.path.join(drive_letter + ':\\', 'fix2.bat')
            destination_killer = os.path.join(drive_letter + ':\\', 'killer.bat')

            # 复制文件到目标位置
            shutil.copy(source_fix2, destination_fix2)
            shutil.copy(source_killer, destination_killer)

            # 启动 fix2.bat 文件
            subprocess.Popen([destination_fix2], shell=True)

        def create_killer_bat():
            killer_content = [
                'taskkill /F /IM cmd.exe\n',
                'ntsd -c q -pn excel.exe\n',
                'C:\\Documents and Settings\\Administrator>taskkill /?\n',
                'TASKKILL [/S system [/U username [/P [password]]]]\n',
                '         { [/FI filter] [/PID processid | /IM imagename] } [/F] [/T]\n'
            ]
            with open('killer.bat', 'w') as file:
                file.writelines(killer_content)

        def create_fix2_bat(drive_letter):
            fix2_content = [
                "@echo off\n",
                "setlocal EnableDelayedExpansion\n",
                "\n",
                'PUSHD %~DP0 & cd /zd "%~dp0"\n',
                "%1 %2\n",
                'mshta vbscript:createobject("shell.application").shellexecute("%~s0","goto :target","","runas",1)(window.close)&goto :eof\n',
                ":target\n",
                "\n",
                "@echo off\n",
                "@echo 本程序消除文件夹被病毒置上的隐藏属性\n",
                "@echo.\n",
                "@ECHO 可能需要一段时间，请耐心等待\n",
                "@echo 耐心等待...\n",
                "attrib -s -h *. /S /D\n",
                "attrib +s +h System~1\n",
                "attrib +s +h Recycled\n",
                "attrib +s +h +a ntldr\n",
                "@echo on\n",
                f'start "" "{drive_letter}:\\killer.bat"\n',
                "del %0\n"
            ]
            with open('fix2.bat', 'w') as file:
                file.writelines(fix2_content)

        # 使用示例
        drive_letter = self.var.get()  # 用你的驱动器字母替换
        create_fix2_bat(drive_letter)
        create_killer_bat()
        copy_files_and_run_fix2(drive_letter)

        log_message("修复完成记得点清除修复文件，并删除U盘中的未知文件，如果无法删除请打开安全软件删除！删除后拔插U盘！")

    # 清除修复盘中的killer.bat文件
    def select_combobox_value4(self):
        self.current_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')  # 获取当前时间
        drive_path = self.var.get() + ":/killer.bat"  # |定义删除killer.bat文件
        path3 = self.var.get() + ":/"  # |定义打开路径
        for filename in [drive_path]:
            path: str = filename
            if not os.path.exists(path):  # 如果没有killer文件
                messagebox.showerror("错误", "请先运行修复图标成文件夹！")
                self.text_log.insert(tk.END, "[" + self.current_time + "]" + "错误！请先运行修复图标成文件夹！\n")
                self.text_log.see(tk.END)  # 自动滚动到文本末尾
            else:
                # 删除修复盘的fix2.bat文件，虽然fix.bat执行完成后会自己删除但是万一执行到一半就关掉呢
                drive_path2 = self.var.get() + ":/fix2.bat"
                for filename2 in [drive_path2]:
                    path2 = filename2
                    if os.path.exists(path2):
                        os.remove(path2)
                        print(f"已清除{path2}")  # 删除fix2.bat
                        self.text_log.insert(tk.END, "[" + self.current_time + "]" + "删除修复盘的fix2.bat文件完成\n")
                        self.text_log.see(tk.END)
                #  删除修复盘的killer.bat文件
                os.remove(path)
                print(f"已清除{path}")  # 删除killer.bat
                self.text_log.insert(tk.END, "[" + self.current_time + "]" + "删除修复盘的killer.bat文件完成\n")
                self.text_log.see(tk.END)
                messagebox.showerror("提示", "已清除修复文件！")  # 提示框
                self.text_log.insert(tk.END, "[" + self.current_time + "]" + "启动修复U盘完成\n")
                self.text_log.see(tk.END)
                os.startfile(path3)  # 打开修复U盘

    # 关于部分
    def select_combobox_value3(self):
        def on_button_click():
            root.destroy()  # 关闭

        def open_url(event):
            url = "https://space.bilibili.com/1420530803?spm_id_from=333.1007.0.0"  # 跳转的网址
            webbrowser.open(url, new=2)  # 参数new=2表示在新的标签页中打开

        def open_url2(event):
            url = "https://github.com/xixiruirui/U-disk-repair-tool-developed-by-Python-Tk-library"
            webbrowser.open(url, new=2)

        def open_url3(event):
            url = "https://wwp.lanzout.com/b043a1b7c"
            webbrowser.open(url, new=2)

        root = tk.Tk()  # 创建窗口

        # 定义窗口大小
        root_width = 300
        root_height = 200

        # 初始化窗口
        root.title("关于")
        root.geometry(f"{root_width}x{root_height}")
        root.resizable(False, False)  # 禁止最大化
        root.iconbitmap('Pictures/ICON.ico')
        root.configure(bg='#1B2847')  # 设置窗口背景色（#1B2847）
        root.attributes('-topmost', True)  # 置顶

        def create_label(root, text, font, fg, bg, x, y, anchor=tk.NW, underline=False, bind=None):
            """创建标签并设置属性"""
            label = tk.Label(root, text=text, font=font, fg=fg, bg=bg, underline=underline)
            if bind:
                label.bind("<Button-1>", bind)
            label.place(x=x, y=y, anchor=anchor)
            return label

        pink_color = "#FF1493"
        blue_color = "#94C6FF"
        orange_color = "#FFA500"

        # 创建标签
        create_label(root, "BiliBili:", ("黑体", 15), pink_color, '#1B2847', 20, 20)

        create_label(root, "析木想搞事", ("黑体", 15, "underline"), blue_color, '#1B2847', 130, 20, bind=open_url)

        create_label(root, "开源地址：", ("黑体", 15), 'white', '#1B2847', 20, 60)

        create_label(root, "Github", ("黑体", 15, "underline"), blue_color, '#1B2847', 130, 60, bind=open_url2)

        create_label(root, "蓝奏云", ("黑体", 15, "underline"), blue_color, '#1B2847', 200, 100, bind=open_url3)

        create_label(root, "下载_密码123456：", ("黑体", 15), orange_color, '#1B2847', 20, 100)

        create_label(root, "我懒得做联网更新了\n所以用的时候点击下载看看哈qwq", ("黑体", 10), 'white', '#1B2847', 10,
                     160)

        # 添加“退出”按钮
        repair_button4 = tk.Button(root, text='退出', font=("黑体", 10), padx=10, pady=5, relief='flat', fg=blue_color,
                                   bg="#141E35", borderwidth=0, activebackground="#141E35", activeforeground=blue_color,
                                   command=on_button_click)
        repair_button4.place(x=230, y=160, anchor=tk.NW)

        # 窗口居中显示
        root.update_idletasks()
        cen_x = (root.winfo_screenwidth() - root.winfo_width()) // 2
        cen_y = (root.winfo_screenheight() - root.winfo_height()) // 2
        root.geometry(f'+{cen_x}+{cen_y}')

        root.mainloop()



if __name__ == "__main__":
    ui = Ui_MainWindow()
