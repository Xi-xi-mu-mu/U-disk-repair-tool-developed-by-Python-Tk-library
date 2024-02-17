import os
import string
import sys
import pystray
import subprocess
import win32api
import win32file
import win32con
import webbrowser
import tkinter as tk
from tkinter import *
from PIL import ImageTk, Image
from tkinter import messagebox, Label
from tkinter import ttk
from datetime import datetime

"""
Python Tk库开发的U盘修复工具
U disk repair tool developed by Python Tk library
"""

global tray_icon  # 全局变量

# 最小化提示
def show_info():
    messagebox.showinfo("提示", "已缩小至系统托盘")


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
        # 创建主窗口
        self.window = tk.Tk()

        # 定义窗口大小
        self.window_width = 770
        self.window_height = 450

        # 初始化窗口
        self.window.title("U盘修复工具")
        self.window.geometry(f"{self.window_width}x{self.window_height}")
        self.window.resizable(False, False) # 禁止最大化
        self.window.iconbitmap('Pictures/ICON.ico')
        self.button(self.window)  # 这是定义按钮的方法
        self.setup_combobox()  # 这是定义下拉框的方法
        self.current_time = None  # 时间
        # 在窗口初始化后立即创建并保存Text控件为类属性
        self.text_log = tk.Text(self.window, width=57, height=23)
        self.text_log.place(x=350, y=60, anchor=tk.NW)
        # 在窗口初始化后立即获取并打印USB设备信息
        self.update_usb_info()

        # 图片
        TRANSCOLOUR = 'gray'  # 'gray'为透明色（我称为心灵的窗户，你也可以用这个方法做一个启动画面）
        self.window.wm_attributes("-transparentcolor", TRANSCOLOUR)  # 设置'gray'为透明色
        photo = ImageTk.PhotoImage(file=r"Pictures/USB flash drive.png"),
        label = Label(self.window, image=photo, bg=TRANSCOLOUR)
        label.place(x=50, y=20, anchor=tk.NW)

        # 为了能够移动窗口，可以添加鼠标绑定事件来实现拖动功能
        self.window.x = None
        self.window.y = None

        self.window.bind("<ButtonPress-1>", self.on_mouse_press)
        self.window.bind("<B1-Motion>", self.on_mouse_drag)

        # 加载背景图片并显示
        img = Image.open("Pictures/background.jpg")
        photo_img = self.get_img(img.filename, self.window.winfo_width(), self.window.winfo_height())
        background_label = tk.Label(self.window, image=photo_img)
        background_label.place(x=0, y=0, relwidth=1, relheight=1)
        background_label.lower()  # 让背景图片层级在最下方

        # 窗口居中显示
        self.window.update_idletasks()
        cen_x = (self.window.winfo_screenwidth() - self.window.winfo_width()) // 2
        cen_y = (self.window.winfo_screenheight() - self.window.winfo_height()) // 2
        self.window.geometry(f'+{cen_x}+{cen_y}')

        self.window.mainloop()

    # 获取USB设备
    def get_usb_drive_info(self):
        drives = win32api.GetLogicalDriveStrings()
        usb_drives = []

        for drive in drives.split('\000')[:-1]:  # Split by null character and exclude last empty string
            if win32file.GetDriveType(drive) == win32con.DRIVE_REMOVABLE:  # Check if it's a removable drive (e.g., USB)
                volume_name, max_len, file_system, flags, fs_name = win32api.GetVolumeInformation(drive)
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

    # 背景图片
    def get_img(self, filename, width=None, height=None):
        img = Image.open(filename)
        if width and height:
            img = img.resize((width, height))
        return ImageTk.PhotoImage(img)

    # 按钮部分和分割线
    def button(self, window):
        # 添加“开始修复”按钮
        repair_button1 = tk.Button(self.window, text='开始修复', command=self.select_combobox_value,
                                   font=("黑体", 18), padx=100, pady=12, relief="flat")
        repair_button1.place(x=20, y=295, anchor=tk.NW)

        # 添加“修复图标成文件夹”按钮
        repair_button2 = tk.Button(self.window, text='修复图标成文件夹', command=self.select_combobox_value2,
                                   font=("黑体", 15), padx=20, pady=12, relief="flat")
        repair_button2.place(x=20, y=375, anchor=tk.NW)

        # 添加“设置”按钮
        repair_button3 = tk.Button(self.window, text='关于', command=self.select_combobox_value3,
                                   font=("黑体", 15), padx=20, pady=12, relief="flat")
        repair_button3.place(x=240, y=375, anchor=tk.NW)

        # 添加“清除修复文件按钮”按钮
        repair_button4 = tk.Button(self.window, text='清除修复文件', command=self.select_combobox_value4,
                                   font=("黑体", 12), padx=15, pady=10, relief="flat")
        repair_button4.place(x=615, y=380, anchor=tk.NW)

        # 分割线
        vertical = Frame(self.window, bg="white", height=450, width=2)
        vertical.place(x=340, y=0)

    # 下拉框
    def setup_combobox(self):
        self.var = tk.StringVar()  # 将 var 定义为实例变量
        self.var.trace_add('write', self.on_combobox_selected)

        comboxlist = ttk.Combobox(self.window, textvariable=self.var, justify=tk.CENTER, width=27, height=13,
                                  font=("黑体", 20))
        comboxlist['values'] = tuple(string.ascii_uppercase)  # 24个字母
        comboxlist.current(0)
        comboxlist.place(x=351, y=20, anchor=tk.NW)

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
        self.text_log.insert(tk.END, "[" + self.current_time + "]" + "创建快速修复bat文件\n")
        self.text_log.see(tk.END)  # 自动滚动到文本末尾
        file = open('fix.bat', 'w')  # 没有就创建
        file.write("chkdsk ")
        file.write(self.var.get())
        file.write(": /f\n")
        file.write("del %0\n")
        file.close()
        self.text_log.insert(tk.END, "[" + self.current_time + "]" + "创建完成\n")
        self.text_log.see(tk.END)
        os.startfile("fix.bat")  # 启动本目录下的fix.bat文件
        self.text_log.insert(tk.END, "[" + self.current_time + "]" + "启动快速修复文件bat\n")
        self.text_log.see(tk.END)

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

        simulate_repair(selected_drive)  # 开始修复提示
        self.text_log.insert(tk.END, "[" + self.current_time + "]" + "创建bat文件\n")
        self.text_log.see(tk.END)
        messagebox.showerror("提示",
                             "修复完成后请点击清除修复文件按钮！\n并删除U盘中的未知文件，如果无法删除请打开安全软件删除！\n删除后拔插U盘！")
        file = open('fix2.bat', 'w')
        # 这里显示在文件名、目录名或卷标语法不正确的命令提示符上
        file.write("@echo 显示文件名、目录名或卷标语法不正确是正常现象，不用在意！\n")
        file.write("@echo ----------------------------------------------------------\n")
        file.write("@echo ----------------------------------------------------------\n")
        file.write("@echo off\n")
        file.write("setlocal EnableDelayedExpansion\n")
        file.write("  \n")
        file.write('PUSHD %~DP0 & cd /zd "%~dp0"\n')
        file.write("%1 %2\n")
        file.write('mshta vbscript:createobject("shell.application").shellexecute("%~s0","goto :target","","runas",1)(window.close)&goto :eof\n')
        file.write(":target\n")
        file.write('  \n')
        file.write("@echo off\n")
        file.write("@echo 本程序消除文件夹被病毒置上的隐藏属性\n")
        file.write("@echo.\n")
        file.write("@ECHO 可能需要一段时间，请耐心等待\n")
        file.write("@echo 耐心等待...\n")
        file.write("attrib -s -h *. /S /D\n")
        file.write("attrib +s +h System~1\n")
        file.write("attrib +s +h Recycled\n")
        file.write("attrib +s +h +a ntldr\n")
        file.write("@echo on\n")
        file.write('start "" "')
        file.write(self.var.get())
        file.write(':/killer.bat"\n')
        file.write("del %0\n")  # 执行完成后删除自己
        file.close()
        self.text_log.insert(tk.END, "[" + self.current_time + "]" + "创建修复图标成文件夹bat文件完成\n")
        self.text_log.see(tk.END)
        # 复制和启动fix.bat和killer.bat（用bat实现，因为我懒）
        file = open('copy.bat', 'w')
        file.write('xcopy "%~dp0fix2.bat" "')
        file.write(self.var.get())
        file.write(':\ "\n')
        file.write('xcopy "%~dp0killer.bat" "')
        file.write(self.var.get())
        file.write(':\ "\n')
        # 这里用bat启动修复U盘中的fix2
        file.write('start "" "')
        file.write(self.var.get())
        file.write(':/fix2.bat"\n')
        file.write('del %0\n')  # 删除自己
        file.close()
        self.text_log.insert(tk.END, "[" + self.current_time + "]" + "创建复制bat文件完成\n")
        self.text_log.see(tk.END)
        # 启动copy.bat
        os.startfile("copy.bat")
        self.text_log.insert(tk.END, "[" + self.current_time + "]" + "启动复制bat文件完成\n")
        self.text_log.see(tk.END)
        # 结束cmd，我也不知道为什么用bat的start会显示“文件名、目录名或卷标语法不正确”所以就结束cmd窗口，但是会留下killer.bat，所以叫用户修复完成点清除修复文件
        file = open('killer.bat', 'w')  # 没有就创建
        file.write('taskkill /F /IM cmd.exe\n')
        file.write('ntsd -c q -pn excel.exe\n')
        file.write('C:\Documents and Settings\Administrator>taskkill /?\n')
        file.write('TASKKILL [/S system [/U username [/P [password]]]]\n')
        file.write('         { [/FI filter] [/PID processid | /IM imagename] } [/F] [/T]\n')
        file.close()
        self.text_log.insert(tk.END, "[" + self.current_time + "]" + "修复完成记得点清除修复文件，并删除U盘中的未知文件，如果无法删除请打开安全软件删除！删除后拔插U"
                                                                     "盘！\n")
        self.text_log.see(tk.END)  # 自动滚动到文本末尾

    # 关于部分
    def select_combobox_value3(self):
        def on_button_click():
            root.destroy()   # 关闭

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
        root.attributes('-topmost', True)  # 置顶

        pink_color = "#FF1493"  # 粉色的十六进制代码
        link_label = tk.Label(root, text="BiliBili:", font=("黑体", 15), fg=pink_color)
        link_label.place(x=20, y=20, anchor=tk.NW)

        # 添加哔哩哔哩跳转
        link_label2 = tk.Label(root, text="析木想搞事", font=("黑体", 15, "underline"), fg="blue")
        link_label2.bind("<Button-1>", open_url)
        link_label2.place(x=130, y=20, anchor=tk.NW)

        link_label3 = tk.Label(root, text="开源地址：", font=("黑体", 15))
        link_label3.place(x=20, y=60, anchor=tk.NW)

        # 添加Github跳转
        link_label3 = tk.Label(root, text="Github", font=("黑体", 15, "underline"), fg="blue")
        link_label3.bind("<Button-1>", open_url2)
        link_label3.place(x=130, y=60, anchor=tk.NW)

        link_label4 = tk.Label(root, text="开源地址：", font=("黑体", 15))
        link_label4.place(x=20, y=60, anchor=tk.NW)

        # 添加下载跳转
        link_label4 = tk.Label(root, text="蓝奏云", font=("黑体", 15, "underline"), fg="blue")
        link_label4.bind("<Button-1>", open_url3)
        link_label4.place(x=200, y=100, anchor=tk.NW)

        # 橙色的十六进制代码
        orange_color = "#FFA500"
        link_label4 = tk.Label(root, text="下载_密码123456：", font=("黑体", 15), fg=orange_color)
        link_label4.place(x=20, y=100, anchor=tk.NW)

        # 懒
        link_label5 = tk.Label(root, text="我懒得做联网更新了\n所以用的时候点击下载看看哈qwq", font=("黑体", 10,))
        link_label5.place(x=10, y=160, anchor=tk.NW)

        # 添加“退出”按钮
        repair_button4 = tk.Button(root, text='退出', font=("黑体", 10), padx=10, pady=5, relief="groove",
                                   command=on_button_click)
        repair_button4.place(x=230, y=160, anchor=tk.NW)

        # 窗口居中显示
        root.update_idletasks()
        cen_x = (root.winfo_screenwidth() - root.winfo_width()) // 2
        cen_y = (root.winfo_screenheight() - root.winfo_height()) // 2
        root.geometry(f'+{cen_x}+{cen_y}')

        root.mainloop()

    # 清除修复盘中的killer.bat文件
    def select_combobox_value4(self):
        drive_path = self.var.get() + ":/killer.bat"  # |定义删除killer.bat文件
        path3 = self.var.get() + ":/"                 # |定义打开路径
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


# 创建系统托盘图标
def create_tray_icon():
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
    # 这里使用subprocess启动主窗口，因为之前我们关闭了主窗口，所以要再次启动一边，我用这个方法，你也可以改成别的
    subprocess.Popen(["U盘修复2.0.exe"])  # 应用程序的实际路径


# 点击“退出”菜单项的事件处理函数
def exit_app():
    global tray_icon  # 引用全局变量

    # 关闭托盘图标并退出程序
    tray_icon.stop()
    sys.exit(0)
    # 正常退出


if __name__ == "__main__":
    ui = Ui_MainWindow()
    remove_files(files_to_remove)  # 退出就删除工作目录下的bat文件
    create_tray_icon()  # 退出就创建系统托盘图标
