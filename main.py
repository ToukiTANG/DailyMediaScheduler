import ctypes
import locale
import os
import subprocess
import sys
import time
from datetime import datetime

import pythoncom
import win32api
import win32com.client
import win32con
import win32gui
import win32process

# ================== 配置区域 ==================
VIDEO_PATHS = [
    r"C:\Users\Touki\Desktop\temp\100727\BOCCHI_NUDE.mp4",
    r"C:\Users\Touki\Desktop\temp\100727\BOCCHI_NUDE.mp4"
]

PPT_PATHS = [
    r"C:\Users\Touki\Desktop\temp\test.pptx",
    r"C:\Users\Touki\Desktop\temp\test.pptx"
]

PLAYER_PATH = r"D:\SysTools\PotPlayer\PotPlayerMini64.exe"


# =============================================

# 获取系统编码
def get_system_encoding():
    try:
        return locale.getpreferredencoding()
    except:
        return 'utf-8'


# 日志函数 - 修复乱码问题
def log(message):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    log_message = f"[{timestamp}] {message}"

    # 控制台输出
    try:
        print(log_message)
    except UnicodeEncodeError:
        print(log_message.encode('utf-8', errors='replace').decode('utf-8'))

    # 文件日志 - 使用系统编码
    try:
        with open("scheduler_log.txt", "a", encoding='utf-8') as log_file:
            log_file.write(log_message.encode('utf-8', errors='replace').decode('utf-8') + "\n")
    except UnicodeEncodeError:
        print(log_message.encode('utf-8', errors='replace').decode('utf-8'))


# 检查管理员权限
def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False


class DisplayManager:
    @staticmethod
    def set_display_mode(mode):
        try:
            if mode == 0:
                subprocess.run(["DisplaySwitch.exe", "/internal"], check=True)
                log("显示模式已切换到: 仅电脑屏幕")
            elif mode == 1:
                subprocess.run(["DisplaySwitch.exe", "/clone"], check=True)
                log("显示模式已切换到: 复制")
        except Exception as e:
            log(f"显示模式切换失败: {str(e)}")


class MediaController:
    def __init__(self):
        self.player_process = None
        self.powerpoint = None
        self.presentation = None

    def play_video(self, video_path):
        self.close_all()
        try:
            if os.path.exists(PLAYER_PATH) and os.path.exists(video_path):
                # 使用PotPlayer全屏循环播放
                self.player_process = subprocess.Popen(
                    [
                        PLAYER_PATH,
                        video_path,
                        "/play",  # 立即开始播放
                        "/repeat",  # 循环播放
                        "/fullscreen"  # 全屏模式
                    ],
                    creationflags=subprocess.CREATE_NO_WINDOW
                )
                log(f"开始循环播放视频: {os.path.basename(video_path)}")

                # 等待播放器启动并确保全屏
                time.sleep(3)

                # 确保窗口在前台
                self._ensure_foreground("PotPlayer")
            else:
                log(f"错误: 播放器或视频文件不存在 - Player: {PLAYER_PATH}, Video: {video_path}")
        except Exception as e:
            log(f"视频播放失败: {str(e)}")

    def play_ppt(self, ppt_path):
        self.close_all()
        try:
            if os.path.exists(ppt_path):
                pythoncom.CoInitialize()
                self.powerpoint = win32com.client.Dispatch("PowerPoint.Application")
                self.powerpoint.Visible = True
                self.powerpoint.WindowState = 2  # ppWindowMinimized
                self.presentation = self.powerpoint.Presentations.Open(ppt_path, WithWindow=True)
                self.presentation.SlideShowSettings.Run()
                log(f"开始播放PPT: {os.path.basename(ppt_path)}")
                time.sleep(5)
                self._ensure_foreground("幻灯片放映")
            else:
                log(f"错误: PPT文件不存在 - {ppt_path}")
        except Exception as e:
            log(f"PPT播放失败: {str(e)}")

    def close_all(self):
        try:
            if self.player_process:
                self.player_process.terminate()
                self.player_process = None

            if self.presentation:
                self.presentation.Close()
                self.presentation = None

            if self.powerpoint:
                self.powerpoint.Quit()
                self.powerpoint = None

            # 强制终止残留进程
            subprocess.run(["taskkill", "/f", "/im", "PotPlayerMini64.exe"],
                           stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            subprocess.run(["taskkill", "/f", "/im", "POWERPNT.EXE"],
                           stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        except:
            pass

    def _ensure_foreground(self, window_title):
        """更健壮的方法确保窗口在前台"""
        try:
            def callback(hwnd, hwnds):
                if win32gui.IsWindowVisible(hwnd) and window_title in win32gui.GetWindowText(hwnd):
                    hwnds.append(hwnd)
                return True

            hwnds = []
            win32gui.EnumWindows(callback, hwnds)

            if hwnds:
                hwnd = hwnds[0]

                # 方法3: 使用更可靠的前景窗口设置方法
                self._set_foreground(hwnd)

                log(f"已成功激活窗口: {window_title}")
            else:
                log(f"警告: 未找到窗口 - {window_title}")
        except Exception as e:
            log(f"窗口激活失败: {str(e)}")

    def _set_foreground(self, hwnd):
        try:
            fg_window = win32gui.GetForegroundWindow()
            fg_thread = win32process.GetWindowThreadProcessId(fg_window)[0]
            target_thread = win32process.GetWindowThreadProcessId(hwnd)[0]

            if fg_thread != target_thread:
                ctypes.windll.user32.AttachThreadInput(fg_thread, target_thread, True)

            # 恢复窗口
            if win32gui.IsIconic(hwnd):
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            else:
                win32gui.ShowWindow(hwnd, win32con.SW_SHOW)

            # 窗口置顶再取消置顶，触发激活
            win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, 0, 0, 0, 0,
                                  win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
            win32gui.SetWindowPos(hwnd, win32con.HWND_NOTOPMOST, 0, 0, 0, 0,
                                  win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)

            # 允许设置前台窗口（关键！）
            ctypes.windll.user32.AllowSetForegroundWindow(-1)

            # ALT键辅助激活
            win32api.keybd_event(win32con.VK_MENU, 0, 0, 0)
            time.sleep(0.05)
            win32api.keybd_event(win32con.VK_MENU, 0, win32con.KEYEVENTF_KEYUP, 0)

            # 尝试前台激活
            win32gui.SetForegroundWindow(hwnd)

            if fg_thread != target_thread:
                ctypes.windll.user32.AttachThreadInput(fg_thread, target_thread, False)

            # 模拟鼠标点击窗口中心
            rect = win32gui.GetWindowRect(hwnd)
            x = (rect[0] + rect[2]) // 2
            y = (rect[1] + rect[3]) // 2
            win32api.SetCursorPos((x, y))
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0)
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0)

            log(f"窗口激活成功: hwnd={hwnd}")
            return True

        except Exception as e:
            log(f"高级窗口激活失败: {str(e)}")
            return False


class DailyScheduler:
    def __init__(self):
        self.media = MediaController()
        self.running = True

    def run_schedule(self):
        log("调度器开始运行...")
        while self.running:
            now = datetime.now().time()
            hour, minute = now.hour, now.minute

            if hour == 7 and minute == 30:
                log("7:30 - 切换到复制模式并播放视频")
                DisplayManager.set_display_mode(1)
                self.media.play_video(VIDEO_PATHS[0])

            elif hour == 8 and minute == 30:
                log("8:30 - 关闭视频并播放PPT")
                self.media.play_ppt(PPT_PATHS[0])

            elif hour == 20 and minute == 46:
                log("11:30 - 关闭PPT并播放视频")
                self.media.play_video(VIDEO_PATHS[0])

            elif hour == 12 and minute == 30:
                log("12:30 - 关闭视频并播放PPT")
                self.media.play_ppt(PPT_PATHS[1])

            elif hour == 17 and minute == 30:
                log("17:30 - 关闭PPT并播放视频")
                self.media.play_video(VIDEO_PATHS[1])

            elif hour == 18 and minute == 0:
                log("18:00 - 关闭视频并切换到仅电脑屏幕")
                self.media.close_all()
                DisplayManager.set_display_mode(0)
                self.running = False

            time.sleep(50)

    def start(self):
        log("每日媒体调度器启动")
        try:
            self.run_schedule()
        except KeyboardInterrupt:
            log("程序被用户中断")
        except Exception as e:
            log(f"发生未处理异常: {str(e)}")
        finally:
            log("清理资源...")
            self.media.close_all()
            DisplayManager.set_display_mode(0)
            log("程序已停止")


# 创建系统启动任务
def create_startup_task():
    try:
        script_path = os.path.abspath(__file__)
        task_name = "DailyMediaScheduler"

        # 获取Pythonw.exe路径
        pythonw_path = sys.executable.replace("python.exe", "pythonw.exe")
        if not os.path.exists(pythonw_path):
            pythonw_path = sys.executable  # 回退到python.exe

        # 创建任务命令
        cmd = f'schtasks /create /tn "{task_name}" /tr "\"{pythonw_path}\" \"{script_path}\"" /sc daily /st 07:25 /f'
        log(f"创建任务命令: {cmd}")

        # 执行命令
        result = subprocess.run(cmd, shell=True, capture_output=True, text=True)

        if result.returncode == 0:
            log(f"计划任务创建成功: {task_name}")
            return True
        else:
            log(f"计划任务创建失败: {result.stderr}")
            return False
    except Exception as e:
        log(f"创建任务时出错: {str(e)}")
        return False


# ================== 主程序 ==================
if __name__ == "__main__":
    # 设置工作目录为脚本所在目录
    os.chdir(os.path.dirname(os.path.abspath(__file__)))

    # 设置控制台编码为UTF-8（解决控制台乱码）
    try:
        if sys.stdout.encoding != 'utf-8':
            sys.stdout.reconfigure(encoding='utf-8')
    except:
        pass

    # 检查是否已配置
    if not os.path.exists("scheduler_configured.flag"):
        log("首次运行 - 开始配置")

        # 检查管理员权限
        if not is_admin():
            log("错误: 需要管理员权限创建计划任务")
            log("请右键点击脚本 -> '以管理员身份运行'")
            input("按回车键退出...")
            sys.exit(1)

        # 创建任务
        if create_startup_task():
            with open("scheduler_configured.flag", "w") as f:
                f.write("1")
            log("配置完成，程序将在后台运行")

            # 启动调度器
            scheduler = DailyScheduler()
            scheduler.start()
        else:
            log("计划任务创建失败，请尝试手动创建")
            log("手动创建步骤:")
            log("1. 打开'任务计划程序'")
            log("2. 创建新任务")
            log(f"3. 名称: DailyMediaScheduler")
            log("4. 触发器: 每天 7:25")
            log(f"5. 操作: 启动程序 '{sys.executable.replace('python.exe', 'pythonw.exe')}'")
            log(f"6. 参数: \"{os.path.abspath(__file__)}\"")
            log("7. 勾选'使用最高权限运行'")
            input("按回车键退出...")
    else:
        log("检测到已配置，启动调度器")
        scheduler = DailyScheduler()
        scheduler.start()
