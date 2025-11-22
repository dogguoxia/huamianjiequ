import tkinter as tk
from tkinter import ttk, filedialog
import win32gui
import win32con
import ctypes
from PIL import ImageGrab
import time
import os

try:
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
except Exception:
    ctypes.windll.user32.SetProcessDPIAware()


class ScreenCaptureTool:
    def __init__(self, root):
        self.root = root
        self.root.title("截图工具 v1.0.9")
        self.root.geometry("300x290")
        self.root.attributes("-topmost", True)
        self.window_dict = {}
        self.save_path = os.getcwd()
        self.is_auto_capturing = False

        self.frame = tk.Frame(root)
        self.frame.pack(expand=True, fill=tk.BOTH, padx=10, pady=10)

        self.win_label = tk.Label(self.frame, text="选择截图窗口:", anchor="w")
        self.win_label.pack(fill=tk.X, pady=(0, 2))

        self.window_list_cb = ttk.Combobox(self.frame, state="readonly")
        self.window_list_cb.pack(fill=tk.X, pady=(0, 10))

        self.path_label = tk.Label(self.frame, text="图片保存路径:", anchor="w")
        self.path_label.pack(fill=tk.X, pady=(0, 2))

        self.path_frame = tk.Frame(self.frame)
        self.path_frame.pack(fill=tk.X, pady=(0, 10))

        self.path_var = tk.StringVar(value=self.save_path)
        self.path_entry = tk.Entry(self.path_frame, textvariable=self.path_var, state="readonly")
        self.path_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)

        self.path_btn = tk.Button(self.path_frame, text="...", width=3, command=self.select_path)
        self.path_btn.pack(side=tk.RIGHT, padx=(5, 0))

        self.btn_frame = tk.Frame(self.frame)
        self.btn_frame.pack(fill=tk.X)

        self.refresh_btn = tk.Button(self.btn_frame, text="刷新", command=self.refresh_windows)
        self.refresh_btn.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(0, 2))

        self.capture_btn = tk.Button(self.btn_frame, text="截图", command=self.capture_window)
        self.capture_btn.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(2, 2))

        self.open_btn = tk.Button(self.btn_frame, text="打开目录", command=self.open_folder)
        self.open_btn.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(2, 0))

        self.auto_frame = tk.Frame(self.frame)
        self.auto_frame.pack(fill=tk.X, pady=(10, 0))

        self.auto_btn = tk.Button(self.auto_frame, text="开始自动截图 (10s/张)", command=self.toggle_auto_capture)
        self.auto_btn.pack(fill=tk.X)

        self.status_label = tk.Label(self.frame, text="准备就绪", fg="gray", wraplength=280)
        self.status_label.pack(pady=(15, 0), fill=tk.X)

        self.refresh_windows()

    def select_path(self):
        path = filedialog.askdirectory()
        if path:
            self.save_path = path
            self.path_var.set(path)

    def open_folder(self):
        try:
            os.startfile(self.save_path)
        except Exception as e:
            self.status_label.config(text=f"打开失败: {str(e)}", fg="red")

    def refresh_windows(self):
        self.window_dict.clear()
        win32gui.EnumWindows(self._enum_window_callback, None)
        sorted_titles = sorted(self.window_dict.keys())
        self.window_list_cb['values'] = sorted_titles
        if sorted_titles:
            self.window_list_cb.current(0)

    def _enum_window_callback(self, hwnd, _):
        if win32gui.IsWindowVisible(hwnd) and win32gui.GetWindowText(hwnd):
            rect = win32gui.GetWindowRect(hwnd)
            if rect[2] - rect[0] > 0 and rect[3] - rect[1] > 0:
                title = win32gui.GetWindowText(hwnd)
                if "截图工具" not in title:
                    self.window_dict[title] = hwnd

    def toggle_auto_capture(self):
        if not self.is_auto_capturing:
            self.is_auto_capturing = True
            self.auto_btn.config(text="停止自动截图")
            self.status_label.config(text="自动截图已启动", fg="blue")
            self.auto_capture_loop()
        else:
            self.is_auto_capturing = False
            self.auto_btn.config(text="开始自动截图 (10s/张)")
            self.status_label.config(text="自动截图已停止", fg="blue")

    def auto_capture_loop(self):
        if self.is_auto_capturing:
            try:
                self.capture_window()
            except Exception as e:
                self.status_label.config(text=f"自动截图异常: {str(e)}", fg="red")
            finally:
                if self.is_auto_capturing:
                    self.root.after(10000, self.auto_capture_loop)

    def capture_window(self):
        if not self.is_auto_capturing:
            self.status_label.config(text="正在处理...", fg="blue")
            self.root.update()

        selected_title = self.window_list_cb.get()
        if not selected_title:
            self.status_label.config(text="未选择窗口", fg="red")
            return

        hwnd = self.window_dict.get(selected_title)
        if not hwnd:
            self.status_label.config(text="窗口句柄无效", fg="red")
            return

        try:
            self.root.attributes("-alpha", 0)

            if win32gui.IsIconic(hwnd):
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)

            try:
                win32gui.SetForegroundWindow(hwnd)
            except Exception:
                pass

            time.sleep(0.2)

            rect = win32gui.GetWindowRect(hwnd)
            bbox = (rect[0], rect[1], rect[2], rect[3])

            img = ImageGrab.grab(bbox)

            self.root.attributes("-alpha", 1.0)

            i = 1
            while True:
                filename = f"test{i}.png"
                full_path = os.path.join(self.save_path, filename)
                if not os.path.exists(full_path):
                    break
                i += 1

            img.save(full_path)
            self.status_label.config(text=f"已保存: {filename}", fg="green")

        except Exception as e:
            self.root.attributes("-alpha", 1.0)
            self.status_label.config(text=f"出错: {str(e)}", fg="red")


if __name__ == "__main__":
    root = tk.Tk()
    app = ScreenCaptureTool(root)
    root.mainloop()