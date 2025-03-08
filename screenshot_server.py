import io
import os
import sys
import json
import base64
import logging
import threading
import subprocess
import time  # 添加time模块导入
from http.server import HTTPServer, BaseHTTPRequestHandler
from PIL import ImageGrab, Image
import tkinter as tk
from tkinter import ttk, messagebox

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(os.path.join(os.path.dirname(os.path.abspath(__file__)), "screenshot_server.log")),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger("ScreenshotServer")

# 默认配置
DEFAULT_CONFIG = {
    "port": 8000,
    "quality": 80,
    "autostart": False
}

CONFIG_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")

def load_config():
    """加载配置文件"""
    try:
        if os.path.exists(CONFIG_PATH):
            with open(CONFIG_PATH, 'r', encoding='utf-8') as f:
                return json.load(f)
        return DEFAULT_CONFIG.copy()
    except Exception as e:
        logger.error(f"加载配置文件失败: {e}")
        return DEFAULT_CONFIG.copy()

def save_config(config):
    """保存配置文件"""
    try:
        with open(CONFIG_PATH, 'w', encoding='utf-8') as f:
            json.dump(config, f, indent=4, ensure_ascii=False)
        return True
    except Exception as e:
        logger.error(f"保存配置文件失败: {e}")
        return False

class ScreenshotHandler(BaseHTTPRequestHandler):
    """处理屏幕截图请求的HTTP处理器"""
    
    def log_message(self, format, *args):
        """重写日志方法，将服务器日志输出到我们的日志系统"""
        logger.info("%s - - [%s] %s" %
                     (self.address_string(),
                      self.log_date_time_string(),
                      format % args))
    
    def do_GET(self):
        """处理GET请求"""
        try:
            logger.debug(f"收到GET请求: {self.path}")
            if self.path == '/screenshot':
                self.send_screenshot()
            elif self.path == '/':
                self.send_response(200)
                self.send_header('Content-type', 'text/html')
                self.end_headers()
                self.wfile.write(b"<html><body><h1>Screenshot Server</h1><p>Use /screenshot to get a screenshot.</p></body></html>")
            else:
                self.send_response(404)
                self.send_header('Content-type', 'text/plain')
                self.end_headers()
                self.wfile.write(b"404 Not Found")
        except Exception as e:
            logger.error(f"处理GET请求时出错: {e}", exc_info=True)
            try:
                self.send_response(500)
                self.send_header('Content-type', 'text/plain')
                self.end_headers()
                self.wfile.write(f"服务器内部错误: {str(e)}".encode('utf-8'))
            except:
                logger.error("无法发送错误响应", exc_info=True)
    
    def send_screenshot(self):
        """捕获屏幕截图并发送"""
        try:
            # 获取当前配置
            config = load_config()
            quality = config.get("quality", 80)
            
            logger.debug("开始捕获屏幕")
            # 捕获屏幕
            screenshot = ImageGrab.grab()
            logger.debug(f"屏幕捕获成功，尺寸: {screenshot.size}")
            
            # 转换为JPEG格式
            img_byte_arr = io.BytesIO()
            screenshot.save(img_byte_arr, format='JPEG', quality=quality)
            img_byte_arr.seek(0)
            img_data = img_byte_arr.getvalue()
            logger.debug(f"图像转换成功，大小: {len(img_data)} 字节")
            
            # 发送响应
            self.send_response(200)
            self.send_header('Content-type', 'image/jpeg')
            self.send_header('Content-length', str(len(img_data)))
            self.end_headers()
            self.wfile.write(img_data)
            
            logger.info(f"已发送屏幕截图，质量: {quality}%")
        except Exception as e:
            logger.error(f"截图失败: {e}", exc_info=True)
            self.send_response(500)
            self.send_header('Content-type', 'text/plain')
            self.end_headers()
            error_msg = f"截图失败: {str(e)}\n{traceback.format_exc()}"
            self.wfile.write(error_msg.encode('utf-8'))

class ScreenshotServer:
    """屏幕截图服务器"""
    
    def __init__(self):
        self.config = load_config()
        self.server = None
        self.server_thread = None
    
    def start_server(self):
        """启动HTTP服务器"""
        if self.server:
            return
        
        try:
            port = self.config.get("port", 8000)
            logger.debug(f"尝试在端口 {port} 上启动服务器")
            self.server = HTTPServer(('0.0.0.0', port), ScreenshotHandler)
            self.server_thread = threading.Thread(target=self.server.serve_forever)
            self.server_thread.daemon = True
            self.server_thread.start()
            logger.info(f"服务器已启动，监听端口: {port}")
            return True
        except Exception as e:
            logger.error(f"启动服务器失败: {e}", exc_info=True)
            messagebox.showerror("错误", f"启动服务器失败: {e}")
            return False
    
    def stop_server(self):
        """停止HTTP服务器"""
        if self.server:
            # 在单独的线程中关闭服务器，避免阻塞GUI
            server_to_stop = self.server
            self.server = None
            
            def shutdown_server():
                try:
                    server_to_stop.shutdown()
                    logger.info("服务器已停止")
                except Exception as e:
                    logger.error(f"停止服务器时出错: {e}", exc_info=True)
            
            threading.Thread(target=shutdown_server, daemon=True).start()
            return True
        return False
    
    def restart_server(self):
        """重启HTTP服务器"""
        self.stop_server()
        # 等待一小段时间确保服务器完全关闭
        time.sleep(0.5)
        return self.start_server()

def setup_autostart(enable=True):
    """设置开机自启动"""
    try:
        # 获取当前脚本的绝对路径
        script_path = os.path.abspath(sys.argv[0])
        if script_path.endswith('.py'):
            # 如果是Python脚本，使用pythonw.exe运行（无控制台窗口）
            exe_path = sys.executable
            if exe_path.endswith('python.exe'):
                exe_path = exe_path.replace('python.exe', 'pythonw.exe')
            command = f'"{exe_path}" "{script_path}"'
        else:
            # 如果是已编译的可执行文件，直接运行
            command = f'"{script_path}"'
        
        task_name = "ScreenshotServer"
        
        # 创建启动目录（如果不存在）
        startup_dir = os.path.join(os.environ['APPDATA'], 'Microsoft', 'Windows', 'Start Menu', 'Programs', 'Startup')
        os.makedirs(startup_dir, exist_ok=True)
        
        # 创建快捷方式路径
        shortcut_path = os.path.join(startup_dir, f"{task_name}.bat")
        
        if enable:
            try:
                # 首先尝试使用任务计划
                cmd = [
                    'schtasks', '/create', '/tn', task_name, 
                    '/tr', command, 
                    '/sc', 'onlogon', 
                    '/rl', 'highest', 
                    '/f'
                ]
                result = subprocess.run(cmd, check=False, capture_output=True, text=True)
                
                if result.returncode != 0:
                    # 如果任务计划失败，使用启动文件夹方法
                    logger.warning(f"使用任务计划设置自启动失败: {result.stderr}，尝试使用启动文件夹方法")
                    
                    # 创建批处理文件
                    with open(shortcut_path, 'w') as f:
                        f.write(f"@echo off\n{command}")
                    
                    logger.info(f"已在启动文件夹创建自启动脚本: {shortcut_path}")
                else:
                    logger.info("已使用任务计划设置开机自启动")
                
                return True
            except Exception as e:
                # 如果任务计划方法完全失败，使用启动文件夹方法
                logger.warning(f"使用任务计划设置自启动失败: {e}，尝试使用启动文件夹方法")
                
                # 创建批处理文件
                with open(shortcut_path, 'w') as f:
                    f.write(f"@echo off\n{command}")
                
                logger.info(f"已在启动文件夹创建自启动脚本: {shortcut_path}")
                return True
        else:
            # 删除任务计划和启动文件
            try:
                subprocess.run(['schtasks', '/delete', '/tn', task_name, '/f'], check=False)
            except:
                pass
            
            # 删除启动文件夹中的脚本
            if os.path.exists(shortcut_path):
                os.remove(shortcut_path)
                
            logger.info("已取消开机自启动")
            return True
    except Exception as e:
        logger.error(f"设置开机自启动失败: {e}", exc_info=True)
        return False

class SettingsGUI:
    """设置界面"""
    
    def __init__(self, root, server):
        self.root = root
        self.server = server
        self.config = server.config
        
        self.root.title("屏幕截图服务器设置")
        self.root.geometry("400x300")
        self.root.resizable(False, False)
        
        # 创建样式
        style = ttk.Style()
        style.configure('TButton', font=('微软雅黑', 10))
        style.configure('TLabel', font=('微软雅黑', 10))
        style.configure('TCheckbutton', font=('微软雅黑', 10))
        
        # 创建主框架
        main_frame = ttk.Frame(root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 端口设置
        ttk.Label(main_frame, text="服务端口:").grid(row=0, column=0, sticky=tk.W, pady=10)
        self.port_var = tk.StringVar(value=str(self.config.get("port", 8000)))
        port_entry = ttk.Entry(main_frame, textvariable=self.port_var, width=10)
        port_entry.grid(row=0, column=1, sticky=tk.W, pady=10)
        
        # 图片质量设置
        ttk.Label(main_frame, text="图片质量:").grid(row=1, column=0, sticky=tk.W, pady=10)
        self.quality_var = tk.IntVar(value=self.config.get("quality", 80))
        quality_scale = ttk.Scale(main_frame, from_=10, to=100, orient=tk.HORIZONTAL,
                                 variable=self.quality_var, length=200)
        quality_scale.grid(row=1, column=1, sticky=tk.W, pady=10)
        self.quality_label = ttk.Label(main_frame, text=f"{self.quality_var.get()}%")
        self.quality_label.grid(row=1, column=2, sticky=tk.W, pady=10)
        quality_scale.bind("<Motion>", self.update_quality_label)
        
        # 开机自启动设置
        self.autostart_var = tk.BooleanVar(value=self.config.get("autostart", False))
        autostart_check = ttk.Checkbutton(main_frame, text="开机自启动", variable=self.autostart_var)
        autostart_check.grid(row=2, column=0, columnspan=2, sticky=tk.W, pady=10)
        
        # 服务器状态
        self.status_var = tk.StringVar(value="服务器状态: 未运行")
        status_label = ttk.Label(main_frame, textvariable=self.status_var)
        status_label.grid(row=3, column=0, columnspan=3, sticky=tk.W, pady=10)
        
        # 按钮框架
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=4, column=0, columnspan=3, pady=20)
        
        # 启动/停止服务器按钮
        self.toggle_button = ttk.Button(button_frame, text="启动服务器", command=self.toggle_server)
        self.toggle_button.pack(side=tk.LEFT, padx=5)
        
        # 保存设置按钮
        save_button = ttk.Button(button_frame, text="保存设置", command=self.save_settings)
        save_button.pack(side=tk.LEFT, padx=5)
        
        # 退出按钮
        exit_button = ttk.Button(button_frame, text="退出", command=self.exit_app)
        exit_button.pack(side=tk.LEFT, padx=5)
        
        # 更新服务器状态显示
        self.update_server_status()
        
        # 设置关闭窗口的处理
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        # 创建系统托盘图标
        self.create_tray_icon()
    
    def update_quality_label(self, event=None):
        """更新质量标签"""
        self.quality_label.config(text=f"{self.quality_var.get()}%")
    
    def toggle_server(self):
        """切换服务器状态"""
        if self.server.server:
            self.server.stop_server()
        else:
            # 先应用当前端口设置
            try:
                port = int(self.port_var.get())
                if port < 1 or port > 65535:
                    raise ValueError("端口必须在1-65535之间")
                self.config["port"] = port
                self.server.config = self.config
            except ValueError as e:
                messagebox.showerror("错误", f"无效的端口: {e}")
                return
            
            self.server.start_server()
        
        self.update_server_status()
    
    def update_server_status(self):
        """更新服务器状态显示"""
        if self.server.server:
            self.status_var.set(f"服务器状态: 运行中 (端口: {self.config.get('port', 8000)})")
            self.toggle_button.config(text="停止服务器")
        else:
            self.status_var.set("服务器状态: 未运行")
            self.toggle_button.config(text="启动服务器")
    
    def save_settings(self):
        """保存设置"""
        try:
            # 验证端口
            port = int(self.port_var.get())
            if port < 1 or port > 65535:
                raise ValueError("端口必须在1-65535之间")
            
            # 更新配置
            self.config["port"] = port
            self.config["quality"] = self.quality_var.get()
            self.config["autostart"] = self.autostart_var.get()
            
            # 保存配置
            if save_config(self.config):
                # 更新服务器配置
                self.server.config = self.config
                
                # 设置开机自启动
                setup_autostart(self.config["autostart"])
                
                # 如果服务器正在运行，重启服务器以应用新端口
                if self.server.server and port != self.server.server.server_address[1]:
                    self.server.restart_server()
                    self.update_server_status()
                
                messagebox.showinfo("成功", "设置已保存")
            else:
                messagebox.showerror("错误", "保存设置失败")
        except ValueError as e:
            messagebox.showerror("错误", f"无效的端口: {e}")
    
    def exit_app(self):
        """退出应用"""
        self.on_closing()
    
    def on_closing(self):
        """关闭窗口时的处理"""
        if messagebox.askyesno("确认", "是否退出程序？\n选择\"否\"将最小化到系统托盘。"):
            if self.server.server:
                self.server.stop_server()
            self.root.destroy()
            sys.exit(0)
        else:
            self.root.withdraw()  # 隐藏主窗口
    
    def create_tray_icon(self):
        """创建系统托盘图标"""
        try:
            import pystray
            from PIL import Image, ImageDraw
            
            # 创建图标
            icon_size = 64
            image = Image.new('RGB', (icon_size, icon_size), color=(255, 255, 255))
            draw = ImageDraw.Draw(image)
            draw.rectangle(
                [(8, 8), (icon_size - 8, icon_size - 8)],
                outline=(0, 0, 0),
                width=2
            )
            draw.rectangle(
                [(16, 16), (icon_size - 16, icon_size - 16)],
                fill=(0, 120, 215)
            )
            
            # 创建菜单
            def show_window():
                self.root.deiconify()  # 显示主窗口
                self.root.lift()       # 将窗口提升到顶层
                self.root.focus_force()  # 强制获取焦点
            
            def exit_app():
                if self.server.server:
                    self.server.stop_server()
                icon.stop()
                self.root.destroy()
                sys.exit(0)
            
            menu = (
                pystray.MenuItem('显示设置', show_window),
                pystray.MenuItem('退出', exit_app)
            )
            
            # 创建托盘图标
            icon = pystray.Icon("screenshot_server", image, "屏幕截图服务器", menu)
            
            # 在单独的线程中运行托盘图标
            threading.Thread(target=icon.run, daemon=True).start()
            
        except ImportError:
            logger.warning("未安装pystray模块，无法创建系统托盘图标")
            messagebox.showwarning("警告", "未安装pystray模块，无法创建系统托盘图标。\n请使用pip安装: pip install pystray")

def main():
    """主函数"""
    # 检查是否已经有实例在运行
    try:
        import socket
        sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        sock.bind(('localhost', 0))  # 绑定到随机端口
        sock.close()
    except socket.error:
        # 如果无法绑定，可能是另一个实例已经在运行
        logger.warning("可能已有另一个实例在运行")
        if not sys.argv or len(sys.argv) <= 1 or sys.argv[1] != "--force":
            messagebox.showwarning("警告", "屏幕截图服务器可能已经在运行。\n如果确定没有运行，请使用--force参数启动。")
            sys.exit(1)
    
    # 创建服务器实例
    server = ScreenshotServer()
    
    # 创建GUI
    root = tk.Tk()
    app = SettingsGUI(root, server)
    
    # 如果配置了自动启动服务器，则启动服务器
    if server.config.get("autostart", False):
        server.start_server()
        app.update_server_status()
    
    # 运行GUI主循环
    root.mainloop()

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        logger.error(f"程序异常: {e}", exc_info=True)
        messagebox.showerror("错误", f"程序发生异常: {e}")
        sys.exit(1)