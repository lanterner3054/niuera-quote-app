"""exe 入口:启动本地服务并自动打开浏览器。PyInstaller 用它做单文件 exe。"""
import os, sys, threading, time, webbrowser, uvicorn

def _resource_dir():
    # PyInstaller 解包目录
    return getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))

# 确保能 import 到打包进去的 app/engine
sys.path.insert(0, _resource_dir())
os.chdir(_resource_dir())

def open_browser():
    time.sleep(1.5)
    webbrowser.open("http://127.0.0.1:8000")

if __name__ == "__main__":
    from app import app
    threading.Thread(target=open_browser, daemon=True).start()
    uvicorn.run(app, host="127.0.0.1", port=8000, log_level="warning")
