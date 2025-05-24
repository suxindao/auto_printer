import os
import sys
import time
import subprocess
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

# 获取监听目录（来自命令行参数）
if len(sys.argv) < 2:
    print("❌ 用法错误：请指定要监听的文件夹路径")
    print("✅ 示例：python auto_printer_mac.py \"/Users/yourname/Documents/ToPrint\"")
    sys.exit(1)

WATCH_FOLDER = sys.argv[1]

def print_pdf(file_path):
    print(f"🖨️ 打印文件: {file_path}")
    try:
        # 使用系统的 lp 命令进行打印
        subprocess.run(["lp", file_path], check=True)
        print("✅ 打印成功")
    except subprocess.CalledProcessError as e:
        print(f"❌ 打印失败: {e}")

class PDFHandler(FileSystemEventHandler):
    def on_created(self, event):
        if event.is_directory:
            return
        filepath = event.src_path
        if filepath.lower().endswith(".pdf"):
            time.sleep(15)
            print_pdf(filepath)

if __name__ == "__main__":
    if not os.path.exists(WATCH_FOLDER):
        print(f"❌ 目录不存在：{WATCH_FOLDER}")
        sys.exit(1)

    print(f"📂 正在监听目录：{WATCH_FOLDER}")
    print(f"🖨️ 使用系统默认打印机")

    event_handler = PDFHandler()
    observer = Observer()
    observer.schedule(event_handler, path=WATCH_FOLDER, recursive=False)
    observer.start()

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()
