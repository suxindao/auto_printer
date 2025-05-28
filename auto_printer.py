import os
import sys
import time
import win32api
import win32print
import win32com.client
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import pythoncom

# 获取监听目录（来自命令行参数）
if len(sys.argv) < 2:
    print("❌ 用法错误：请指定要监听的文件夹路径")
    print("✅ 示例：python auto_printer.py \"C:\\ToPrint\"")
    sys.exit(1)

WATCH_FOLDER = sys.argv[1]
PRINTER_NAME = win32print.GetDefaultPrinter()

# # 获取打印机支持的纸张数量
# try:
#     hprinter = win32print.OpenPrinter(PRINTER_NAME)
#     level = 1
#     forms = win32print.EnumForms(hprinter)
#     print(f"\n打印机 '{PRINTER_NAME}' 支持的纸张大小:")
#     for i, form in enumerate(forms, 1):
#         print(f"{i}. {form['Name']} (宽度: {form['Size']['cx']/1000:.1f}cm × 高度: {form['Size']['cy']/1000:.1f}cm)")
# except Exception as e:
#     print(f"获取纸张大小时出错: {e}")

def print_pdf(file_path):
    print(f"🖨️ 正在打印 PDF 文件: {file_path}")
    win32api.ShellExecute(
        0,
        "print",
        file_path,
        f'/d:"{PRINTER_NAME}"',
        ".",
        0
    )


def print_excel(file_path):
    print(f"📊 正在打印 Excel 文件: {file_path}")
    pythoncom.CoInitialize()  # 初始化当前线程的 COM 环境
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    workbook = None
    try:
        workbook = excel.Workbooks.Open(file_path)

        for sheet in workbook.Sheets:
            # 设置打印纸张为 A4（枚举值 9），其他常见值见下方
            sheet.PageSetup.PaperSize = 132  # A4
            # 设置为缩放：1 页宽，1 页高（即适应一页打印）
            sheet.PageSetup.Zoom = 75
            sheet.PageSetup.FitToPagesWide = 1
            sheet.PageSetup.FitToPagesTall = 1

        workbook.PrintOut()
        print("✅ Excel 打印成功")
    except Exception as e:
        print(f"❌ Excel 打印失败: {e}")
    finally:
        # 关闭工作簿（忽略可能的错误）
        try:
            if workbook:
                workbook.Close(False)
        except Exception as e:
            print(f"⚠️ 无法关闭工作簿: {e}")
        # 关闭 Excel 实例
        try:
            excel.Quit()
        except Exception as e:
            print(f"⚠️ 无法退出 Excel 应用程序: {e}")
        pythoncom.CoUninitialize()  # 清理 COM 环境

class AutoPrintHandler(FileSystemEventHandler):
    def on_created(self, event):
        if event.is_directory:
            return
        filepath = event.src_path.lower()

        # 跳过 Excel 临时文件
        filename = os.path.basename(filepath)
        if filename.startswith("~$"):
            # print(f"⚠️ 忽略临时文件: {filename}")
            return

        # 延迟确保文件写入完成
        time.sleep(8)

        if filepath.endswith(".pdf"):
            print_pdf(event.src_path)
        elif filepath.endswith(".xls") or filepath.endswith(".xlsx"):
            print_excel(event.src_path)


if __name__ == "__main__":
    if not os.path.exists(WATCH_FOLDER):
        print(f"❌ 目录不存在：{WATCH_FOLDER}")
        sys.exit(1)

    print(f"📂 正在监听目录：{WATCH_FOLDER}")
    print(f"🖨️ 默认打印机：{PRINTER_NAME}")

    event_handler = AutoPrintHandler()
    observer = Observer()
    observer.schedule(event_handler, path=WATCH_FOLDER, recursive=False)
    observer.start()

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()
