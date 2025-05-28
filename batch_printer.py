import os
import sys
import time
import shutil
import win32api
import win32print
import pythoncom
import win32com.client
import logging
from datetime import datetime

# 设置打印机名称
DEFAULT_PRINTER = win32print.GetDefaultPrinter()
MONTHLY_PRINTER_NAME = "Microsoft XPS Document Writer"  # 请替换成你的月结单打印机名称


def is_monthly_file(filename):
    return "月结单" in filename


# === 打印 PDF 文件 ===
def print_pdf(file_path, use_alt_printer=False):
    logging.info(f"📊️ 打印 PDF: {file_path}")
    printer_name = MONTHLY_PRINTER_NAME if use_alt_printer else DEFAULT_PRINTER
    logging.info(f"🖨️ 打印机: {printer_name}")
    try:
        win32api.ShellExecute(
            0,
            "print",
            file_path,
            f'/d:"{printer_name}"',
            ".",
            0
        )
        logging.info("✅ PDF 打印成功")
    except Exception as e:
        logging.info(f"❌ PDF 打印失败: {e}")

    return True


# === 打印 Excel 文件 ===
def print_excel(file_path, use_alt_printer=False):
    logging.info(f"📊 打印 Excel: {file_path}")
    printer_name = MONTHLY_PRINTER_NAME if use_alt_printer else DEFAULT_PRINTER
    logging.info(f"🖨️ 打印机: {printer_name}")

    pythoncom.CoInitialize()
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    workbook = None
    try:
        workbook = excel.Workbooks.Open(file_path, ReadOnly=True)

        for sheet in workbook.Sheets:
            if use_alt_printer:
                # 设置打印纸张为 A4（枚举值 9），其他常见值见下方
                sheet.PageSetup.PaperSize = 9  # A4
                # 设置为缩放：1 页宽，1 页高（即适应一页打印）
                sheet.PageSetup.Zoom = False
                sheet.PageSetup.FitToPagesWide = 1
                sheet.PageSetup.FitToPagesTall = 1
            else:
                # 设置打印纸张为 A4（枚举值 9），其他常见值见下方
                try:
                    sheet.PageSetup.PaperSize = 132  # 132列纸
                except:
                    sheet.PageSetup.PaperSize = 9  # A4
                # 设置为缩放：75% 不缩放
                sheet.PageSetup.Zoom = 75
                sheet.PageSetup.FitToPagesWide = False
                sheet.PageSetup.FitToPagesTall = False

        workbook.PrintOut(ActivePrinter=printer_name)
        logging.info(f"✅ Excel 打印成功")
    except Exception as e:
        logging.info(f"❌ Excel 打印失败: {file_path}\n   原因: {e}")
    finally:
        try:
            if workbook:
                workbook.Close(False)
        except:
            pass
        try:
            excel.Quit()
        except:
            pass
        pythoncom.CoUninitialize()

    return True


# === 移动文件，保持目录结构 ===
def move_file_preserve_structure(src_file, src_root, dest_root):
    relative_path = os.path.relpath(src_file, src_root)
    dest_path = os.path.join(dest_root, relative_path)
    dest_dir = os.path.dirname(dest_path)
    os.makedirs(dest_dir, exist_ok=True)
    shutil.move(src_file, dest_path)
    logging.info(f"📁 文件已移动至: {dest_path}")


def delete_if_empty(dir_path):
    try:
        files = [f for f in os.listdir(dir_path) if not f.startswith("~$")]
        if not files:
            os.rmdir(dir_path)
            logging.info(f"🗑️ 删除空目录: {dir_path}")
            # 向上递归删除空目录
            parent = os.path.dirname(dir_path)
            if os.path.isdir(parent) and parent != dir_path:
                delete_if_empty(parent)
    except Exception as e:
        logging.info(f"⚠️ 删除目录失败 {dir_path}: {e}")


# === 主函数 ===
def main():
    if len(sys.argv) < 3:
        logging.info("❗ 用法: python batch_printer_recursive_move.py <源目录> <打印成功保存目录>")
        sys.exit(1)

    source_root = sys.argv[1]
    done_root = sys.argv[2]

    if not os.path.exists(source_root):
        logging.info(f"❌ 源目录不存在: {source_root}")
        sys.exit(1)

    os.makedirs(done_root, exist_ok=True)

    # 获取当前程序所在的目录（兼容 .py 和 .exe）
    base_dir = os.path.dirname(os.path.abspath(sys.argv[0]))

    # 构建日志目录路径
    log_dir = os.path.join(base_dir, "logs")
    os.makedirs(log_dir, exist_ok=True)

    # 初始化日志文件
    log_filename = datetime.now().strftime("log_%Y-%m-%d_%H-%M-%S.log")
    log_path = os.path.join(log_dir, log_filename)
    logging.basicConfig(
        filename=log_path,
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        encoding='utf-8'
    )

    # 将日志输出同时发送到控制台
    console = logging.StreamHandler()
    console.setLevel(logging.INFO)
    formatter = logging.Formatter('%(asctime)s - %(message)s')
    console.setFormatter(formatter)
    logging.getLogger().addHandler(console)

    logging.info(f"📂 开始递归打印目录: {source_root}")
    logging.info(f"🖨️ 默认打印机: {DEFAULT_PRINTER}")
    logging.info(f"🖨️ 月结单打印机: {MONTHLY_PRINTER_NAME}")

    for root, dirs, files in os.walk(source_root):
        for filename in files:
            if filename.startswith("~$"):
                continue  # 忽略 Excel 临时文件

            filepath = os.path.join(root, filename)
            is_monthly = is_monthly_file(filename)
            success = False

            if filename.lower().endswith(".pdf"):
                success = print_pdf(filepath, use_alt_printer=is_monthly)

            elif filename.lower().endswith((".xls", ".xlsx")):
                success = print_excel(filepath, use_alt_printer=is_monthly)

            # ✅ 每打印完一个文件，不论成功失败，暂停 5 秒
            time.sleep(5)

            if success:
                move_file_preserve_structure(filepath, source_root, done_root)
                delete_if_empty(root)

            logging.info(f"")


if __name__ == "__main__":
    main()
