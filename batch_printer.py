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
import configparser
import ctypes  # 顶部添加此模块

# 省略 imports，与你一致

# 全局变量
DEFAULT_PRINTER = win32print.GetDefaultPrinter()
MONTHLY_PRINTER_NAME = ""
DEFAULT_PAPER_SIZE = 9
DEFAULT_PAPER_ZOOM = 75
DELAY_SECONDS = 5
ENABLE_WAIT_PROMPT = True
WAIT_PROMPT_SLEEP = 30

def is_monthly_file(filename):
    return "月结单" in filename


def setup_logging(log_dir):
    os.makedirs(log_dir, exist_ok=True)
    log_filename = datetime.now().strftime("log_%Y-%m-%d_%H-%M-%S.log")
    log_path = os.path.join(log_dir, log_filename)

    logging.basicConfig(
        filename=log_path,
        level=logging.INFO,
        format="%(asctime)s - %(levelname)s - %(message)s",
        encoding="utf-8"
    )

    console = logging.StreamHandler()
    console.setLevel(logging.INFO)
    formatter = logging.Formatter('%(asctime)s - %(message)s')
    console.setFormatter(formatter)
    logging.getLogger().addHandler(console)


def read_config(config_path):
    config = configparser.ConfigParser()
    config.read(config_path, encoding="utf-8")

    global MONTHLY_PRINTER_NAME, DEFAULT_PAPER_SIZE, DEFAULT_PAPER_ZOOM, DELAY_SECONDS, ENABLE_WAIT_PROMPT, WAIT_PROMPT_SLEEP

    source = config.get("settings", "source_dir")
    target = config.get("settings", "target_dir")
    MONTHLY_PRINTER_NAME = config.get("settings", "monthly_printer_name")
    DEFAULT_PAPER_SIZE = int(config.get("settings", "default_paper_size"))
    DEFAULT_PAPER_ZOOM = int(config.get("settings", "default_paper_zoom"))
    DELAY_SECONDS = float(config.get("settings", "delay_seconds"))
    ENABLE_WAIT_PROMPT = config.getboolean("settings", "enable_wait_prompt", fallback=True)
    WAIT_PROMPT_SLEEP = float(config.get("settings", "wait_prompt_sleep"))

    logging.info(f"-------------------------")
    logging.info(f"⚙️ 配置文件信息:")
    logging.info(f"📂 源目录: {source}")
    logging.info(f"📂 保存目录: {target}")
    logging.info(f"🖨️ 月结单使用的打印机名称️: {MONTHLY_PRINTER_NAME}")
    logging.info(f"📄 针式打印机纸张编号: {DEFAULT_PAPER_SIZE}")
    logging.info(f"📄 针式打印机打印缩放比例: {DEFAULT_PAPER_ZOOM}")
    logging.info(f"📄 打印间隔: {DELAY_SECONDS}")
    logging.info(f"🔔 打印完目录是否弹窗并等待: {ENABLE_WAIT_PROMPT}")
    logging.info(f"-------------------------")

    return source, target


def print_pdf(path, use_alt=False):
    printer = MONTHLY_PRINTER_NAME if use_alt else DEFAULT_PRINTER

    logging.info(f"📄 打印 PDF: {path}")
    logging.info(f"🖨️ 打印机: {printer}")

    try:
        win32api.ShellExecute(0, "print", path, f'/d:"{printer}"', ".", 0)
        logging.info(f"✅ 打印成功 (PDF)")
        return True
    except Exception as e:
        logging.error(f"❌ 打印失败 (PDF): {e}")
        return False


def print_excel(path, use_alt=False):
    printer = MONTHLY_PRINTER_NAME if use_alt else DEFAULT_PRINTER

    logging.info(f"📊 打印 Excel: {path}")
    logging.info(f"🖨️ 打印机: {printer}")

    pythoncom.CoInitialize()
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    try:
        wb = excel.Workbooks.Open(path, ReadOnly=True)
        for sheet in wb.Sheets:
            if use_alt:
                sheet.PageSetup.PaperSize = 9  # A4
                sheet.PageSetup.Zoom = False
                sheet.PageSetup.FitToPagesWide = 1
                sheet.PageSetup.FitToPagesTall = 1
            else:
                try:
                    sheet.PageSetup.PaperSize = DEFAULT_PAPER_SIZE  # 132列纸
                except:
                    sheet.PageSetup.PaperSize = 9  # A4
                sheet.PageSetup.Zoom = DEFAULT_PAPER_ZOOM
                sheet.PageSetup.FitToPagesWide = False
                sheet.PageSetup.FitToPagesTall = False

        wb.PrintOut(ActivePrinter=printer)
        logging.info(f"✅ 打印成功 (Excel)")
        return True
    except Exception as e:
        logging.error(f"❌ 打印失败 (Excel): {e}")
        return False
    finally:
        try:
            wb.Close(False)
        except:
            pass
        try:
            excel.Quit()
        except:
            pass
        pythoncom.CoUninitialize()


def move_and_cleanup(src_file, src_root, target_root):
    rel_path = os.path.relpath(src_file, src_root)
    dest_file = os.path.join(target_root, rel_path)
    os.makedirs(os.path.dirname(dest_file), exist_ok=True)
    shutil.move(src_file, dest_file)
    logging.info(f"📁 已移动文件: {dest_file}")

    # 删除空目录
    src_dir = os.path.dirname(src_file)
    if not any(f for f in os.listdir(src_dir) if not f.startswith("~$")):
        try:
            os.rmdir(src_dir)
            logging.info(f"🗑️ 删除空目录: {src_dir}")
        except Exception as e:
            logging.warning(f"⚠️ 删除目录失败: {src_dir} - {e}")

    # 打印一个空行
    logging.warning(f"")


def main():
    base_dir = os.path.dirname(os.path.abspath(sys.argv[0]))
    config_path = os.path.join(base_dir, "config.ini")
    log_dir = os.path.join(base_dir, "logs")

    if not os.path.exists(config_path):
        print(f"❌ 配置文件不存在: {config_path}")
        return

    setup_logging(log_dir)

    source_root, target_root = read_config(config_path)
    logging.info(f"📂 监听目录: {source_root}")
    logging.info(f"📁 目标目录: {target_root}")

    for root, _, files in os.walk(source_root):

        any_printed = False

        for name in files:
            if name.startswith("~$"):
                continue
            full_path = os.path.join(root, name)
            is_monthly = is_monthly_file(name)
            success = False

            if name.lower().endswith(".pdf"):
                success = print_pdf(full_path, use_alt=is_monthly)
            elif name.lower().endswith((".xls", ".xlsx")):
                success = print_excel(full_path, use_alt=is_monthly)

            time.sleep(DELAY_SECONDS)

            if success:
                move_and_cleanup(full_path, source_root, target_root)
                any_printed = True
            else:
                sys.exit(1)

        if any_printed:
            msg = f"📁 当前目录打印完成: \n{root}\n\n📢 将在 {WAIT_PROMPT_SLEEP} 秒后继续打印下一个目录..."
            logging.info(msg)

            # 0x04 = MB_YESNO + MB_ICONQUESTION
            response = ctypes.windll.user32.MessageBoxW(
                0,
                msg,
                "📢 打印完成提示",
                0x04 | 0x20  # MB_YESNO | MB_ICONQUESTION
            )

            if response == 6:  # IDYES
                logging.info(f"✅ 用户选择等待，等待 {WAIT_PROMPT_SLEEP} 秒...")
                time.sleep(WAIT_PROMPT_SLEEP)
            else:
                logging.info("⏩ 用户选择跳过等待")

    logging.info("✅ 所有文件打印完成")


if __name__ == "__main__":
    main()
