import os
import sys
import time
import win32api
import win32print
import pythoncom
import win32com.client

# === 打印 PDF 文件 ===
def print_pdf(file_path):
    print(f"🖨️ 打印 PDF: {file_path}")
    try:
        win32api.ShellExecute(
            0,
            "print",
            file_path,
            f'/d:"{win32print.GetDefaultPrinter()}"',
            ".",
            0
        )
        time.sleep(5)
    except Exception as e:
        print(f"❌ PDF 打印失败: {e}")

# === 打印 Excel 文件 ===
def print_excel(file_path):
    print(f"📊 打印 Excel: {file_path}")
    pythoncom.CoInitialize()
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    workbook = None
    try:
        workbook = excel.Workbooks.Open(file_path, ReadOnly=True)
        workbook.PrintOut()
        print("✅ Excel 打印成功")
    except Exception as e:
        print(f"❌ Excel 打印失败: {e}")
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

# === 主函数 ===
def main():
    if len(sys.argv) < 2:
        print("❗ 用法: python batch_printer.py <要打印的目录路径>")
        sys.exit(1)

    folder_path = sys.argv[1]

    if not os.path.exists(folder_path):
        print(f"❌ 目录不存在: {folder_path}")
        sys.exit(1)

    print(f"📂 开始打印目录: {folder_path}")
    for filename in os.listdir(folder_path):
        if filename.startswith("~$"):  # 忽略 Excel 临时文件
            continue

        filepath = os.path.join(folder_path, filename)
        if filename.lower().endswith(".pdf"):
            print_pdf(filepath)
        elif filename.lower().endswith((".xls", ".xlsx")):
            print_excel(filepath)

if __name__ == "__main__":
    main()
