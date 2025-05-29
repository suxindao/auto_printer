import sys
import win32print


def main():
    PRINTER_NAME = win32print.GetDefaultPrinter()
    try:
        hprinter = win32print.OpenPrinter(PRINTER_NAME)
        printer_info = win32print.GetPrinter(hprinter, 2)
        # 获取打印机支持的纸张数量
        level = 1
        forms = win32print.EnumForms(hprinter)
        print(f"\n打印机 '{PRINTER_NAME}' 支持的纸张大小:")
        for i, form in enumerate(forms, 1):
            print(
                f"{i}. {form['Name']} (宽度: {form['Size']['cx'] / 1000:.1f}cm × 高度: {form['Size']['cy'] / 1000:.1f}cm)")
    except Exception as e:
        print(f"设置纸张大小时出错: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
