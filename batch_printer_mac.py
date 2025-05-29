import os
import sys
import time
import shutil
import logging
import subprocess
from datetime import datetime
import configparser

# 参数默认值
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
    logging.getLogger().addHandler(console)


def read_config(config_path):
    config = configparser.ConfigParser()
    config.read(config_path, encoding="utf-8")

    global DEFAULT_PAPER_ZOOM, DELAY_SECONDS, ENABLE_WAIT_PROMPT, WAIT_PROMPT_SLEEP

    source = config.get("settings", "source_dir")
    target = config.get("settings", "target_dir")
    DEFAULT_PAPER_ZOOM = int(config.get("settings", "default_paper_zoom"))
    DELAY_SECONDS = float(config.get("settings", "delay_seconds"))
    ENABLE_WAIT_PROMPT = config.getboolean("settings", "enable_wait_prompt", fallback=True)
    WAIT_PROMPT_SLEEP = float(config.get("settings", "wait_prompt_sleep"))

    logging.info(f"📂 源目录: {source}")
    logging.info(f"📂 保存目录: {target}")
    return source, target


def print_file_mac(filepath):
    try:
        subprocess.run(["lp", filepath], check=True)
        logging.info(f"✅ 打印成功: {filepath}")
        return True
    except Exception as e:
        logging.error(f"❌ 打印失败: {filepath} - {e}")
        return False


def move_and_cleanup(src_file, src_root, target_root):
    rel_path = os.path.relpath(src_file, src_root)
    dest_file = os.path.join(target_root, rel_path)
    os.makedirs(os.path.dirname(dest_file), exist_ok=True)
    shutil.move(src_file, dest_file)
    logging.info(f"📁 已移动文件: {dest_file}")


def main():
    base_dir = os.path.dirname(os.path.abspath(sys.argv[0]))
    config_path = os.path.join(base_dir, "config.ini")
    log_dir = os.path.join(base_dir, "logs")

    if not os.path.exists(config_path):
        print(f"❌ 配置文件不存在: {config_path}")
        return

    setup_logging(log_dir)
    source_root, target_root = read_config(config_path)

    for root, _, files in os.walk(source_root):
        for name in files:
            if name.startswith("~$") or is_monthly_file(name):
                continue

            full_path = os.path.join(root, name)
            success = print_file_mac(full_path)

            time.sleep(DELAY_SECONDS)

            if success:
                move_and_cleanup(full_path, source_root, target_root)
            else:
                sys.exit(1)

    logging.info("✅ 所有文件打印完成")

    try:
        shutil.rmtree(source_root)
        logging.info(f"🧹 已删除源目录: {source_root}")
    except Exception as e:
        logging.warning(f"⚠️ 无法删除源目录: {source_root} - {e}")


if __name__ == "__main__":
    main()
