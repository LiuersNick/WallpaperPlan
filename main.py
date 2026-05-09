import os
import sys
import win32com.client as win32
import fitz
from PIL import Image, ImageOps
import cv2
import numpy as np
import ctypes
from pathlib import Path
import logging
import yaml
from datetime import datetime
from logging.handlers import RotatingFileHandler
import inspect

def get_app_dir():
    if getattr(sys, 'frozen', False):
        return Path(sys.executable).parent
    else:
        return Path(inspect.getframeinfo(inspect.currentframe()).filename).parent

APP_DIR = get_app_dir()

# ================== Configuration ==================
CONFIG_PATH = APP_DIR / "config" / "config.yaml"

def load_config():
    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)

CONFIG = load_config()

EXCEL_FILE = CONFIG["excel"]["file"]
SHEET_NAME = CONFIG["excel"]["sheet"]
OUTPUT_DIR = APP_DIR / "temp"
TABLE_SCALE = CONFIG["image"]["table_scale"]
INVERT_COLORS = CONFIG["image"]["invert_colors"]
CROP_THRESHOLD = CONFIG["image"]["crop_threshold"]
DPI = CONFIG["image"]["dpi"]
LOG_LEVEL = CONFIG["logging"]["level"]
LOG_FILE_ENABLED = CONFIG["logging"]["file_enabled"]
# ================== Configuration ==================

PDF_PATH = OUTPUT_DIR / "temp_export.pdf"
RAW_PNG = OUTPUT_DIR / "raw_table.png"
CROPPED_PNG = OUTPUT_DIR / "cropped.png"
FINAL_WALLPAPER = OUTPUT_DIR / "final_wallpaper.png"

def setup_logging():
    logs_dir = APP_DIR / "logs"
    logs_dir.mkdir(exist_ok=True)

    handlers = [logging.StreamHandler(sys.stdout)]

    if LOG_FILE_ENABLED:
        log_file = logs_dir / f"wallpaper_{datetime.now().strftime('%Y%m%d')}.log"
        file_handler = RotatingFileHandler(
            log_file,
            maxBytes=5 * 1024 * 1024,
            backupCount=3,
            encoding="utf-8"
        )
        file_handler.setFormatter(logging.Formatter(
            '%(asctime)s [%(levelname)s] %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        ))
        handlers.append(file_handler)

    logging.basicConfig(
        level=getattr(logging, LOG_LEVEL),
        format='%(asctime)s [%(levelname)s] %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S',
        handlers=handlers
    )
    return logging.getLogger()

logger = setup_logging()

def get_screen_resolution():
    user32 = ctypes.windll.user32
    user32.SetProcessDPIAware()
    return user32.GetSystemMetrics(0), user32.GetSystemMetrics(1)

def export_excel_to_pdf(excel_file, sheet_name, pdf_path):
    excel = None
    try:
        excel = win32.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        wb = excel.Workbooks.Open(os.path.abspath(excel_file))
        ws = wb.Sheets(sheet_name)

        ps = ws.PageSetup
        ps.PaperSize = 9          
        ps.Orientation = 2        
        ps.LeftMargin = ps.RightMargin = ps.TopMargin = ps.BottomMargin = 0.25  
        ps.FitToPagesWide = 1
        ps.FitToPagesTall = 1

        ws.ExportAsFixedFormat(
            Type=0, Filename=str(pdf_path),
            Quality=0, IncludeDocProperties=True,
            IgnorePrintAreas=False, OpenAfterPublish=False
        )
        wb.Close(SaveChanges=False)
        logger.info("Excel 已导出为 PDF: %s", pdf_path)
    except Exception as e:
        logger.error("PDF 导出失败: %s", e)
        raise RuntimeError(f"导出 PDF 失败: {e}")
    finally:
        if excel:
            excel.Quit()

def pdf_to_png(pdf_path, png_path, dpi=200):
    doc = fitz.open(pdf_path)
    page = doc[0]
    mat = fitz.Matrix(dpi / 72, dpi / 72)
    pix = page.get_pixmap(matrix=mat, alpha=False)
    pix.save(png_path)
    doc.close()
    logger.info("PDF 已转换为 PNG: %s", png_path)

def auto_crop_white_borders(image_path, output_path, threshold=240):
    img = cv2.imread(str(image_path))
    if img is None:
        logger.error("无法读取图像: %s", image_path)
        raise FileNotFoundError(f"无法读取图像: {image_path}")

    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    _, binary = cv2.threshold(gray, threshold, 255, cv2.THRESH_BINARY_INV)

    coords = cv2.findNonZero(binary)
    if coords is None:
        logger.warning("未检测到非白色区域，跳过裁剪")
        cv2.imwrite(str(output_path), img)
        return

    x, y, w, h = cv2.boundingRect(coords)
    cropped = img[y:y+h, x:x+w]
    cv2.imwrite(str(output_path), cropped)
    logger.info("自动裁剪白边完成: %s", output_path)

def invert_image_if_needed(image_path, output_path, invert=True):
    if not invert:
        img = Image.open(image_path).convert("RGB")
        img.save(output_path)
        return

    try:
        img = Image.open(image_path).convert("RGB")
        inverted = ImageOps.invert(img)
        inverted.save(output_path)
        logger.info("图像已反转（白色背景黑色文字）: %s", output_path)
    except Exception as e:
        logger.warning("图像反转失败（可能包含透明度）: %s", e)
        img.save(output_path)

def embed_into_desktop_bg(table_img_path, final_path, scale=0.85):
    screen_w, screen_h = get_screen_resolution()
    table_img = Image.open(table_img_path).convert("RGB")

    target_w = int(screen_w * scale)
    target_h = int(screen_h * scale)

    table_w, table_h = table_img.size
    scale_factor = min(target_w / table_w, target_h / table_h)
    new_w = int(table_w * scale_factor)
    new_h = int(table_h * scale_factor)
    resized_table = table_img.resize((new_w, new_h), Image.LANCZOS)

    bg = Image.new("RGB", (screen_w, screen_h), (0, 0, 0))
    x = (screen_w - new_w) // 2
    y = (screen_h - new_h) // 2
    bg.paste(resized_table, (x, y))

    bg.save(final_path)
    logger.info("壁纸已生成: %dx%d, 表格缩放比例: %.0f%%", screen_w, screen_h, scale*100)

def set_wallpaper(image_path):
    if not os.path.exists(image_path):
        logger.error("壁纸文件不存在: %s", image_path)
        raise FileNotFoundError(f"壁纸文件不存在: {image_path}")
    ctypes.windll.user32.SystemParametersInfoW(20, 0, image_path, 3)
    logger.info("桌面壁纸已更新")

def main():
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    try:
        export_excel_to_pdf(EXCEL_FILE, SHEET_NAME, PDF_PATH)
        pdf_to_png(PDF_PATH, RAW_PNG, dpi=DPI)
        auto_crop_white_borders(RAW_PNG, CROPPED_PNG, threshold=CROP_THRESHOLD)

        temp_after_invert = OUTPUT_DIR / "after_invert.png"
        invert_image_if_needed(CROPPED_PNG, temp_after_invert, invert=INVERT_COLORS)

        embed_into_desktop_bg(temp_after_invert, FINAL_WALLPAPER, scale=TABLE_SCALE)
        set_wallpaper(str(FINAL_WALLPAPER))

        logger.info("执行完成，壁纸已设置为: %s", FINAL_WALLPAPER)
    except Exception as e:
        logger.error("执行过程中发生错误: %s", e)
        import traceback
        traceback.print_exc()
        logger.info("按回车键退出...")
        input()
    finally:
        import shutil
        if OUTPUT_DIR.exists():
            shutil.rmtree(OUTPUT_DIR)
            logger.info("临时文件已清理")

if __name__ == "__main__":
    main()