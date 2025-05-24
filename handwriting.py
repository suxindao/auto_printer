import cv2
import numpy as np
from PIL import Image, ImageDraw, ImageFont
import random


def generate_signature_image(text, font_path, output_path="signature.png"):
    # 1. 创建基本签名图像（白底 + 手写字体）
    img_width, img_height = 400, 150
    img = Image.new("RGB", (img_width, img_height), color=(255, 255, 255))
    draw = ImageDraw.Draw(img)

    font_size = 80
    font = ImageFont.truetype(font_path, font_size)
    # text_width, text_height = draw.textsize(text, font=font)
    bbox = draw.textbbox((0, 0), text, font=font)
    text_width = bbox[2] - bbox[0]
    text_height = bbox[3] - bbox[1]

    x = (img_width - text_width) / 2
    y = (img_height - text_height) / 2
    draw.text((x, y), text, font=font, fill=(0, 0, 139))  # 深蓝墨水效果

    img.save(output_path)  # 初始图像保存
    return output_path


def distort_image(image_path, output_path="signature_distorted.png"):
    # 2. 使用 OpenCV 加载图像
    img = cv2.imread(image_path, cv2.IMREAD_COLOR)
    h, w = img.shape[:2]
    # 3. 模拟波浪扭曲（模拟笔迹波动）
    def wave_distort(img, magnitude=5):
        distorted = np.zeros_like(img)
        for i in range(h):
            offset = int(magnitude * np.sin(2 * np.pi * i / 50))
            for j in range(w):
                if 0 <= j + offset < w:
                    distorted[i, j] = img[i, (j + offset) % w]
        return distorted

    # 4. 添加轻微旋转
    angle = random.uniform(-5, 5)
    M = cv2.getRotationMatrix2D((w // 2, h // 2), angle, 1)
    img_rotated = cv2.warpAffine(img, M, (w, h), borderValue=(255, 255, 255))

    # 5. 模糊模拟笔锋自然过渡
    img_blurred = cv2.GaussianBlur(img_rotated, (3, 3), 0)

    # 6. 扭曲图像
    img_distorted = wave_distort(img_blurred)

    # 7. 可选：模拟墨迹缺损（遮罩）
    # mask = np.random.randint(0, 2, (h, w), dtype=np.uint8) * 255
    # img_distorted[mask == 0] = 255

    # 8. 保存最终图像
    cv2.imwrite(output_path, img_distorted)
    print(f"伪造签名图像已保存：{output_path}")


# --- 使用示例 ---
if __name__ == "__main__":
    text = "老大来啦"
    font_path = "./font/HYYanYunW.ttf"  # 替换为你下载的中文手写字体路径
    clean_img = generate_signature_image(text, font_path)
    distort_image(clean_img)
