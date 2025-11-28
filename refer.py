import os
import cv2
import numpy as np
from PIL import Image, ImageDraw, ImageFont


# ---------------- 中文路径读写函数 ----------------
def imread_unicode(path):
    return cv2.imdecode(np.fromfile(path, dtype=np.uint8), cv2.IMREAD_COLOR)


def imwrite_unicode(path, img):
    ext = os.path.splitext(path)[1]
    success, buf = cv2.imencode(ext, img)
    if success:
        buf.tofile(path)


# ---------------- 绘制中文文本函数 ----------------
def draw_chinese_text(image, text, position, font_size, color):
    pil_img = Image.fromarray(image)
    draw = ImageDraw.Draw(pil_img)
    font_path = os.path.join("font", "SimHei.ttf")  # 确保有 SimHei.ttf
    font = ImageFont.truetype(font_path, font_size)

    try:
        bbox = draw.textbbox(position, text, font=font)
        text_width = bbox[2] - bbox[0]
        text_height = bbox[3] - bbox[1]
    except AttributeError:
        text_width, text_height = draw.textsize(text, font=font)

    x, y = position
    padding = 6
    bg_x1 = max(x - padding, 0)
    bg_y1 = max(y - padding, 0)
    bg_x2 = min(x + text_width + padding, image.shape[1])
    bg_y2 = min(y + text_height + padding, image.shape[0])

    # 绘制背景色（红色背景）
    draw.rectangle([bg_x1, bg_y1, bg_x2, bg_y2], fill=color)

    # 绘制白色文字
    draw.text((x, y), text, font=font, fill=(255, 255, 255))

    return np.array(pil_img)


# ---------------- 提取缺陷名 ----------------
def extract_defect_name(filename):
    """
    从文件名提取缺陷名
    格式示例: CS001-大里程侧右侧施工垃圾残留.jpg
    结果: 施工垃圾残留
    """
    basename = os.path.splitext(os.path.basename(filename))[0]
    parts = basename.split("-")
    if len(parts) >= 2:
        full_desc = parts[-1]  # e.g. 大里程侧右侧施工垃圾残留
        # 去掉“里程侧左侧右侧”等字样，只取最后部分
        for keyword in ["大里程侧左侧", "大里程侧右侧", "小里程侧左侧", "小里程侧右侧"]:
            if full_desc.startswith(keyword):
                return full_desc.replace(keyword, "")
        return full_desc
    return basename


# ---------------- 主函数 ----------------
def annotate_image(image_path, output_path, min_area=1200):
    img = imread_unicode(image_path)
    if img is None:
        print(f"❌ 无法读取图片: {image_path}")
        return

    defect_name = extract_defect_name(image_path)
    confidence = 0.95  # 可以改成实际检测的置信度
    text_to_draw = f"{defect_name} {confidence:.2f}"

    # 转 HSV 找纯红框
    hsv = cv2.cvtColor(img, cv2.COLOR_BGR2HSV)

    # 更严格的红色阈值（高饱和度、高亮度）
    lower_red1 = np.array([0, 150, 180])  # 纯正鲜红
    upper_red1 = np.array([8, 255, 255])
    lower_red2 = np.array([172, 150, 180])
    upper_red2 = np.array([179, 255, 255])

    mask1 = cv2.inRange(hsv, lower_red1, upper_red1)
    mask2 = cv2.inRange(hsv, lower_red2, upper_red2)
    mask = mask1 | mask2

    # 形态学操作，去噪点，闭合缺口
    kernel = np.ones((3, 3), np.uint8)
    mask = cv2.morphologyEx(mask, cv2.MORPH_OPEN, kernel, iterations=1)
    mask = cv2.morphologyEx(mask, cv2.MORPH_CLOSE, kernel, iterations=2)

    contours, _ = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

    for cnt in contours:
        x, y, w, h = cv2.boundingRect(cnt)
        area = w * h
        if area < min_area:  # 忽略太小的红色矩形
            continue

        # 绘制缺陷名（矩形上方）
        img = draw_chinese_text(img, text_to_draw, (x, y - 40 if y > 40 else y), 35, (255, 0, 0))

    imwrite_unicode(output_path, img)
    print(f"✅ 已处理并保存: {output_path}")


# ---------------- 批量处理（支持多级文件夹并保持原结构） ----------------
def process_all_images(input_root, output_root):
    for root, dirs, files in os.walk(input_root):
        # 计算当前root相对于输入根目录的相对路径
        relative_path = os.path.relpath(root, input_root)
        # 构造对应的输出文件夹路径
        output_folder = os.path.join(output_root, relative_path)
        os.makedirs(output_folder, exist_ok=True)

        for fname in files:
            if fname.lower().endswith((".jpg", ".png", ".jpeg")):
                in_path = os.path.join(root, fname)
                out_path = os.path.join(output_folder, fname)
                annotate_image(in_path, out_path)


if __name__ == "__main__":
    input_dir = r"D:\重复照片"
    output_dir = r"D:\ai处理重复照片"

    # 调用递归处理函数
    process_all_images(input_dir, output_dir)