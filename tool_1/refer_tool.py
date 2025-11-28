import os
try:
    import cv2
    import numpy as np
except Exception:
    cv2 = None
    np = None
try:
    from PIL import Image, ImageDraw, ImageFont
except Exception:
    Image = None
    ImageDraw = None
    ImageFont = None
try:
    from langchain.tools import tool
except Exception:
    def tool(*args, **kwargs):
        def _wrap(f):
            return f
        return _wrap

def load_env_file():
    env_path = os.path.join(os.getcwd(), ".env")
    if not os.path.exists(env_path):
        project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        env_path = os.path.join(project_root, ".env")
    if os.path.exists(env_path):
        with open(env_path, 'r', encoding='utf-8') as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith('#') and '=' in line:
                    key, value = line.split('=', 1)
                    if value.startswith(("\"", "'")) and value.endswith(("\"", "'")):
                        value = value[1:-1]
                    os.environ[key.strip()] = value.strip()

load_env_file()

def imread_unicode(path):
    return cv2.imdecode(np.fromfile(path, dtype=np.uint8), cv2.IMREAD_COLOR) if cv2 is not None and np is not None else None

def imwrite_unicode(path, img):
    if cv2 is None:
        return False
    ext = os.path.splitext(path)[1]
    success, buf = cv2.imencode(ext, img)
    if success:
        buf.tofile(path)
    return bool(success)

def draw_chinese_text(image, text, position, font_size, color):
    pil_img = Image.fromarray(image)
    draw = ImageDraw.Draw(pil_img)
    font_path = os.path.join("font", "SimHei.ttf")
    font = ImageFont.truetype(font_path, font_size)
    try:
        bbox = draw.textbbox(position, text, font=font)
        text_width = bbox[2] - bbox[0]
        text_height = bbox[3] - bbox[1]
    except Exception:
        w, h = draw.textsize(text, font=font)
        text_width = w
        text_height = h
    x, y = position
    padding = 6
    bg_x1 = max(x - padding, 0)
    bg_y1 = max(y - padding, 0)
    bg_x2 = min(x + text_width + padding, image.shape[1])
    bg_y2 = min(y + text_height + padding, image.shape[0])
    draw.rectangle([bg_x1, bg_y1, bg_x2, bg_y2], fill=color)
    draw.text((x, y), text, font=font, fill=(255, 255, 255))
    return np.array(pil_img)

def extract_defect_name(filename):
    basename = os.path.splitext(os.path.basename(filename))[0]
    parts = basename.split("-")
    if len(parts) >= 2:
        full_desc = parts[-1]
        for keyword in ["大里程侧左侧", "大里程侧右侧", "小里程侧左侧", "小里程侧右侧"]:
            if full_desc.startswith(keyword):
                return full_desc.replace(keyword, "")
        return full_desc
    return basename

def _annotate(image_path, output_path, min_area=1200):
    if cv2 is None or np is None or Image is None:
        raise RuntimeError("依赖缺失: 请安装 opencv-python pillow numpy")
    img = imread_unicode(image_path)
    if img is None:
        raise FileNotFoundError(f"无法读取图片: {image_path}")
    defect_name = extract_defect_name(image_path)
    confidence = 0.95
    text_to_draw = f"{defect_name} {confidence:.2f}"
    hsv = cv2.cvtColor(img, cv2.COLOR_BGR2HSV)
    lower_red1 = np.array([0, 150, 180])
    upper_red1 = np.array([8, 255, 255])
    lower_red2 = np.array([172, 150, 180])
    upper_red2 = np.array([179, 255, 255])
    mask1 = cv2.inRange(hsv, lower_red1, upper_red1)
    mask2 = cv2.inRange(hsv, lower_red2, upper_red2)
    mask = mask1 | mask2
    kernel = np.ones((3, 3), np.uint8)
    mask = cv2.morphologyEx(mask, cv2.MORPH_OPEN, kernel, iterations=1)
    mask = cv2.morphologyEx(mask, cv2.MORPH_CLOSE, kernel, iterations=2)
    contours, _ = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    for cnt in contours:
        x, y, w, h = cv2.boundingRect(cnt)
        area = w * h
        if area < min_area:
            continue
        img = draw_chinese_text(img, text_to_draw, (x, y - 40 if y > 40 else y), 35, (255, 0, 0))
    ok = imwrite_unicode(output_path, img)
    if not ok:
        raise RuntimeError("写入失败")
    return output_path

@tool
def annotate_image_tool(image_path: str, output_path: str, min_area: int = 1200) -> str:
    """
    单图中文标注工具：在检测到的红色框附近叠加“缺陷名+置信度”中文标注，并保存到输出路径。

    参数:
        image_path: 输入图片路径（支持中文路径）。
        output_path: 输出图片保存路径（支持中文路径）。
        min_area: 红色矩形最小面积阈值，过滤噪点。

    返回:
        输出图片的路径字符串。
    """
    return _annotate(image_path, output_path, min_area)

@tool
def process_all_images(input_root: str, output_root: str, min_area: int = 1200) -> str:
    """
    批量中文标注工具：递归处理输入根目录下的所有图片，保持原目录结构到输出根目录。

    参数:
        input_root: 输入图片根目录（支持中文路径）。
        output_root: 输出图片根目录（支持中文路径）。
        min_area: 红色矩形最小面积阈值，过滤噪点。

    返回:
        处理统计信息字符串，例如“处理完成: 42 张”。
    """
    if not os.path.exists(input_root):
        raise FileNotFoundError(input_root)
    count = 0
    for root, dirs, files in os.walk(input_root):
        relative_path = os.path.relpath(root, input_root)
        output_folder = os.path.join(output_root, relative_path)
        os.makedirs(output_folder, exist_ok=True)
        for fname in files:
            lf = fname.lower()
            if lf.endswith((".jpg", ".png", ".jpeg")):
                in_path = os.path.join(root, fname)
                out_path = os.path.join(output_folder, fname)
                _annotate(in_path, out_path, min_area)
                count += 1
    return f"处理完成: {count} 张"

class ReferHandler:
    @staticmethod
    def annotate_image(image_path: str, output_path: str, min_area: int = 1200) -> str:
        return _annotate(image_path, output_path, min_area)
    @staticmethod
    def process_all_images(input_root: str, output_root: str, min_area: int = 1200) -> str:
        return process_all_images(input_root, output_root, min_area)

REFER_TOOLS = [annotate_image_tool, process_all_images]