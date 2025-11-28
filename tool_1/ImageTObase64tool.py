import base64
import os
from typing import Optional, Dict
from PIL import Image
from io import BytesIO

# 加载环境变量（如果使用.env文件）
if os.path.exists('.env'):
    try:
        with open('.env', 'r', encoding='utf-8') as f:
            for line in f:
                if '=' in line and not line.strip().startswith('#'):
                    key, value = line.strip().split('=', 1)
                    os.environ[key.strip()] = value.strip().strip('"').strip('\'')
    except Exception as e:
        print(f"警告：加载.env文件失败：{str(e)}")


class ImageToBase64Tool:
    """
    图片转Base64工具类
    """

    def __init__(
        self,
        compress: bool = True,
        max_width: int = 1280,
        quality: int = 85
    ):
        self.compress = compress
        self.max_width = max_width
        self.quality = quality

        self.supported_formats = {
            "jpg": "image/jpeg",
            "jpeg": "image/jpeg",
            "png": "image/png",
            "bmp": "image/bmp",
            "gif": "image/gif"
        }

    # ---------------------
    # 获取 MIME 类型
    # ---------------------
    def _get_image_mime_type(self, image_path: str) -> str:
        print(f"调试信息 - 处理路径: {image_path}")

        file_ext = os.path.splitext(image_path)[-1].lower().lstrip(".")
        print(f"调试信息 - splitext获取的扩展名: '{file_ext}'")

        filename = os.path.basename(image_path).lower()
        detected_ext = None

        for ext in sorted(self.supported_formats.keys(), key=len, reverse=True):
            if filename.endswith("." + ext):
                detected_ext = ext
                print(f"调试信息 - 从文件名检测到扩展名: '{detected_ext}'")
                break

        final_ext = detected_ext if detected_ext else file_ext

        print(f"调试信息 - 最终确定的扩展名: '{final_ext}'")

        if not final_ext:
            raise ValueError(f"unknown file extension: {image_path}")

        if final_ext not in self.supported_formats:
            raise ValueError(
                f"不支持的图片格式：{final_ext}，仅支持{list(self.supported_formats.keys())}"
            )

        return self.supported_formats[final_ext]

    # ---------------------
    # 压缩图片（关键修复点）
    # ---------------------
    def _compress_image(self, image_bytes: bytes, ext: str) -> bytes:
        """
        ext: 由文件扩展名决定的格式（jpg / jpeg / png...）
        """
        with Image.open(BytesIO(image_bytes)) as img:

            # 修复：从 BytesIO 打开时 img.format 可能为 None，因此不再依赖它
            if img.mode not in ("RGB", "RGBA"):
                img = img.convert("RGB")

            width, height = img.size

            # 缩放
            if width > self.max_width:
                scale_ratio = self.max_width / width
                new_height = int(height * scale_ratio)
                img = img.resize((self.max_width, new_height), Image.Resampling.LANCZOS)

            output = BytesIO()


            if ext in ["jpg", "jpeg"]:
                save_format = "JPEG"
            elif ext == "png":
                save_format = "PNG"
            else:
                save_format = "JPEG"  # 默认降级为 JPEG，防止未知格式崩溃

            img.save(
                output,
                format=save_format,
                quality=self.quality,
                optimize=True
            )

            return output.getvalue()

    # ---------------------
    # 转 Base64
    # ---------------------
    def image_to_base64(self, image_path: str) -> str:

        if not os.path.exists(image_path):
            raise FileNotFoundError(f"图片文件不存在：{image_path}")

        try:
            with open(image_path, "rb") as f:
                image_bytes = f.read()
        except Exception as e:
            raise IOError(f"读取图片失败：{str(e)}")

        # 获取扩展名用于压缩功能
        ext = os.path.splitext(image_path)[-1].lower().lstrip(".")

        if self.compress:
            image_bytes = self._compress_image(image_bytes, ext)

        mime_type = self._get_image_mime_type(image_path)
        base64_str = base64.b64encode(image_bytes).decode("utf-8")

        return f"{mime_type};base64,{base64_str}"

    # ---------------------
    # API 使用格式
    # ---------------------
    def get_api_image_param(self, image_path: str) -> Dict[str, Dict[str, str]]:
        base64_str = self.image_to_base64(image_path)
        return {
            "type": "image_url",
            "image_url": {
                "url": base64_str
            }
        }


# ----------------------------------------------------------
# 示例执行
# ----------------------------------------------------------
if __name__ == "__main__":
    img_tool = ImageToBase64Tool()

    local_image_path = os.getenv("LOCAL_IMAGE_PATH")

    if not local_image_path:
        print("错误：环境变量LOCAL_IMAGE_PATH未设置")
        exit(1)

    local_image_path = os.path.normpath(local_image_path)
    print(f"当前使用的图片路径：{local_image_path}")

    try:
        api_image_param = img_tool.get_api_image_param(local_image_path)
        print("图片转换成功！")
        print("Base64 图片 URL：")
        print(api_image_param["image_url"]["url"])


    except Exception as e:
        print(f"图片处理失败：{str(e)}")
        exit(1)
