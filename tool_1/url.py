import os
import requests
from typing import Dict

# 尝试导入dotenv库来加载.env文件
try:
    from dotenv import load_dotenv
    # 加载.env文件中的环境变量
    load_dotenv()
except ImportError:
    print("警告: dotenv库未安装，将尝试直接读取环境变量")

class ArkFileUploader:
    """
    火山方舟文件上传器
    上传任意文件到方舟文件仓库，返回在线可访问 URL：
    https://ark-file.volces.com/xxx
    """

    def __init__(self):
        # 从环境变量获取ARK_API_KEY
        self.api_key = os.getenv("ARK_API_KEY")
        if not self.api_key:
            raise ValueError("缺少 ARK_API_KEY 环境变量，请在.env文件中设置")

        # 火山方舟文件上传接口（固定）
        self.upload_url = "https://ark.cn-beijing.volces.com/api/v3/files"

        self.headers = {
            "Authorization": f"Bearer {self.api_key}"
        }

    def upload(self, file_path: str) -> str:
        """
        上传本地文件并返回公网可访问 URL
        """

        if not os.path.exists(file_path):
            raise FileNotFoundError(f"文件不存在：{file_path}")

        with open(file_path, "rb") as f:
            files = {"file": f}
            response = requests.post(self.upload_url, headers=self.headers, files=files)

        if response.status_code != 200:
            raise RuntimeError(f"上传失败：{response.text}")

        resp = response.json()

        # 返回 URL（多模态 API 直接可用）
        return resp["data"]["url"]
if __name__ == "__main__":
    uploader = ArkFileUploader()

    # 从环境变量获取图片路径
    file_path = os.getenv("LOCAL_IMAGE_PATH")
    if not file_path:
        print("错误: 环境变量LOCAL_IMAGE_PATH未设置")
        print("请在.env文件中设置LOCAL_IMAGE_PATH变量，指定完整的图片路径")
        exit(1)
    
    # 规范化路径，处理Windows路径分隔符问题
    file_path = os.path.normpath(file_path)
    print("上传中:", file_path)

    url = uploader.upload(file_path)

    print("\n上传成功！图片 URL：")
    print(url)
