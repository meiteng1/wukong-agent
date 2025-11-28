from langchain_openai import ChatOpenAI
from openai import OpenAI
from dotenv import load_dotenv
import os
import chardet
from docx import Document

# 加载环境变量
load_dotenv()

# 聊天模型主类
class MyChatModel:
    def __init__(self):
        # 基础配置
        self.model_name = os.getenv("MODEL_NAME")
        
        # 豆包API配置
        self.api_key = os.getenv("ARK_API_KEY")
        self.base_url = os.getenv("ARK_API_BASE")
        self.raw_report_path = os.getenv("RAW_REPORT_PATH")
        self.template_report_path = os.getenv("TEMPLATE_REPORT_PATH")
        
        
        # 懒加载实例
        self._llm = None  # langchain的ChatOpenAI实例
        self._openai_client = None  # 原生OpenAI客户端实例
        self._prompt_system = self._load_system_prompt()

    def _load_system_prompt(self):
        """加载桥梁检测报告生成的系统提示词"""
        return """
        你是一名专业的桥梁工程质检与检测报告撰写工程师，负责根据程序自动统计所得的文本内容，生成符合《桥梁支座检查报告模板》的正式、规范、专业的桥梁缺陷支座检查分析报告...
        """.strip()  # 实际使用时替换为完整prompt

    @property
    def openai_client(self):
        """懒加载原生OpenAI客户端"""
        if not self._openai_client:
            if not self.api_key:
                raise ValueError("未找到有效的API密钥，请设置ARK_API_KEY")
            self._openai_client = OpenAI(
                base_url=self.base_url,
                api_key=self.api_key
            )
        return self._openai_client

    def get_langchain_llm(self):
        """获取langchain的ChatOpenAI实例（配置豆包API）"""
        if not self._llm:
            # 验证配置是否存在
            if not self.api_key:
                raise ValueError("未找到有效的API密钥，请设置ARK_API_KEY环境变量")
            if not self.base_url:
                raise ValueError("未找到有效的API基座地址，请设置ARK_API_BASE环境变量")
            # 实例化ChatOpenAI并传递豆包API配置
            self._llm = ChatOpenAI(
                model_name=self.model_name,
                api_key=self.api_key,
                base_url=self.base_url,
                # 可选配置：根据需求调整
                temperature=0.2,  # 控制生成的随机性（0-1，越小越严谨）
                # max_tokens=   # 最大生成 tokens 数
            )
        return self._llm

    def generate_bridge_report(self, stream_callback=None):
        """生成桥梁检测报告（流式处理）"""
        # 通过文档处理器读取输入数据
        raw_report = self.doc_handler.read_text_auto(self.raw_report_path)
        template = self.doc_handler.read_text_auto(self.template_report_path)
        
        # 构建消息
        messages = [
            {"role": "system", "content": self._prompt_system},
            {"role": "user", "content": (
                "请根据以下数据，生成完整的桥梁支座检测报告：\n\n"
                f"【统计报告】\n{raw_report}\n\n"
                f"【报告模板】\n{template}"
            )}
        ]
        
        # 流式调用
        stream = self.openai_client.chat.completions.create(
            model=self.model_name,
            messages=messages,
            stream=True,
            reasoning_effort="high"
        )
        
        content = ""
        reasoning_content = ""
        for chunk in stream:
            delta = chunk.choices[0].delta
            if getattr(delta, "reasoning_content", None):
                reasoning_content += delta.reasoning_content
                if stream_callback:
                    stream_callback("reasoning", delta.reasoning_content)
            if delta.content:
                content += delta.content
                if stream_callback:
                    stream_callback("content", delta.content)
        
        return {
            "full_content": content,
            "reasoning": reasoning_content
        }

if __name__ == '__main__':
    model = MyChatModel()
    # 测试报告生成
    try:
        result = model.generate_bridge_report(
            stream_callback=lambda t, c: print(c, end="")
        )
        # 使用文档处理器保存结果
        model.doc_handler.save_to_docx(result["full_content"], "桥梁检测报告_测试.docx")
    except Exception as e:
        print(f"错误: {str(e)}")