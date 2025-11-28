import time
import sys
import os
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from langchain.agents import create_tool_calling_agent, AgentExecutor
from Tool.word_Imagetool import insert_images_to_docx
from Tool.documentRead_tool import read_text_auto, save_to_docx
from Model.mychat_doubao import MyChatModel
from langchain_core.prompts import ChatPromptTemplate, MessagesPlaceholder

def create_agent():
    chat = MyChatModel()
    llm = chat.get_langchain_llm()
    tools = [insert_images_to_docx, read_text_auto, save_to_docx]
    prompt = ChatPromptTemplate.from_messages([
        ("system", "可使用以下工具：1) read_text_auto 读取/预览模板；2) save_to_docx 保存文本；3) insert_images_to_docx 执行图片插入。必须调用 insert_images_to_docx，并使用参数 template_path={template_path}，output_path={output_path}，static_dir={static_dir}。不要生成其它描述性文本。"),
        ("human", "开始执行图片插入。"),
        MessagesPlaceholder(variable_name="agent_scratchpad"),
    ])
    agent = create_tool_calling_agent(llm=llm, tools=tools, prompt=prompt)
    executor = AgentExecutor(agent=agent, tools=tools, verbose=True)
    tpl = os.environ.get("TEMPLATE_REPORT_PATH") or "报告模板.docx"
    ts = time.strftime("%Y%m%d_%H%M%S")
    out = os.path.abspath(f"模板_插图_{ts}.docx")
    static_dir = os.environ.get("STATIC_DIR") or "static"
    data = {
        "template_path": tpl,
        "output_path": out,
        "static_dir": static_dir,
    }
    try:
        rs = executor.invoke(data)
        print(rs)
        print(f"输出文件：{out}")
        if not os.path.exists(out):
            print("提示：未检测到输出文件，直接调用工具执行一次...")
            res = insert_images_to_docx.invoke({
                "template_path": tpl,
                "output_path": out,
                "static_dir": static_dir,
            })
            print(res)
            print(f"输出文件：{out}")
    except Exception as e:
        try:
            res = insert_images_to_docx.invoke({
                "template_path": tpl,
                "output_path": out,
                "static_dir": static_dir,
            })
            print(res)
            print(f"输出文件：{out}")
        except Exception as ee:
            print(str(ee))

if __name__ == "__main__":
    s = time.time()
    create_agent()
    e = time.time()
    print("耗时:", e - s)