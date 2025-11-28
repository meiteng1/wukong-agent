import os
import sys
import time
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from tool_1.handle_fault_tool import read_and_format_defects
from Tool.word_tool import create_complete_report
from Tool.documentRead_tool import read_text_auto, save_to_docx
from Tool.word_Imagetool import insert_images_to_docx
from Tool.excel_reader_tool import read_filtered_excel_tables
import pandas as pd

def run():
    raw = os.environ.get("REFER_FILE_PATH") or os.environ.get("RAW_REPORT_PATH")
    tpl = os.environ.get("TEMPLATE_REPORT_PATH")
    ts = time.strftime("%Y%m%d_%H%M%S")
    out = os.path.abspath(f"表格插入测试_{ts}.docx")
    imgs_out = os.path.abspath(f"表格插入测试_插图_{ts}.docx")
    static_dir = os.environ.get("STATIC_DIR") or "static"
    if not raw or not os.path.exists(raw):
        print("缺陷汇总路径未配置或不存在")
        return
    if not tpl or not os.path.exists(tpl):
        print("模板路径未配置或不存在")
        return
    try:
        defects = read_and_format_defects.invoke({"input_file": raw})
        tables = read_filtered_excel_tables.invoke({"file_path": os.environ.get("REFER_FILE_OUT_PATH")})
        t31 = tables.get("table31", [])
        t32 = tables.get("table32", [])
        lines = list(t31) + list(t32)
        data = {
            "beam_pier_defects": defects.get("beam_pier_defects", []),
            "support_system_defects": defects.get("support_system_defects", []),
            "excel_filtered_table": "\n".join(lines),
            "table31": t31,
            "table32": t32
        }
        res = create_complete_report.invoke({"output_path": out, "data": data, "template_path": tpl})
        print(res)
        res_imgs = insert_images_to_docx.invoke({"template_path": out, "output_path": imgs_out, "static_dir": static_dir})
        print(res_imgs)
    except Exception as e:
        print(str(e))

if __name__ == "__main__":
    run()
