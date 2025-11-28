import os
import re
import pandas as pd
from Tool.documentRead_tool import read_text_auto
try:
    from langchain.tools import tool
except Exception:
    def tool(*args, **kwargs):
        def _wrap(f):
            return f
        return _wrap

def load_env_file():
    """加载.env文件环境变量"""
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

def get_base_number(pier_code):
    """提取桥墩编号中的数字部分作为基础编号"""
    if pier_code is None:
        return 0
    pier_code = str(pier_code).strip()
    numbers = re.findall(r"\d+", pier_code)
    pier_number = int(numbers[0]) if numbers else 0
    return pier_number

def get_component_name(defect_type):
    """根据缺陷类型映射构件名称（行业规范映射）"""
    mapping = [
        ("桥墩", "墩"), ("垃圾残留", "墩"), ("墩台", "墩"), ("落水管", "墩"), ("梁体", "梁"),
        ("垫石", "垫石"), ("环氧砂浆", "垫石"), ("涂装漆", "垫石"), ("麻面", "垫石"),
        ("支座板", "支座板"), ("连接件", "支座板"), ("螺栓", "防落梁块"),
        ("防滑块", "防落梁块"), ("梁块", "防落梁块"), ("预埋件", "防落梁块"),
        ("球形支座", "支座"), ("防尘围挡", "支座"), ("刻度", "支座"),
        ("缺棱断角", "垫石"), ("破损", "垫石"), ("掉角", "垫石")
    ]
    for keyword, comp in mapping:
        if keyword in str(defect_type):
            return comp
    return "未知构件"

def generate_component(row):
    """生成符合模板规则的构件编号（梁体从1开始，其余从0开始）"""
    pier_code = row.get("桥墩编号", "")
    defect_type = row.get("缺陷类型", "")
    base = get_base_number(pier_code)
    component_name = get_component_name(defect_type)
    if component_name == "梁":
        return f"{base + 1}#梁"
    else:
        return f"{base}#{component_name}"

def _deduplicate_rows(rows):
    """保留原始行（与 handle_fault.py 保持一致，不做去重）"""
    return rows

def _format_rows(df):
    df = df.copy()
    formatted_cols = ["桥墩", "构件", "部位", "缺陷类型", "现场照片"]
    if set(formatted_cols).issubset(df.columns):
        df = df[formatted_cols].fillna("")
        return df.to_dict(orient="records")
    for col in ["桥墩编号", "缺陷类型", "缺陷部位（里程/侧别）"]:
        if col not in df.columns:
            df[col] = ""
        df[col] = df[col].fillna("")
    df["构件"] = df.apply(generate_component, axis=1)
    df["桥墩"] = df["桥墩编号"]
    df["部位"] = df["缺陷部位（里程/侧别）"]
    df["现场照片"] = df.apply(lambda r: f"{r['桥墩编号']}-{r['部位']}{r['缺陷类型']}.jpg", axis=1)
    return df[["桥墩", "构件", "部位", "缺陷类型", "现场照片"]].to_dict(orient="records")

def _read_all_sheets(input_file):
    """读取Excel所有sheet并合并处理（严格对齐 handle_fault.py）"""
    xls = pd.ExcelFile(input_file)
    all_rows = []
    for sheet in xls.sheet_names:
        df = pd.read_excel(input_file, sheet_name=sheet)
        rows = _format_rows(df)
        all_rows.extend(rows)
    return all_rows

def _parse_text_rows(text):
    """保留占位（与 handle_fault.py 一致，不解析文本）"""
    return []

def _split_tables(rows):
    """按构件类型拆分表格（表3.1.1：梁/墩；表3.2.1：支座系统）"""
    table_31 = []  # 梁体、桥墩、墩台
    table_32 = []  # 支座系统（垫石、支座板等）
    for r in rows:
        comp = r.get("构件", "")
        if any(tag in comp for tag in ["#梁", "#墩"]):
            table_31.append({
                "pier": r.get("桥墩", ""),
                "component": r.get("构件", ""),
                "position": r.get("部位", ""),
                "defect_type": r.get("缺陷类型", ""),
                "photo": r.get("现场照片", "")
            })
        elif any(tag in comp for tag in ["#垫石", "#支座板", "#防落梁块", "#支座"]):
            table_32.append({
                "pier": r.get("桥墩", ""),
                "component": r.get("构件", ""),
                "position": r.get("部位", ""),
                "defect_type": r.get("缺陷类型", ""),
                "photo": r.get("现场照片", "")
            })
    return table_31, table_32

@tool
def read_and_format_defects(input_file: str = None) -> dict:
    """
    读取并格式化缺陷汇总表：合并所有Sheet，按【桥墩+部位+缺陷类型】去重，输出两类表格数据。

    参数:
        input_file: 缺陷汇总Excel路径；未提供时自动使用.env的RAW_REPORT_PATH。

    返回:
        字典：{"beam_pier_defects": [...], "support_system_defects": [...]}，
        每条记录包含pier/component/position/defect_type/photo五字段，符合表3.1.1和3.2.1格式要求。
    """
    # 自动获取文件路径（优先input_file，其次.env）
    fpath = input_file
    if not fpath or not os.path.exists(fpath):
        fpath = os.getenv("REFER_FILE_PATH") or os.getenv("RAW_REPORT_PATH")
    if not fpath or not os.path.exists(fpath):
        raise FileNotFoundError("未找到缺陷汇总表，请检查路径或.env配置")
    
    # 根据文件类型处理数据
    ext = os.path.splitext(fpath)[1].lower()
    rows = _read_all_sheets(fpath)
    
    # 拆分表格并返回
    beam_pier, support_system = _split_tables(rows)
    return {
        "beam_pier_defects": beam_pier,
        "support_system_defects": support_system
    }

@tool
def export_formatted_excel(input_file: str = None, output_file: str = None) -> str:
    """
    导出标准化缺陷表：将去重后的五列表（桥墩、构件、部位、缺陷类型、现场照片）导出到Excel。

    参数:
        input_file: 缺陷汇总Excel路径；未提供时使用.env的RAW_REPORT_PATH。
        output_file: 输出Excel路径；未提供时默认“缺陷汇总_格式化.xlsx”。

    返回:
        输出文件路径字符串。
    """
    fpath = input_file
    if not fpath or not os.path.exists(fpath):
        fpath = os.getenv("REFER_FILE_PATH") or os.getenv("RAW_REPORT_PATH")
    if not fpath or not os.path.exists(fpath):
        raise FileNotFoundError("未找到缺陷汇总表，请检查路径或.env配置")
    
    # 处理输出路径
    out_path = output_file or os.getenv("REFER_FILE_OUT_PATH")
    xls = pd.ExcelFile(fpath)
    try:
        with pd.ExcelWriter(out_path, engine='openpyxl') as writer:
            for sheet in xls.sheet_names:
                df = pd.read_excel(fpath, sheet_name=sheet)
                formatted_rows = _format_rows(df)
                df_out = pd.DataFrame(formatted_rows)
                df_out.columns = ["桥墩", "构件", "部位", "缺陷类型", "现场照片"]
                df_out.to_excel(writer, sheet_name=sheet, index=False)
        return out_path
    except PermissionError:
        base_dir = os.path.join(os.getcwd(), "输出")
        try:
            os.makedirs(base_dir, exist_ok=True)
        except Exception:
            base_dir = os.getcwd()
        alt_path = os.path.join(base_dir, os.path.basename(out_path) or "缺陷汇总_格式化.xlsx")
        with pd.ExcelWriter(alt_path, engine='openpyxl') as writer:
            for sheet in xls.sheet_names:
                df = pd.read_excel(fpath, sheet_name=sheet)
                formatted_rows = _format_rows(df)
                df_out = pd.DataFrame(formatted_rows)
                df_out.columns = ["桥墩", "构件", "部位", "缺陷类型", "现场照片"]
                df_out.to_excel(writer, sheet_name=sheet, index=False)
        return alt_path

class HandleFaultHandler:
    """工具类封装，便于外部调用"""
    @staticmethod
    def read_and_format_defects(input_file: str = None) -> dict:
        return read_and_format_defects(input_file)
    
    @staticmethod
    def export_formatted_excel(input_file: str = None, output_file: str = None) -> str:
        return export_formatted_excel(input_file, output_file)

# 工具列表，供agent调用
HANDLE_FAULT_TOOLS = [read_and_format_defects, export_formatted_excel]
