import pandas as pd
import re

# ===== 解析桥墩编号，通用方法 =====
def get_base_number(pier_code):
    """
    支持各种桥墩编号形式：
    - 纯数字: 3, 01, 001
    - 字母+数字: QR-02, RH-11
    - 数字+字母: 2-Y
    - 特殊前缀统一用数字提取
    """
    if pier_code is None:
        return 0

    pier_code = str(pier_code).strip()
    numbers = re.findall(r"\d+", pier_code)
    pier_number = int(numbers[0]) if numbers else 0
    return pier_number

# ===== 构件关键词映射 =====
def get_component_name(defect_type):
    mapping = [
        ("桥墩", "墩"), ("垃圾残留", "墩"), ("墩台", "墩"),("落水管","墩"),("梁体","梁"),
        ("垫石", "垫石"), ("环氧砂浆", "垫石"), ("涂装漆", "垫石"), ("麻面", "垫石"),
        ("支座板", "支座板"), ("连接件", "支座板"),("螺栓","防落梁块"),
        ("防滑块", "防落梁块"), ("梁块", "防落梁块"), ("预埋件", "防落梁块"),
        ("球形支座", "支座"), ("防尘围挡", "支座"), ("刻度", "支座"),
        ("缺棱断角", "垫石"), ("破损", "垫石"), ("掉角", "垫石")
    ]
    for keyword, comp in mapping:
        if keyword in str(defect_type):
            return comp
    return "未知构件"

# ===== 生成构件编号 =====
def generate_component(row):
    pier_code = row["桥墩编号"]
    defect_type = row["缺陷类型"]
    base = get_base_number(pier_code)

    component_name = get_component_name(defect_type)

    if component_name == "梁":
        return f"{base + 1}#梁"
    else:
        return f"{base}#{component_name}"

# ===== 主程序 =====
def process_excel():
    input_file = r"F:\厦门轨道3号线和4号线桥梁支座缺陷\缺陷汇总表.xlsx"
    output_file = r"F:\总结.xlsx"

    xls = pd.ExcelFile(input_file)
    writer = pd.ExcelWriter(output_file, engine='openpyxl')

    for sheet in xls.sheet_names:
        df = pd.read_excel(input_file, sheet_name=sheet)

        # 防止空值报错
        df["桥墩编号"] = df["桥墩编号"].fillna("")
        df["缺陷类型"] = df["缺陷类型"].fillna("")
        df["缺陷部位（里程/侧别）"] = df["缺陷部位（里程/侧别）"].fillna("")

        df["构件"] = df.apply(generate_component, axis=1)
        df["桥墩"] = df["桥墩编号"]
        df["部位"] = df["缺陷部位（里程/侧别）"]
        df["现场照片"] = df.apply(
            lambda r: f"{r['桥墩编号']}-{r['部位']}{r['缺陷类型']}.jpg", axis=1
        )

        output_df = df[["桥墩", "构件", "部位", "缺陷类型", "现场照片"]]
        output_df.to_excel(writer, sheet_name=sheet, index=False)

    writer.close()
    print(f"✅ 处理完成，输出文件：{output_file}")

# ===== 程序入口 =====
if __name__ == "__main__":
    process_excel()
