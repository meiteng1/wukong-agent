import pandas as pd
from collections import defaultdict

# ===== 构件类别映射 =====
# 用于将不同编号的构件统一归类
component_category_map = {
    "梁体": "梁体",
    "墩": "桥墩及墩台",
    "墩台": "桥墩及墩台",
    "垫石": "垫石",
    "支座板": "支座板",
    "防落梁块": "防落梁块",
    "支座": "球形支座",
    "螺栓":"防落梁块",
    "指针":"刻度",
}

# ===== 统计并生成报告 =====
def generate_report(input_file, output_file_txt):
    xls = pd.ExcelFile(input_file)
    all_reports = []

    for sheet in xls.sheet_names:
        df = pd.read_excel(input_file, sheet_name=sheet)

        # 确保列存在
        for col in ["构件", "缺陷类型"]:
            if col not in df.columns:
                raise ValueError(f"Sheet {sheet} 缺少必要列: {col}")

        # 统计构件大类-缺陷类型出现次数
        stats = defaultdict(lambda: defaultdict(int))
        for _, row in df.iterrows():
            comp_full = str(row["构件"]).strip()
            defect = str(row["缺陷类型"]).strip()

            # 映射成构件大类
            comp_category = None
            for key in component_category_map:
                if key in comp_full:
                    comp_category = component_category_map[key]
                    break
            if comp_category is None:
                comp_category = "未知构件"

            stats[comp_category][defect] += 1

        # 按构件大类生成文字
        report_lines = [f"Sheet: {sheet}"]
        for comp_cat, defect_dict in stats.items():
            total = sum(defect_dict.values())
            defects_list = [f"{k}_{v}处" for k, v in defect_dict.items()]
            defects_str = "，".join(defects_list)
            report_lines.append(f"{comp_cat}共发现缺陷{total}处，其中{defects_str}。")
        all_reports.append("\n".join(report_lines))

    # 写入文本文件
    with open(output_file_txt, "w", encoding="utf-8") as f:
        f.write("\n\n".join(all_reports))

    print(f"✅ 统计完成，报告已生成：{output_file_txt}")


# ===== 程序入口 =====
if __name__ == "__main__":
    input_excel = r"F:\总结.xlsx"  # 新生成的表格
    output_txt = r"F:\缺陷统计报告.txt"
    generate_report(input_excel, output_txt)
