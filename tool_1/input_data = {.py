input_data = {
    "input": "根据要求和输入的文档内容，完成桥梁支座检查报告的编写",
    # 报告基本信息
    "project_name": "厦门轨道交通桥梁支座检测项目",
    "bridge_name": "厦门轨道交通各区间桥梁支座",
    "bridge_code": "未提供，暂按'厦门轨道交通 + 区段名称'分类",
    "inspect_date": "未提供",
    
    # 缺陷和检查相关信息
    "inspection_result": "",
    "defect_summary": "",
    # 其他必要变量...
}
rs = agent_executor.invoke(input_data)