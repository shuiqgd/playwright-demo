import pandas as pd
import json
import sys
import os
from openpyxl.styles import Font, Alignment, Border, Side

def format_excel(worksheet, columns):
    """通用样式美化：加粗表头、边框、自动换行、列宽。"""
    header_font = Font(bold=True)
    thin_border = Border(
        left=Side(style='thin'), 
        right=Side(style='thin'), 
        top=Side(style='thin'), 
        bottom=Side(style='thin')
    )
    
    # 表头样式
    for cell in worksheet[1]:
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = thin_border

    # 内容样式
    for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row, max_col=len(columns)):
        for cell in row:
            cell.border = thin_border
            cell.alignment = Alignment(wrap_text=True, vertical='top')

    # 列宽设为 30
    for idx, _ in enumerate(columns):
        worksheet.column_dimensions[chr(65+idx)].width = 30

def generate_reports(test_results, output_dir="output"):
    """
    生成两个报告：
    1. bugList.xlsx: 仅包含状态为 '失败' 的用例。
    2. 测试成果.xlsx: 包含所有执行用例。
    """
    os.makedirs(output_dir, exist_ok=True)
    
    # 1. 生成测试成果 (All Results)
    df_all = pd.DataFrame(test_results)
    all_cols = ["模块", "用例 ID", "测试项", "重要程度", "状态", "结果"]
    # 确保列齐全
    for col in all_cols:
        if col not in df_all.columns: df_all[col] = ""
    df_all = df_all[all_cols]
    
    res_path = os.path.join(output_dir, "测试成果.xlsx")
    with pd.ExcelWriter(res_path, engine='openpyxl') as writer:
        df_all.to_excel(writer, index=False, sheet_name="测试汇总")
        format_excel(writer.sheets["测试汇总"], all_cols)
    print(f"成功生成测试成果: {os.path.abspath(res_path)}")

    # 2. 生成 Bug 列表 (Only Failures)
    df_bugs = df_all[df_all["状态"].str.contains("失败", na=False)].copy()
    if not df_bugs.empty:
        # 转换为 Bug 列表格式要求
        # 假设 test_results 中已经包含了详细的 '重现步骤' 格式字符串
        bug_cols = ["Bug 标题", "严重程度", "优先级", "重现步骤"]
        df_bugs["Bug 标题"] = df_bugs["测试项"] + " - 执行失败"
        df_bugs["严重程度"] = "一般" # 默认值，可由 AI 在生成结果时指定
        df_bugs["优先级"] = "中"
        # 重现步骤需要拼接 [步骤][结果][期望]
        if "重现步骤" not in df_bugs.columns:
            df_bugs["重现步骤"] = "[步骤]\n" + df_bugs["测试项"] + "\n[结果]\n测试失败\n[期望]\n通过测试"
            
        df_bugs = df_bugs[bug_cols]
        bug_path = os.path.join(output_dir, "bugList.xlsx")
        with pd.ExcelWriter(bug_path, engine='openpyxl') as writer:
            df_bugs.to_excel(writer, index=False, sheet_name="Bug清单")
            format_excel(writer.sheets["Bug清单"], bug_cols)
        print(f"成功生成 Bug 列表: {os.path.abspath(bug_path)}")
    else:
        print("未发现失败用例，不生成 bugList.xlsx")

if __name__ == "__main__":
    try:
        if len(sys.argv) > 1:
            with open(sys.argv[1], 'r', encoding='utf-8') as f:
                data = json.load(f)
        else:
            print("请输入测试结果 JSON 数据 (Ctrl+D 或 Ctrl+Z 结束):")
            data = json.load(sys.stdin)
        
        generate_reports(data)
    except Exception as e:
        print(f"生成失败: {e}")
