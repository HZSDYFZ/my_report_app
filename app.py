# -*- coding: utf-8 -*-
import io
import re
import zipfile
import pandas as pd
import streamlit as st
from docx import Document
from datetime import datetime

st.set_page_config(page_title="认证评定报告自动化生成系统", page_icon="📄", layout="wide")

BOX_PATTERN = re.compile(r"[□☐\[\]口]")

def extract_first_person(lead_str):
    """提取审核组长姓名"""
    if pd.isna(lead_str) or not str(lead_str).strip():
        return ""
    s = str(lead_str).strip()
    s = re.sub(r"^(审核组长|组长|Lead)[:：]\s*", "", s, flags=re.IGNORECASE)
    s = re.sub(r"[\（\(].*?[\）\)]", "", s).strip()
    parts = re.split(r"[ ,，/、+&\t\n]+", s)
    return parts[0] if parts and parts[0] else ""

def get_clean_col_val(row, possible_keys, default=""):
    """多列名模糊匹配与格式化"""
    for key in possible_keys:
        for col in row.index:
            col_clean = str(col).replace(" ", "").replace("\n", "").lower()
            key_clean = key.replace(" ", "").lower()
            if key_clean in col_clean:
                val = row[col]
                if pd.isna(val):
                    continue
                if isinstance(val, (pd.Timestamp, datetime)) or "date" in type(val).__name__.lower():
                    return str(val).split(" ")[0]
                val_str = str(val).strip()
                if val_str and val_str.lower() not in ["nan", "none", "null", "nat", "0", "undefined"]:
                    return val_str
    return default

def is_real_data_row(row):
    """过滤 Excel 尾部空白/汇总行"""
    if row.dropna().empty:
        return False
    comp = get_clean_col_val(row, ["公司名称", "客户名称", "企业名称", "公司"], default="")
    task = get_clean_col_val(row, ["任务号", "合同号", "项目编号"], default="")
    lead = get_clean_col_val(row, ["审核组长", "组长", "审核员"], default="")

    invalid_words = ["", "nan", "none", "null", "未知企业", "未填写", "0"]
    if comp.lower() in invalid_words and task.lower() in invalid_words and lead.lower() in invalid_words:
        return False
    if any(kw in comp for kw in ["合计", "小计", "统计", "说明", "备注", "填表说明"]):
        return False
    return True

def process_row_data(row, index):
    """提取 Excel 数据并计算结论勾选逻辑"""
    company_name = get_clean_col_val(row, ["公司名称", "客户名称", "企业名称", "公司"], default="")
    task_no = get_clean_col_val(row, ["任务号", "合同号", "项目编号"], default=f"TASK_{index+1}")
    lead_raw = get_clean_col_val(row, ["审核组长", "组长", "审核员"], default="")
    address = get_clean_col_val(row, ["审核地址", "地址", "企业地址"], default="")
    scope = get_clean_col_val(row, ["审核范围", "认证范围", "范围"], default="")
    audit_type_raw = get_clean_col_val(row, ["审核类型", "audit type"], default="")
    eval_date = get_clean_col_val(row, ["评定日期", "决定日期", "日期"], default="")

    lead_first = extract_first_person(lead_raw)

    # 结论勾选判定:
    # 红色 (初审/再认证)   -> 第 1 个勾
    # 黄色 (转移)          -> 第 3 个勾
    # 青色 (监一/监二/监审)-> 第 5 个勾
    decision_option = 0
    if "转移" in audit_type_raw:
        decision_option = 3
    elif "监" in audit_type_raw:
        decision_option = 5
    elif "初" in audit_type_raw or "再认证" in audit_type_raw:
        decision_option = 1

    return {
        "company_name": company_name,
        "task_no": task_no,
        "lead": lead_first,
        "address": address,
        "scope": scope,
        "audit_type_raw": audit_type_raw,
        "decision_option": decision_option,
        "eval_date": eval_date
    }

def replace_placeholders_in_paragraph(p, data):
    """段落文本内精准替换占位符，不损毁样式"""
    if not p.text:
        return
    replacements = {
        "{{公司名称}}": data["company_name"],
        "{{客户名称}}": data["company_name"],
        "{{任务号}}": data["task_no"],
        "{{审核组长}}": data["lead"],
        "{{组长}}": data["lead"],
        "{{审核地址}}": data["address"],
        "{{地址}}": data["address"],
        "{{审核范围}}": data["scope"],
        "{{范围}}": data["scope"],
        "{{评定日期}}": data["eval_date"],
        "{{日期}}": data["eval_date"],
        "{{审核类型}}": data["audit_type_raw"],
    }
    text = p.text
    for k, v in replacements.items():
        if k in text:
            text = text.replace(k, str(v))
    if text != p.text:
        p.text = text

def process_table_safely(table, data):
    """去重保护合并单元格，仅修改 paragraph.text，不破坏表格框架"""
    visited_tcs = set()

    for row in table.rows:
        cells = row.cells
        for idx, cell in enumerate(cells):
            # 防止重复修改合并单元格 (colspan / rowspan)
            if cell._tc in visited_tcs:
                continue
            visited_tcs.add(cell._tc)

            cell_text = cell.text.strip()
            clean_cell_text = cell_text.replace(" ", "").replace("\n", "")

            # 1. 替换单元格内已有的占位符
            for p in cell.paragraphs:
                replace_placeholders_in_paragraph(p, data)

            # 2. 认证决定结论区域：精准处理第 N 个复选框与评定日期
            if "认证决定" in clean_cell_text or "决定结论" in clean_cell_text:
                opt_target = data["decision_option"]
                if opt_target > 0:
                    box_count = 0
                    for p in cell.paragraphs:
                        if BOX_PATTERN.search(p.text):
                            def replace_box(m):
                                nonlocal box_count
                                box_count += 1
                                return "☑" if box_count == opt_target else "☐"
                            p.text = BOX_PATTERN.sub(replace_box, p.text)
                
                # 回填评定日期
                if data["eval_date"]:
                    for p in cell.paragraphs:
                        if "日期" in p.text and ("：" in p.text or ":" in p.text or "{{" in p.text):
                            if "{{" in p.text:
                                p.text = re.sub(r"\{\{.*?\}\}", data["eval_date"], p.text)
                            else:
                                prefix = re.split(r"[:：]", p.text)[0]
                                p.text = f"{prefix}：{data['eval_date']}"

            # 3. 模板纯文本绑定（如“公司名称”匹配右侧格子）
            mappings = [
                (["公司名称", "客户名称"], data["company_name"]),
                (["任务号", "合同号"], data["task_no"]),
                (["审核组长", "组长"], data["lead"]),
                (["审核地址", "企业地址"], data["address"]),
                (["审核范围", "认证范围"], data["scope"]),
            ]

            for keywords, val in mappings:
                if any(kw in clean_cell_text for kw in keywords) and val:
                    # 情况 A: 单元格形如 "公司名称：" 且后方为空
                    if ("：" in cell_text or ":" in cell_text) and not cell_text.split("：")[-1].strip().split(":")[-1].strip():
                        prefix = re.split(r"[:：]", cell_text)[0]
                        if cell.paragraphs:
                            cell.paragraphs[0].text = f"{prefix}：{val}"
                    # 情况 B: 当前格为标题，值写在右侧相邻单元格
                    elif idx + 1 < len(cells):
                        next_cell = cells[idx + 1]
                        next_text = next_cell.text.strip()
                        if not next_text or "{{" in next_text:
                            if next_cell.paragraphs:
                                next_cell.paragraphs[0].text = str(val)

def fill_word_template(template_bytes, data):
    doc = Document(io.BytesIO(template_bytes))

    # 1. 替换正文段落
    for p in doc.paragraphs:
        replace_placeholders_in_paragraph(p, data)

    # 2. 遍历表格填充
    for table in doc.tables:
        process_table_safely(table, data)

    out_stream = io.BytesIO()
    doc.save(out_stream)
    out_stream.seek(0)
    return out_stream.getvalue()

# ==========================================
# Streamlit 前端界面
# ==========================================
st.title("🛡️ 认证评定报告自动化生成系统")

c1, c2 = st.columns(2)
with c1:
    excel_file = st.file_uploader("1. 上传认证 Excel 数据文件 (.xlsx / .xls)", type=["xlsx", "xls"])
with c2:
    template_file = st.file_uploader("2. 上传 Word 报告模板 (.docx)", type=["docx"])

st.markdown("---")

if excel_file is not None and template_file is not None:
    try:
        raw_df = pd.read_excel(excel_file)
        template_bytes = template_file.getvalue()

        parsed_records = []
        for idx, row in raw_df.iterrows():
            if is_real_data_row(row):
                parsed_records.append(process_row_data(row, idx))

        if not parsed_records:
            st.warning("⚠️ Excel 文件中未读取到有效数据。")
        else:
            st.subheader(f"📋 提取数据预览（共 {len(parsed_records)} 条记录）")
            preview_df = pd.DataFrame(parsed_records)[
                ["company_name", "task_no", "lead", "audit_type_raw", "decision_option", "eval_date"]
            ]
            preview_df.columns = ["公司名称", "任务号", "审核组长", "审核类型", "勾选结论位置", "评定日期"]
            
            preview_df["勾选结论位置"] = preview_df["勾选结论位置"].map({
                1: "第 1 项 (初审/再认证)",
                3: "第 3 项 (转移)",
                5: "第 5 项 (监一/监二)"
            }).fillna("未定义")

            st.dataframe(preview_df, use_container_width=True)

            st.markdown("### 🚀 开始生成")

            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                for idx, data in enumerate(parsed_records):
                    doc_bytes = fill_word_template(template_bytes, data)
                    clean_company = re.sub(r'[\\/*?:"<>|]', "_", str(data["company_name"]) or "未命名企业")
                    clean_task = re.sub(r'[\\/*?:"<>|]', "_", str(data["task_no"]) or "未命名任务")
                    zf.writestr(f"{clean_company}_{clean_task}_评定报告.docx", doc_bytes)

            zip_buffer.seek(0)

            st.download_button(
                label=f"📦 一键下载 Word 报告包 ({len(parsed_records)} 份 .zip)",
                data=zip_buffer.getvalue(),
                file_name="认证评定报告包.zip",
                mime="application/zip",
                type="primary",
                use_container_width=True
            )

    except Exception as e:
        st.error(f"发生错误: {str(e)}")
        st.exception(e)
else:
    st.info("👈 请在上方上传 **Excel 数据文件** 和 **Word 模板文件**。")
