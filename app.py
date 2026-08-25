# -*- coding: utf-8 -*-
import io
import re
import zipfile
import pandas as pd
import streamlit as st
from docx import Document

st.set_page_config(page_title="认证评定报告自动化生成系统", page_icon="📄", layout="wide")

BOX_CHARS = r"[□☐☑✔\[\]口]"

def extract_first_person(lead_str):
    """提取组长姓名"""
    if pd.isna(lead_str) or not str(lead_str).strip():
        return ""
    s = str(lead_str).strip()
    s = re.sub(r"^(审核组长|组长|Lead)[:：]\s*", "", s, flags=re.IGNORECASE)
    s = re.sub(r"[\（\(].*?[\）\)]", "", s).strip()
    parts = re.split(r"[ ,，/、+&\t\n]+", s)
    return parts[0] if parts and parts[0] else ""

def get_clean_col_val(row, possible_keys, default="未填写"):
    """多列名模糊匹配"""
    for key in possible_keys:
        for col in row.index:
            col_clean = str(col).replace(" ", "").replace("\n", "").lower()
            key_clean = key.replace(" ", "").lower()
            if key_clean in col_clean:
                val = row[col]
                if pd.isna(val):
                    continue
                # 日期格式化处理
                if isinstance(val, (pd.Timestamp, datetime)) or "date" in type(val).__name__.lower():
                    return str(val).split(" ")[0]
                val_str = str(val).strip()
                if val_str and val_str.lower() not in ["nan", "none", "null", "nat", "0", "undefined"]:
                    return val_str
    return default

from datetime import datetime

def is_real_data_row(row):
    """过滤 Excel 尾部空白行"""
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
    """提取单行 Excel 数据并计算勾选逻辑"""
    company_name = get_clean_col_val(row, ["公司名称", "客户名称", "企业名称", "公司"], default="未填写公司")
    task_no = get_clean_col_val(row, ["任务号", "合同号", "项目编号"], default=f"TASK_{index+1}")
    lead_raw = get_clean_col_val(row, ["审核组长", "组长", "审核员", "组长姓名"], default="")
    address = get_clean_col_val(row, ["审核地址", "地址", "企业地址"], default="未填写地址")
    scope = get_clean_col_val(row, ["审核范围", "认证范围", "范围"], default="未填写范围")
    audit_type_raw = get_clean_col_val(row, ["审核类型", "audit type"], default="")
    eval_date = get_clean_col_val(row, ["评定日期", "决定日期", "日期", "eval date"], default="")

    # 提取第一组长
    lead_first = extract_first_person(lead_raw) or "未填写组长"

    # 标准匹配
    task_no_upper = task_no.upper()
    has_ts = "TS" in task_no_upper
    has_er = "ER" in task_no_upper

    # 结论勾选位置逻辑判定
    # 红色（初审/再认证） -> 第1个勾
    # 黄色（转移）        -> 第3个勾
    # 青色（监一/监二/监审）-> 第5个勾
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
        "has_ts": has_ts,
        "has_er": has_er,
        "audit_type_raw": audit_type_raw,
        "decision_option": decision_option,  # 1, 3, 5
        "eval_date": eval_date
    }

def set_nth_checkbox(text, target_index):
    """把文本中的第 target_index 个复选框设为 ☑，其余保持/设为 ☐"""
    count = 0
    def repl(m):
        nonlocal count
        count += 1
        return "☑" if count == target_index else "☐"
    return re.sub(BOX_CHARS, repl, text)

def fill_table_safely(table, data):
    """非破坏性填充表格及决定结论处理"""
    for row in table.rows:
        cells = row.cells
        row_full_text = "".join(c.text.strip() for c in cells)

        # 1. 认证决定结论区域处理（勾选第 1/3/5 项）
        if "结论" in row_full_text or "决定" in row_full_text or "同意" in row_full_text:
            for cell in cells:
                # 勾选指定索引的选项
                if re.search(BOX_CHARS, cell.text) and data["decision_option"] > 0:
                    cell.text = set_nth_checkbox(cell.text, data["decision_option"])
                
                # 日期替换
                if "日期" in cell.text and data["eval_date"]:
                    if "{{" in cell.text:
                        cell.text = re.sub(r"\{\{.*?\}\}", data["eval_date"], cell.text)
                    elif "：" in cell.text or ":" in cell.text:
                        prefix = cell.text.split("：")[0] if "：" in cell.text else cell.text.split(":")[0]
                        cell.text = f"{prefix}：{data['eval_date']}"

        # 2. 常规数据行强填（不改左列，只填右列或替换占位符）
        for idx, cell in enumerate(cells):
            # 处理 {{占位符}}
            if "{{" in cell.text:
                cell.text = cell.text.replace("{{公司名称}}", data["company_name"])\
                                     .replace("{{任务号}}", data["task_no"])\
                                     .replace("{{审核组长}}", data["lead"])\
                                     .replace("{{审核地址}}", data["address"])\
                                     .replace("{{审核范围}}", data["scope"])\
                                     .replace("{{评定日期}}", data["eval_date"])

            raw_text = cell.text.strip()
            clean_text = raw_text.replace(" ", "").replace("\n", "")

            # 根据标题补入右格
            mappings = [
                (["公司名称", "客户名称"], data["company_name"]),
                (["任务号", "合同号"], data["task_no"]),
                (["审核组长", "组长"], data["lead"]),
                (["审核地址", "企业地址"], data["address"]),
                (["审核范围", "认证范围"], data["scope"]),
            ]
            for keywords, val in mappings:
                if any(kw in clean_text for kw in keywords):
                    if idx + 1 < len(cells):
                        next_cell_text = cells[idx + 1].text.strip()
                        if not next_cell_text or "{{" in next_cell_text:
                            cells[idx + 1].text = str(val)

def fill_word_template(template_bytes, data):
    doc = Document(io.BytesIO(template_bytes))

    # 替换正文占位符
    for p in doc.paragraphs:
        if "{{" in p.text:
            p.text = p.text.replace("{{公司名称}}", data["company_name"])\
                           .replace("{{任务号}}", data["task_no"])\
                           .replace("{{审核组长}}", data["lead"])\
                           .replace("{{审核地址}}", data["address"])\
                           .replace("{{审核范围}}", data["scope"])\
                           .replace("{{评定日期}}", data["eval_date"])

    # 填充表格
    for table in doc.tables:
        fill_table_safely(table, data)

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
                ["company_name", "task_no", "lead", "audit_type_raw", "decision_option", "eval_date", "address"]
            ]
            preview_df.columns = ["公司名称", "任务号", "审核组长", "审核类型", "勾选结论项", "评定日期", "审核地址"]
            
            # 显示解析后的结论选项说明
            preview_df["勾选结论项"] = preview_df["勾选结论项"].map({
                1: "第 1 项 (初审/再认证)",
                3: "第 3 项 (转移)",
                5: "第 5 项 (监一/监二)"
            }).fillna("未匹配")

            st.dataframe(preview_df, use_container_width=True)

            st.markdown("### 🚀 开始生成")

            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                for idx, data in enumerate(parsed_records):
                    doc_bytes = fill_word_template(template_bytes, data)
                    clean_company = re.sub(r'[\\/*?:"<>|]', "_", str(data["company_name"]))
                    clean_task = re.sub(r'[\\/*?:"<>|]', "_", str(data["task_no"]))
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
