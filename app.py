# -*- coding: utf-8 -*-
import io
import re
import zipfile
import pandas as pd
import streamlit as st
from docx import Document
from datetime import datetime

st.set_page_config(page_title="认证评定报告自动化生成系统", page_icon="📄", layout="wide")

BOX_PATTERN = re.compile(r"^[□☐☑✔\[\]口\s]+")

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
    """表头多别名模糊匹配"""
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
    """过滤 Excel 尾部无效空行"""
    if row.dropna().empty:
        return False
    comp = get_clean_col_val(row, ["公司名称", "客户名称", "企业名称", "单位名称", "公司"], default="")
    task = get_clean_col_val(row, ["任务号", "合同号", "项目编号"], default="")
    lead = get_clean_col_val(row, ["审核组长", "组长", "审核员", "组长姓名"], default="")

    invalid_words = ["", "nan", "none", "null", "未知企业", "未填写", "0"]
    if comp.lower() in invalid_words and task.lower() in invalid_words and lead.lower() in invalid_words:
        return False
    if any(kw in comp for kw in ["合计", "小计", "统计", "说明", "备注", "填表说明"]):
        return False
    return True

def process_row_data(row, index):
    """单行 Excel 数据解析"""
    company_name = get_clean_col_val(row, ["公司名称", "客户名称", "企业名称", "单位名称", "公司"], default="")
    task_no = get_clean_col_val(row, ["任务号", "合同号", "项目编号", "单号"], default="")
    lead_raw = get_clean_col_val(row, ["审核组长", "组长", "审核员", "组长姓名", "lead", "姓名"], default="")
    address = get_clean_col_val(row, ["审核地址", "地址", "企业地址", "注册地址"], default="")
    scope = get_clean_col_val(row, ["审核范围", "认证范围", "范围", "业务范围"], default="")
    audit_type_raw = get_clean_col_val(row, ["审核类型", "类型", "审核阶段"], default="")
    eval_date = get_clean_col_val(row, ["评定日期", "决定日期", "日期", "评审日期", "评定时间", "通过时间"], default="")

    lead_first = extract_first_person(lead_raw)

    task_upper = task_no.upper()
    has_ts = "TS" in task_upper or "16949" in audit_type_raw
    has_er = "ER" in task_upper or "9001" in audit_type_raw

    decision_option = 1
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
        "decision_option": decision_option,
        "eval_date": eval_date
    }

def update_paragraph_checkboxes(p, data):
    """文本框及普通段落占位符与复选框替换"""
    text = p.text
    if not text.strip():
        return

    replacements = {
        "{{公司名称}}": data["company_name"],
        "{{任务号}}": data["task_no"],
        "{{审核组长}}": data["lead"],
        "{{审核地址}}": data["address"],
        "{{审核范围}}": data["scope"],
        "{{评定日期}}": data["eval_date"],
        "{{评定通过时间}}": data["eval_date"],
        "【公司名称】": data["company_name"],
        "【任务号】": data["task_no"],
        "【审核组长】": data["lead"],
        "【审核地址】": data["address"],
        "【审核范围】": data["scope"],
        "【评定日期】": data["eval_date"],
        "【评定通过时间】": data["eval_date"],
    }
    for k, v in replacements.items():
        if k in text:
            text = text.replace(k, str(v))

    if "16949" in text:
        sym = "☑" if data["has_ts"] else "☐"
        text = re.sub(r"[□☐☑✔]\s*(IATF\s*16949)", f"{sym} \\1", text, flags=re.I)
    if "9001" in text:
        sym = "☑" if data["has_er"] else "☐"
        text = re.sub(r"[□☐☑✔]\s*(ISO\s*9001)", f"{sym} \\1", text, flags=re.I)
    if "初审" in text:
        sym = "☑" if ("初" in data["audit_type_raw"] and "监" not in data["audit_type_raw"]) else "☐"
        text = re.sub(r"[□☐☑✔]\s*(初审)", f"{sym} \\1", text)
    if "监审" in text:
        sym = "☑" if "监" in data["audit_type_raw"] else "☐"
        text = re.sub(r"[□☐☑✔]\s*(监审)", f"{sym} \\1", text)
    if "再认证" in text or "转移" in text:
        sym = "☑" if ("再认证" in data["audit_type_raw"] or "转移" in data["audit_type_raw"]) else "☐"
        text = re.sub(r"[□☐☑✔]\s*(再认证/转移)", f"{sym} \\1", text)

    if text != p.text:
        p.text = text

def fill_next_target_cell(cells, current_idx, value):
    """基于底层 XML 单元格(_tc)精准定位标签格后面紧挨着的下一个独立单元格（用于审核组长）"""
    if not str(value).strip():
        return
    current_tc = cells[current_idx]._tc
    for next_idx in range(current_idx + 1, len(cells)):
        if cells[next_idx]._tc != current_tc:
            target_cell = cells[next_idx]
            if target_cell.paragraphs:
                target_cell.paragraphs[0].text = str(value)
            else:
                target_cell.add_paragraph(str(value))
            break

def format_decision_options(cell, data):
    """还原认证决定结论格式"""
    option_paragraphs = [p for p in cell.paragraphs if "通过" in p.text or "不予通过" in p.text]
    target_idx = data["decision_option"]

    for opt_i, p in enumerate(option_paragraphs, start=1):
        raw_text = p.text.strip()
        clean_text = BOX_PATTERN.sub("", raw_text).strip()

        is_selected = (opt_i == target_idx)
        mark = "☑ " if is_selected else "☐ "

        if opt_i == 1 and "适用于：" in clean_text:
            prefix = clean_text.split("适用于：")[0]
            clean_text = f"{prefix}适用于：{data['scope']}）"

        p.text = mark + clean_text

def process_table_safely(table, data):
    """表格安全遍历，严格区分：公司/任务号在同单元格冒号后，组长在后一个格子"""
    visited_tcs = set()

    for row in table.rows:
        cells = row.cells
        for idx, cell in enumerate(cells):
            if cell._tc in visited_tcs:
                continue
            visited_tcs.add(cell._tc)

            cell_text = cell.text.strip()
            clean_text = re.sub(r"[\s:：]", "", cell_text)

            # 内部占位符替换
            for p in cell.paragraphs:
                update_paragraph_checkboxes(p, data)

            # 1. 公司名称 -> 填在“公司名称：”后面（同一单元格内）
            if clean_text in ["公司名称", "客户名称", "企业名称"]:
                if cell.paragraphs:
                    cell.paragraphs[0].text = f"公司名称：{data['company_name']}"

            # 2. 任务号 -> 填在“任务号：”后面（同一单元格内）
            elif clean_text in ["任务号", "合同号"]:
                if cell.paragraphs:
                    cell.paragraphs[0].text = f"任务号：{data['task_no']}"

            # 3. 审核组长 -> 填在后面一个格子里面
            elif (clean_text in ["审核组长", "组长"]) and "报告评定人员" not in clean_text:
                fill_next_target_cell(cells, idx, data['lead'])

            # 4. 段落中的地址、范围、日期匹配
            for p in cell.paragraphs:
                p_clean = re.sub(r"\s+", "", p.text)
                if "审核地址：" in p_clean or "审核地址:" in p_clean:
                    p.text = f"审核地址：{data['address']}"
                elif "认证范围：" in p_clean or "认证范围:" in p_clean:
                    p.text = f"认证范围：{data['scope']}"
                elif "日期：" in p_clean and not p_clean.endswith(data['eval_date']):
                    if any(kw in p_clean for kw in ["报告评定人员", "专业支持人员", "审批"]):
                        prefix_text = p.text.split("日期：")[0]
                        p.text = f"{prefix_text}日期：{data['eval_date']}"

            # 5. 表格内的单独日期格子
            if clean_text in ["日期", "评定日期", "评定通过时间", "评定时间"]:
                if "：" in cell_text or ":" in cell_text:
                    cell.paragraphs[0].text = f"日期：{data['eval_date']}"

            # 6. 认证决定结论选项勾选
            if "认证决定结论" in clean_text or ("通过" in cell_text and "不予通过" in cell_text):
                format_decision_options(cell, data)

def fill_word_template(template_bytes, data):
    doc = Document(io.BytesIO(template_bytes))

    for p in doc.paragraphs:
        update_paragraph_checkboxes(p, data)

    for table in doc.tables:
        process_table_safely(table, data)

    for section in doc.sections:
        for header_p in section.header.paragraphs:
            update_paragraph_checkboxes(header_p, data)
        for footer_p in section.footer.paragraphs:
            update_paragraph_checkboxes(footer_p, data)

    out_stream = io.BytesIO()
    doc.save(out_stream)
    out_stream.seek(0)
    return out_stream.getvalue()

# Streamlit 界面
st.title("📄 认证评定报告自动化生成系统")

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

        parsed_records = [process_row_data(row, idx) for idx, row in raw_df.iterrows() if is_real_data_row(row)]

        if not parsed_records:
            st.warning("⚠️ Excel 文件中未读取到有效数据。")
        else:
            st.subheader(f"📋 预览数据（共 {len(parsed_records)} 条）")
            preview_df = pd.DataFrame(parsed_records)[
                ["company_name", "task_no", "lead", "audit_type_raw", "decision_option", "address", "scope", "eval_date"]
            ]
            preview_df.columns = ["公司名称", "任务号", "审核组长", "审核类型", "勾选结论选项", "审核地址", "认证范围", "评定/通过日期"]
            st.dataframe(preview_df, use_container_width=True)

            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                for idx, data in enumerate(parsed_records):
                    doc_bytes = fill_word_template(template_bytes, data)
                    clean_company = re.sub(r'[\\/*?:"<>|]', "_", str(data["company_name"]) or "未命名企业")
                    clean_task = re.sub(r'[\\/*?:"<>|]', "_", str(data["task_no"]) or "未命名任务")
                    zf.writestr(f"{clean_company}_{clean_task}_评定报告.docx", doc_bytes)

            zip_buffer.seek(0)

            st.download_button(
                label=f"📦 下载批量生成报告包 ({len(parsed_records)} 份 .zip)",
                data=zip_buffer.getvalue(),
                file_name="认证评定报告包.zip",
                mime="application/zip",
                type="primary",
                use_container_width=True
            )

    except Exception as e:
        st.error(f"处理失败: {str(e)}")
else:
    st.info("👈 请上传对应的 Excel 和 Word 模板文件进行处理。")
