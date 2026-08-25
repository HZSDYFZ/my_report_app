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
    """提取组长字段中的第一个姓名"""
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
                val = str(row[col]).strip()
                if val and val.lower() not in ["nan", "none", "null", "nat", "0", "undefined"]:
                    return val
    return default

def is_real_data_row(row):
    """过滤 Excel 尾部空白行"""
    if row.dropna().empty:
        return False
    comp = get_clean_col_val(row, ["公司名称", "客户名称", "企业名称", "client name", "公司"], default="")
    task = get_clean_col_val(row, ["任务号", "file number", "合同号", "项目编号"], default="")
    lead = get_clean_col_val(row, ["审核组长", "组长", "lead", "auditor", "审核员", "团队"], default="")

    invalid_words = ["", "nan", "none", "null", "未知企业", "未填写", "0"]
    if comp.lower() in invalid_words and task.lower() in invalid_words and lead.lower() in invalid_words:
        return False
    if any(kw in comp for kw in ["合计", "小计", "统计", "说明", "备注", "填表说明"]):
        return False
    return True

def process_row_data(row, index):
    """提取单行 Excel 数据"""
    company_name = get_clean_col_val(row, ["公司名称", "客户名称", "企业名称", "client name", "公司"], default="未填写公司")
    task_no = get_clean_col_val(row, ["任务号", "file number", "合同号", "项目编号"], default=f"TASK_{index+1}")
    lead_raw = get_clean_col_val(row, ["审核组长", "组长", "lead", "auditor", "审核员", "团队", "组长姓名"], default="")
    address = get_clean_col_val(row, ["审核地址", "地址", "address", "企业地址"], default="未填写地址")
    scope = get_clean_col_val(row, ["审核范围", "认证范围", "范围", "scope"], default="未填写范围")
    audit_type_raw = get_clean_col_val(row, ["审核类型", "audit type"], default="")

    lead_first = extract_first_person(lead_raw) or "未填写组长"

    task_no_upper = task_no.upper()
    has_ts = "TS" in task_no_upper
    has_er = "ER" in task_no_upper

    is_surveillance = "监" in audit_type_raw
    is_first = "一阶段" in audit_type_raw or "二阶段" in audit_type_raw
    is_recert = "再认证" in audit_type_raw or "转移" in audit_type_raw

    ts_str = "☑ IATF16949:2016" if has_ts else "☐ IATF16949:2016"
    er_str = "☑ ISO9001:2015" if has_er else "☐ ISO9001:2015"
    standards_str = f"{ts_str}   {er_str}"

    first_str = "☑ 初审" if is_first else "☐ 初审"
    surv_str = "☑ 监审" if is_surveillance else "☐ 监审"
    recert_str = "☑ 再认证/转移" if is_recert else "☐ 再认证/转移"
    audit_type_str = f"{first_str}   {surv_str}   {recert_str}"

    return {
        "company_name": company_name,
        "task_no": task_no,
        "lead": lead_first,
        "address": address,
        "scope": scope,
        "has_ts": has_ts,
        "has_er": has_er,
        "is_first": is_first,
        "is_surveillance": is_surveillance,
        "is_recert": is_recert,
        "standards_str": standards_str,
        "audit_type_str": audit_type_str,
    }

def replace_in_paragraph(p, data):
    """替换占位符与切换勾选框状态，不改变原有排版布局"""
    if not p.text or not p.text.strip():
        return

    text = p.text

    # 1. 占位符直接替换
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
        "{{认证标准}}": data["standards_str"],
        "{{审核类型}}": data["audit_type_str"],
    }
    for key, val in replacements.items():
        if key in text:
            text = text.replace(key, str(val))

    # 2. 原生勾选框精准状态切换（不改变文字与布局）
    if "IATF" in text:
        sym = "☑" if data["has_ts"] else "☐"
        text = re.sub(rf"{BOX_CHARS}+\s*(IATF\s*16949)", f"{sym} \\1", text, flags=re.IGNORECASE)
    if "ISO9001" in text or "9001" in text:
        sym = "☑" if data["has_er"] else "☐"
        text = re.sub(rf"{BOX_CHARS}+\s*(ISO\s*9001)", f"{sym} \\1", text, flags=re.IGNORECASE)
    if "初审" in text:
        sym = "☑" if data["is_first"] else "☐"
        text = re.sub(rf"{BOX_CHARS}+\s*(初审)", f"{sym} \\1", text, flags=re.IGNORECASE)
    if "监审" in text:
        sym = "☑" if data["is_surveillance"] else "☐"
        text = re.sub(rf"{BOX_CHARS}+\s*(监审)", f"{sym} \\1", text, flags=re.IGNORECASE)
    if "再认证" in text:
        sym = "☑" if data["is_recert"] else "☐"
        text = re.sub(rf"{BOX_CHARS}+\s*(再认证)", f"{sym} \\1", text, flags=re.IGNORECASE)

    if text != p.text:
        p.text = text

def fill_table_safely(table, data):
    """非破坏性表格填充：绝不修改/清除单元格原本的标题和框线"""
    for row in table.rows:
        cells = row.cells
        for idx, cell in enumerate(cells):
            # 处理单元格内部现有的段落（占位符 & 勾选框）
            for p in cell.paragraphs:
                replace_in_paragraph(p, data)

            raw_text = cell.text.strip()
            clean_text = raw_text.replace(" ", "").replace("\n", "")

            # 智能匹配表格标签
            mappings = [
                (["公司名称", "客户名称"], data["company_name"]),
                (["任务号", "合同号"], data["task_no"]),
                (["审核组长", "组长"], data["lead"]),
                (["审核地址", "企业地址"], data["address"]),
                (["审核范围", "认证范围"], data["scope"]),
            ]

            for keywords, val in mappings:
                if any(kw in clean_text for kw in keywords):
                    # 情况 1：单元格内形如 "公司名称："（含冒号），且后面没值，直接在当前格补充
                    if ("：" in raw_text or ":" in raw_text) and not any(char.isalnum() for char in raw_text.split("：")[-1].split(":")[-1]):
                        prefix = raw_text.split("：")[0] if "：" in raw_text else raw_text.split(":")[0]
                        cell.text = f"{prefix}：{val}"
                    # 情况 2：单元格是纯标题（如 [审核组长] 对应右侧空格），填入右侧单元格
                    elif idx + 1 < len(cells):
                        next_cell_text = cells[idx + 1].text.strip()
                        if not next_cell_text or "{{" in next_cell_text:
                            cells[idx + 1].text = str(val)

def fill_word_template(template_bytes, data):
    doc = Document(io.BytesIO(template_bytes))

    # 1. 替换正文
    for p in doc.paragraphs:
        replace_in_paragraph(p, data)

    # 2. 安全填充表格
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
                ["company_name", "task_no", "lead", "address", "scope", "standards_str", "audit_type_str"]
            ]
            preview_df.columns = ["公司名称", "任务号", "审核组长", "审核地址", "审核范围", "认证标准", "审核类型"]
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
