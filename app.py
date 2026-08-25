# -*- coding: utf-8 -*-
import io
import re
import zipfile
import pandas as pd
import streamlit as st
from docx import Document

# 1. Streamlit 页面配置（必须置于首行，防止页面渲染问题）
st.set_page_config(
    page_title="认证评定报告自动化生成系统",
    page_icon="📄",
    layout="wide"
)

BOX_CHARS = r"[□☐☑✔\[\]口]"

# ==========================================
# 1. 数据解析与严密清洗
# ==========================================
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
    """过滤 Excel 尾部空白/公式残留行"""
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
    """解析单行数据"""
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

# ==========================================
# 2. Word 填充引擎（支持全域替换与强行填入）
# ==========================================
def replace_text_in_paragraph(p, data_dict):
    """对段落文本进行占位符和复选框替换"""
    text = p.text
    if not text or not text.strip():
        return

    original_text = text

    # 常见占位符字典映射
    replacements = {
        "{{公司名称}}": data_dict["company_name"],
        "{{客户名称}}": data_dict["company_name"],
        "{{任务号}}": data_dict["task_no"],
        "{{审核组长}}": data_dict["lead"],
        "{{组长}}": data_dict["lead"],
        "{{审核地址}}": data_dict["address"],
        "{{地址}}": data_dict["address"],
        "{{审核范围}}": data_dict["scope"],
        "{{范围}}": data_dict["scope"],
        "{{认证标准}}": data_dict["standards_str"],
        "{{审核类型}}": data_dict["audit_type_str"],
    }

    for key, val in replacements.items():
        if key in text:
            text = text.replace(key, str(val))

    # 复选框正则替换
    def fix_box(t, kw, checked):
        target = "☑" if checked else "☐"
        patt = rf"{BOX_CHARS}+\s*({re.escape(kw)})"
        return re.sub(patt, rf"{target} \1", t, flags=re.IGNORECASE)

    text = fix_box(text, "IATF16949", data_dict["has_ts"])
    text = fix_box(text, "ISO9001", data_dict["has_er"])
    text = fix_box(text, "初审", data_dict["is_first"])
    text = fix_box(text, "监审", data_dict["is_surveillance"])
    text = fix_box(text, "再认证", data_dict["is_recert"])

    if text != original_text:
        p.text = text


def process_document_tables(doc, data):
    """表格识别：同时支持 {{占位符}} 替换与双列键值对强行注入"""
    for table in doc.tables:
        for row in table.rows:
            cells = row.cells

            # 1. 优先替换单元格内所有的 {{占位符}}
            for cell in cells:
                for p in cell.paragraphs:
                    replace_text_in_paragraph(p, data)

            # 2. 针对双列及以上表格，根据左侧标题行强制写入右侧内容
            if len(cells) >= 2:
                c0_text = cells[0].text.strip().replace(" ", "").replace("：", "").replace(":", "")

                # 公司名称
                if "公司" in c0_text or "客户" in c0_text:
                    if len(cells) > 1 and ("{{" in cells[1].text or not cells[1].text.strip()):
                        cells[1].text = str(data["company_name"])

                # 审核组长
                elif "组长" in c0_text or "审核员" in c0_text:
                    cells[0].text = "审核组长"
                    cells[1].text = str(data["lead"])

                # 审核地址
                elif "地址" in c0_text or "厂址" in c0_text:
                    cells[0].text = "审核地址"
                    cells[1].text = str(data["address"])

                # 审核范围
                elif "范围" in c0_text:
                    cells[0].text = "审核范围"
                    cells[1].text = str(data["scope"])

                # 认证标准
                elif "标准" in c0_text or "认证标准" in c0_text:
                    cells[0].text = "认证标准"
                    cells[1].text = (
                        f"☑ IATF16949:2016" if data["has_ts"] else "☐ IATF16949:2016"
                    ) + "    " + (
                        f"☑ ISO9001:2015" if data["has_er"] else "☐ ISO9001:2015"
                    ) + "\n☐ ISO14001:2015  ☐ ISO45001:2018  ☐ 其他"

                # 审核类型
                elif "类型" in c0_text or "审核类型" in c0_text:
                    cells[0].text = "审核类型"
                    cells[1].text = (
                        f"☑ 初审" if data["is_first"] else "☐ 初审"
                    ) + "    " + (
                        f"☑ 监审" if data["is_surveillance"] else "☐ 监审"
                    ) + "    " + (
                        f"☑ 再认证/转移" if data["is_recert"] else "☐ 再认证/转移"
                    ) + "    ☐ 特殊审核"


def fill_word_template(template_bytes, data):
    doc = Document(io.BytesIO(template_bytes))

    # 1. 遍历并替换所有正文段落
    for p in doc.paragraphs:
        replace_text_in_paragraph(p, data)

    # 2. 遍历并替换所有表格段落与结构
    process_document_tables(doc, data)

    out_stream = io.BytesIO()
    doc.save(out_stream)
    out_stream.seek(0)
    return out_stream.getvalue()


# ==========================================
# 3. Streamlit 网页交互界面
# ==========================================
st.title("🛡️ 认证评定报告全量自动化生成系统")

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
                record = process_row_data(row, idx)
                parsed_records.append(record)

        if not parsed_records:
            st.warning("⚠️ Excel 文件中未读取到有效数据，请检查 Excel 表头列名（应包含：公司名称、任务号、审核组长、地址、范围等）。")
        else:
            st.subheader(f"📋 提取到的有效数据预览（共 {len(parsed_records)} 条记录）")
            st.caption("提示：请核对下方表格中的数据是否已准确提取自 Excel。")

            preview_df = pd.DataFrame(parsed_records)[
                ["company_name", "task_no", "lead", "address", "scope", "standards_str", "audit_type_str"]
            ]
            preview_df.columns = ["公司名称", "任务号", "审核组长", "审核地址", "审核范围", "认证标准", "审核类型"]
            st.dataframe(preview_df, use_container_width=True)

            st.markdown("### 🚀 开始批量生成 Word 报告")

            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                for idx, data in enumerate(parsed_records):
                    doc_bytes = fill_word_template(template_bytes, data)
                    clean_company = re.sub(r'[\\/*?:"<>|]', "_", str(data["company_name"]))
                    clean_task = re.sub(r'[\\/*?:"<>|]', "_", str(data["task_no"]))
                    filename = f"{clean_company}_{clean_task}_评定报告.docx"
                    zf.writestr(filename, doc_bytes)

            zip_buffer.seek(0)

            st.download_button(
                label=f"📦 一键下载所有 Word 报告压缩包 ({len(parsed_records)} 份 .zip)",
                data=zip_buffer.getvalue(),
                file_name="全量认证评定报告包.zip",
                mime="application/zip",
                type="primary",
                use_container_width=True
            )

    except Exception as e:
        st.error(f"处理文件时发生错误: {str(e)}")
        st.exception(e)
else:
    st.info("👈 请在上方上传 **Excel 数据文件** 和 **Word 模板文件**。")
