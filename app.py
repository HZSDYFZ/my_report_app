# -*- coding: utf-8 -*-
import io
import re
import zipfile
import pandas as pd
import streamlit as st
from docx import Document

# 页面基本配置
st.set_page_config(
    page_title="认证评定报告自动化生成系统", page_icon="📄", layout="wide"
)

BOX_CHARS = r"[□☐☑✔\[\]口]"

# ==========================================
# 1. 数据解析与严格空行过滤
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


def get_clean_col_val(row, possible_keys, default=""):
    """多列名模糊匹配"""
    for key in possible_keys:
        for col in row.index:
            col_clean = str(col).replace(" ", "").replace("\n", "").lower()
            key_clean = key.replace(" ", "").lower()
            if key_clean in col_clean:
                val = str(row[col]).strip()
                if val and val.lower() not in [
                    "nan",
                    "none",
                    "null",
                    "nat",
                    "0",
                    "undefined",
                ]:
                    return val
    return default


def is_real_data_row(row):
    """严格数据有效性校验，彻底过滤 Excel 尾部空行/格式残留行"""
    if row.dropna().empty:
        return False

    comp = get_clean_col_val(
        row,
        ["公司名称", "客户名称", "企业名称", "client name", "公司"],
        default="",
    )
    task = get_clean_col_val(
        row, ["任务号", "file number", "合同号", "项目编号"], default=""
    )
    lead = get_clean_col_val(
        row,
        [
            "审核组长",
            "组长",
            "lead",
            "auditor",
            "审核员",
            "审核团队",
            "团队",
        ],
        default="",
    )

    invalid_words = ["", "nan", "none", "null", "未知企业", "未填写", "0"]
    if (
        comp.lower() in invalid_words
        and task.lower() in invalid_words
        and lead.lower() in invalid_words
    ):
        return False

    if any(
        kw in comp for kw in ["合计", "小计", "统计", "说明", "备注", "填表说明"]
    ):
        return False

    return True


def process_row_data(row, index):
    """解析 Excel 单行数据"""
    company_name = get_clean_col_val(
        row,
        ["公司名称", "客户名称", "企业名称", "client name", "公司"],
        default="未知企业",
    )
    task_no = get_clean_col_val(
        row,
        ["任务号", "file number", "合同号", "项目编号"],
        default=f"TASK_{index+1}",
    )
    lead_raw = get_clean_col_val(
        row,
        [
            "审核组长",
            "组长",
            "lead",
            "auditor",
            "审核员",
            "审核团队",
            "团队",
            "组长姓名",
        ],
        default="",
    )
    address = get_clean_col_val(
        row, ["审核地址", "地址", "address", "企业地址"], default="未填写"
    )
    scope = get_clean_col_val(
        row, ["审核范围", "认证范围", "范围", "scope"], default="未填写"
    )
    audit_type_raw = get_clean_col_val(
        row, ["审核类型", "audit type"], default=""
    )

    lead_first = extract_first_person(lead_raw)
    if not lead_first:
        lead_first = "未填写"

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
# 2. Word 表格与段落替换引擎
# ==========================================
def replace_checkbox(text, keyword, should_check):
    target_symbol = "☑" if should_check else "☐"

    if keyword == "IATF16949":
        pattern = rf"{BOX_CHARS}+\s*(IATF\s*16949(:2016)?)"
    elif keyword == "ISO9001":
        pattern = rf"{BOX_CHARS}+\s*(ISO\s*9001(:2015)?)"
    elif keyword == "再认证":
        pattern = rf"{BOX_CHARS}+\s*(再认证(/转移)?)"
    else:
        pattern = rf"{BOX_CHARS}+\s*({re.escape(keyword)})"

    if re.search(pattern, text, re.IGNORECASE):
        return re.sub(
            pattern, rf"{target_symbol} \1", text, flags=re.IGNORECASE
        )
    return text


def process_paragraph(p, data):
    full_text = p.text
    if not full_text or not full_text.strip():
        return

    original_text = full_text

    tags = {
        "{{公司名称}}": data["company_name"],
        "{{任务号}}": data["task_no"],
        "{{审核组长}}": data["lead"],
        "{{组长}}": data["lead"],
        "{{审核地址}}": data["address"],
        "{{审核范围}}": data["scope"],
        "{{认证标准}}": data["standards_str"],
        "{{审核类型}}": data["audit_type_str"],
    }
    for tag, val in tags.items():
        if tag in full_text:
            full_text = full_text.replace(tag, str(val))

    full_text = replace_checkbox(full_text, "IATF16949", data["has_ts"])
    full_text = replace_checkbox(full_text, "ISO9001", data["has_er"])
    full_text = replace_checkbox(full_text, "初审", data["is_first"])
    full_text = replace_checkbox(full_text, "监审", data["is_surveillance"])
    full_text = replace_checkbox(full_text, "再认证", data["is_recert"])

    if full_text != original_text:
        p.text = full_text


def process_table_cells(table, data):
    """精准表格替换：确保左列保留标签名称，右列填入提取的数据内容"""
    for r_idx, row in enumerate(table.rows):
        cells = row.cells
        if len(cells) < 2 or cells[0] is cells[1]:
            continue

        c0, c1 = cells[0], cells[1]
        c0_text = (
            c0.text.strip().replace(" ", "").replace("：", "").replace(":", "")
        )
        c1_text = c1.text.strip()

        # 1. 认证标准
        if (
            "认证标准" in c0_text
            or "IATF16949" in c1_text
            or "ISO9001" in c1_text
        ):
            c0.text = "认证标准"
            std_str = (
                "☑ IATF16949:2016"
                if data.get("has_ts")
                else "☐ IATF16949:2016"
            )
            std_str += (
                "    ☑ ISO9001:2015"
                if data.get("has_er")
                else "    ☐ ISO9001:2015"
            )
            std_str += (
                "    ☐ ISO14001:2015\n☐ ISO 45001:2018  ☐ 其他:"
            )
            c1.text = std_str
            continue

        # 2. 审核类型
        if "审核类型" in c0_text or "初审" in c1_text or "监审" in c1_text:
            c0.text = "审核类型"
            type_str = "☑ 初审" if data.get("is_first") else "☐ 初审"
            type_str += (
                "    ☑ 监审"
                if data.get("is_surveillance")
                else "    ☐ 监审"
            )
            type_str += (
                "    ☑ 再认证/转移"
                if data.get("is_recert")
                else "    ☐ 再认证/转移"
            )
            type_str += "    ☐ 特殊审核    ☐ 其它"
            c1.text = type_str
            continue

        # 3. 审核组长
        if (
            "组长" in c0_text
            or "审核员" in c0_text
            or "{{审核组长}}" in c0.text
            or "{{组长}}" in c0.text
            or r_idx == 1
        ):
            if not any(
                k in c0_text for k in ["公司", "标准", "类型", "地址", "范围"]
            ):
                c0.text = "审核组长"
                if data.get("lead") and data["lead"] != "未填写":
                    c1.text = str(data["lead"])
                continue

        # 4. 审核地址
        if (
            "地址" in c0_text
            or "厂址" in c0_text
            or "{{审核地址}}" in c0.text
            or (r_idx == 4 and "范围" not in c0_text)
        ):
            c0.text = "审核地址"
            if data.get("address") and data["address"] != "未填写":
                c1.text = str(data["address"])
            continue

        # 5. 审核范围
        if "范围" in c0_text or "{{审核范围}}" in c0.text or r_idx == 5:
            c0.text = "审核范围"
            if data.get("scope") and data["scope"] != "未填写":
                c1.text = str(data["scope"])
            continue

        # 6. 第一行公司名称与任务号
        if (
            "公司" in c0_text
            or "客户" in c0_text
            or "{{公司名称}}" in c0.text
            or r_idx == 0
        ):
            if "{{公司名称}}" in c0.text or c0_text == "":
                c0.text = str(data.get("company_name", ""))
            elif "公司" in c0_text or "客户" in c0_text:
                c1.text = str(data.get("company_name", ""))

            if "{{任务号}}" in c1.text or c1_text == "":
                c1.text = str(data.get("task_no", ""))


def fill_word_template(template_bytes, data):
    doc = Document(io.BytesIO(template_bytes))

    for table in doc.tables:
        process_table_cells(table, data)

    for p in doc.paragraphs:
        process_paragraph(p, data)

    out_stream = io.BytesIO()
    doc.save(out_stream)
    out_stream.seek(0)
    return out_stream.getvalue()


# ==========================================
# 3. Streamlit 网页前端界面
# ==========================================
st.title("🛡️ 认证评定报告全量自动化生成系统")

c1, c2 = st.columns(2)
with c1:
    excel_file = st.file_uploader(
        "1. 上传认证 Excel 数据文件 (.xlsx / .xls) 【必选】",
        type=["xlsx", "xls"],
    )
with c2:
    template_file = st.file_uploader(
        "2. 上传 Word 报告模板 (.docx) 【必选】", type=["docx"]
    )

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
            st.warning("⚠️ Excel 文件中未读取到任何有效数据行，请检查表格内容。")
        else:
            st.subheader(
                f"📋 待处理有效数据预览（已过滤空行，共 {len(parsed_records)} 条记录）"
            )

            preview_data = []
            for idx, r in enumerate(parsed_records):
                preview_data.append(
                    {
                        "序号": idx + 1,
                        "公司名称": r["company_name"],
                        "任务号": r["task_no"],
                        "提取的审核组长": r["lead"],
                        "TS(IATF16949)": "☑" if r["has_ts"] else "☐",
                        "ER(ISO9001)": "☑" if r["has_er"] else "☐",
                        "审核类型": r["audit_type_str"],
                        "审核地址": r["address"],
                    }
                )
            st.dataframe(pd.DataFrame(preview_data), use_container_width=True)

            st.markdown("### 🚀 开始生成")

            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                for idx, data in enumerate(parsed_records):
                    doc_bytes = fill_word_template(template_bytes, data)
                    clean_company = re.sub(
                        r'[\\/*?:"<>|]', "_", data["company_name"]
                    )
                    clean_task = re.sub(
                        r'[\\/*?:"<>|]', "_", data["task_no"]
                    )
                    filename = f"{clean_company}_{clean_task}_评定报告.docx"
                    zf.writestr(filename, doc_bytes)

            zip_buffer.seek(0)

            st.download_button(
                label=f"📦 一键下载所有有效 Word 报告压缩包 ({len(parsed_records)} 份 .zip)",
                data=zip_buffer.getvalue(),
                file_name="全量认证评定报告包.zip",
                mime="application/zip",
                type="primary",
                use_container_width=True,
            )

    except Exception as e:
        st.error(f"处理数据或模板时出错: {str(e)}")
else:
    st.info(
        "👈 请在上方同时上传 **Excel 数据文件** 和 **Word 模板文件** 即可一键生成全量报告。"
    )
