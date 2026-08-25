# -*- coding: utf-8 -*-
import io
import re
import zipfile
import pandas as pd
import streamlit as st
from docx import Document

st.set_page_config(
    page_title="认证评定报告自动化生成系统", page_icon="📄", layout="wide"
)

# ==========================================
# 1. 字段清洗与条件判定函数
# ==========================================
def extract_first_person(lead_str):
    """提取审核组长字段中的第一个人名"""
    if pd.isna(lead_str) or not str(lead_str).strip():
        return "未填写"
    parts = re.split(r'[ ,，/、+&\t\n]+', str(lead_str).strip())
    return parts[0] if parts else ""

def get_column_value(row, possible_keys, default=""):
    """安全获取指定列名的值"""
    for key in possible_keys:
        for col in row.index:
            if key.lower() in str(col).lower():
                val = str(row[col]).strip()
                if val and val.lower() not in ["nan", "none", "null"]:
                    return val
    return default

def process_row_data(row, index):
    """将 Excel 的单行数据按照规则解析为填充字典"""
    task_no = get_column_value(row, ["任务号", "file number"], default=f"TASK_{index+1}")
    company_name = get_column_value(row, ["公司名称", "客户名称", "企业名称"], default="未知企业")
    lead_raw = get_column_value(row, ["审核组长", "组长"], default="")
    address = get_column_value(row, ["审核地址", "地址"], default="未填写")
    scope = get_column_value(row, ["审核范围", "认证范围", "范围"], default="未填写")
    audit_type_raw = get_column_value(row, ["审核类型"], default="")

    # 1. 提取组长第一个人名
    lead_first = extract_first_person(lead_raw)

    # 2. 依据任务号判定认证标准勾选 (TS -> IATF16949, ER -> ISO9001)
    task_no_upper = task_no.upper()
    has_ts = "TS" in task_no_upper
    has_er = "ER" in task_no_upper

    # 3. 依据审核类型字段判定勾选
    is_surveillance = "监" in audit_type_raw
    is_first = "一阶段" in audit_type_raw or "二阶段" in audit_type_raw
    is_recert = "再认证" in audit_type_raw or "转移" in audit_type_raw

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
        "is_recert": is_recert
    }


# ==========================================
# 2. Word 模板填充引擎
# ==========================================
def fill_word_template(template_bytes, data):
    """根据解析后的规则填充 Word 模板"""
    doc = Document(io.BytesIO(template_bytes))

    # 构建标准与类型的文本框表示
    ts_box = "☑ IATF16949:2016" if data['has_ts'] else "☐ IATF16949:2016"
    er_box = "☑ ISO9001:2015" if data['has_er'] else "☐ ISO9001:2015"
    standards_str = f"{ts_box}   {er_box}"

    first_box = "☑ 初审" if data['is_first'] else "☐ 初审"
    surv_box = "☑ 监审" if data['is_surveillance'] else "☐ 监审"
    recert_box = "☑ 再认证/转移" if data['is_recert'] else "☐ 再认证/转移"
    audit_type_str = f"{first_box}   {surv_box}   {recert_box}"

    # 标签替换字典
    tags = {
        "{{公司名称}}": data['company_name'],
        "{{任务号}}": data['task_no'],
        "{{审核组长}}": data['lead'],
        "{{组长}}": data['lead'],
        "{{审核地址}}": data['address'],
        "{{审核范围}}": data['scope'],
        "{{认证标准}}": standards_str,
        "{{审核类型}}": audit_type_str,
    }

    def process_paragraph(p):
        full_text = p.text
        if not full_text:
            return

        # A. 占位标签替换 {{...}}
        for tag, val in tags.items():
            if tag in full_text:
                full_text = full_text.replace(tag, str(val))

        # B. 冒号标签定位替换 (例如 "公司名称：" 后的文本)
        if re.search(r"公司名称[:：]", full_text) and data['company_name'] not in full_text:
            full_text = re.sub(r"(公司名称[:：])\s*.*", r"\1 " + str(data['company_name']), full_text)

        if re.search(r"任务号[:：]", full_text) and data['task_no'] not in full_text:
            full_text = re.sub(r"(任务号[:：])\s*.*", r"\1 " + str(data['task_no']), full_text)

        if re.search(r"(审核组长|组长)[:：]", full_text) and data['lead'] not in full_text:
            full_text = re.sub(r"((?:审核组长|组长)[:：])\s*.*", r"\1 " + str(data['lead']), full_text)

        if re.search(r"审核地址[:：]", full_text) and data['address'] not in full_text:
            full_text = re.sub(r"(审核地址[:：])\s*.*", r"\1 " + str(data['address']), full_text)

        if re.search(r"(审核范围|认证范围)[:：]", full_text) and data['scope'] not in full_text:
            full_text = re.sub(r"((?:审核范围|认证范围)[:：])\s*.*", r"\1 " + str(data['scope']), full_text)

        # C. 复选框符号自动勾选替换
        if "IATF16949" in full_text and data['has_ts']:
            full_text = re.sub(r"[☐\[ \]口]\s*IATF16949:2016", "☑ IATF16949:2016", full_text)
        if "ISO9001" in full_text and data['has_er']:
            full_text = re.sub(r"[☐\[ \]口]\s*ISO9001:2015", "☑ ISO9001:2015", full_text)

        if data['is_first']:
            full_text = re.sub(r"[☐\[ \]口]\s*初审", "☑ 初审", full_text)
        if data['is_surveillance']:
            full_text = re.sub(r"[☐\[ \]口]\s*监审", "☑ 监审", full_text)
        if data['is_recert']:
            full_text = re.sub(r"[☐\[ \]口]\s*(再认证/转移|再认证)", "☑ 再认证/转移", full_text)

        if full_text != p.text:
            p.text = full_text

    # 遍历段落与表格
    for p in doc.paragraphs:
        process_paragraph(p)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    process_paragraph(p)

    out_stream = io.BytesIO()
    doc.save(out_stream)
    out_stream.seek(0)
    return out_stream.getvalue()


# ==========================================
# 3. 批量打压 ZIP 包逻辑
# ==========================================
def generate_batch_zip(df, template_bytes):
    """遍历 Excel 每一行生成独立 Word 并打包 ZIP"""
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for idx, row in df.iterrows():
            data = process_row_data(row, idx)
            doc_bytes = fill_word_template(template_bytes, data)

            clean_company = re.sub(r'[\\/*?:"<>|]', "_", data["company_name"])
            clean_task = re.sub(r'[\\/*?:"<>|]', "_", data["task_no"])

            filename = f"{clean_company}_{clean_task}_评定报告.docx"
            zf.writestr(filename, doc_bytes)

    zip_buffer.seek(0)
    return zip_buffer.getvalue()


# ==========================================
# 4. Streamlit 界面
# ==========================================
st.title("🛡️ 认证评定报告自动化生成系统")

col1, col2 = st.columns(2)
with col1:
    excel_file = st.file_uploader(
        "1. 上传 Excel 数据文件 (.xlsx / .xls) 【必选】",
        type=["xlsx", "xls"]
    )
with col2:
    template_file = st.file_uploader(
        "2. 上传 Word 模板文件 (.docx) 【必选】",
        type=["docx"]
    )

st.markdown("---")

tab_single, tab_batch = st.tabs(["🎯 单条记录生成", "📦 生成多个报告 (Batch Generate)"])

# 强制校验：必须同时上传 Excel 和 Word 模板
if excel_file is not None and template_file is not None:
    try:
        df = pd.read_excel(excel_file)
        template_bytes = template_file.getvalue()

        # ------------------------------------------
        # Tab 1: 单条记录生成
        # ------------------------------------------
        with tab_single:
            st.subheader("🎯 选择单条记录进行单独生成")
            
            records_data = [process_row_data(row, idx) for idx, row in df.iterrows()]
            company_names = [d["company_name"] for d in records_data]

            selected_company = st.selectbox("请选择要生成报告的企业：", options=company_names)
            
            selected_idx = company_names.index(selected_company)
            single_data = records_data[selected_idx]

            st.info(f"**已选中企业**：{single_data['company_name']}（任务号: {single_data['task_no']}）")
            
            c_a, c_b = st.columns(2)
            with c_a:
                st.write(f"**提取的组长**: {single_data['lead']}")
                st.write(f"**审核地址**: {single_data['address']}")
                st.write(f"**审核范围**: {single_dict['scope'] if 'single_dict' in locals() else single_data['scope']}")
            with c_b:
                st.write(f"**TS 认证 (IATF16949)**: {'☑ 勾选' if single_data['has_ts'] else '☐ 未勾选'}")
                st.write(f"**ER 认证 (ISO9001)**: {'☑ 勾选' if single_data['has_er'] else '☐ 未勾选'}")
                audit_type_desc = []
                if single_data['is_first']: audit_type_desc.append("初审")
                if single_data['is_surveillance']: audit_type_desc.append("监审")
                if single_data['is_recert']: audit_type_desc.append("再认证/转移")
                st.write(f"**勾选的审核类型**: {', '.join(audit_type_desc) if audit_type_desc else '无'}")

            st.markdown("---")

            single_doc_bytes = fill_word_template(template_bytes, single_data)
            st.download_button(
                label=f"📄 下载【{single_data['company_name']}】 Word 报告",
                data=single_doc_bytes,
                file_name=f"{single_data['company_name']}_{single_data['task_no']}_评定报告.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

        # ------------------------------------------
        # Tab 2: 批量生成多个报告
        # ------------------------------------------
        with tab_batch:
            st.subheader("📦 批量导出所有企业的 Word 报告 (.zip)")
            st.write(f"Excel 中共需生成 **{len(df)}** 份独立报告。")

            zip_data = generate_batch_zip(df, template_bytes)
            st.download_button(
                label="📦 一键打包下载所有 Word 报告 (.zip)",
                data=zip_data,
                file_name="批量认证评定报告包.zip",
                mime="application/zip"
            )

    except Exception as e:
        st.error(f"处理文件时出错，请检查输入格式：{str(e)}")

else:
    warning_msg = "⚠️ 请在上方同时上传 **Excel 数据文件** 和 **Word 模板文件** 以后开启生成功能。"
    with tab_single:
        st.warning(warning_msg)
    with tab_batch:
        st.warning(warning_msg)
