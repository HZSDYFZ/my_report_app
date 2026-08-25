# -*- coding: utf-8 -*-
import io
import re
import zipfile
import pandas as pd
import streamlit as st
from docx import Document

# 页面基础配置
st.set_page_config(
    page_title="认证评定报告生成系统", page_icon="📄", layout="wide"
)


# ==========================================
# 1. 核心解析与清洗引擎
# ==========================================
def parse_and_fix_excel(file_buffer):
    """解析 Excel 多表并完成清洗脱敏"""
    xls = pd.ExcelFile(file_buffer)

    df1 = pd.read_excel(xls, sheet_name="Sheet1")
    df2 = (
        pd.read_excel(xls, sheet_name="Sheet2")
        if "Sheet2" in xls.sheet_names
        else pd.DataFrame()
    )
    df3 = (
        pd.read_excel(xls, sheet_name="Sheet3")
        if "Sheet3" in xls.sheet_names
        else pd.DataFrame()
    )
    df4 = (
        pd.read_excel(xls, sheet_name="Sheet4")
        if "Sheet4" in xls.sheet_names
        else pd.DataFrame()
    )

    if not df3.empty:
        df3 = df3.rename(
            columns={
                df3.columns[0]: "任务号",
                "Observations": "Sheet3_结论",
                "Date": "Sheet3_日期",
            }
        )
        df3 = df3.drop_duplicates(subset=["任务号"], keep="first")

    if not df4.empty:
        df4.columns = df4.iloc[0]
        df4 = df4[1:].reset_index(drop=True)
        df4 = df4.rename(columns={"File number(s)": "任务号"})
        df4 = df4.drop_duplicates(subset=["任务号"], keep="first")

    if not df2.empty and "任务号" in df2.columns:
        df2 = df2.drop_duplicates(subset=["任务号"], keep="first")

    master_list = []

    for idx, row in df1.iterrows():
        task_no = str(row.get("任务号", "")).strip()

        row2 = (
            df2[df2["任务号"] == task_no].iloc[0]
            if (not df2.empty and task_no in df2["任务号"].values)
            else pd.Series()
        )
        row3 = (
            df3[df3["任务号"] == task_no].iloc[0]
            if (not df3.empty and task_no in df3["任务号"].values)
            else pd.Series()
        )
        row4 = (
            df4[df4["任务号"] == task_no].iloc[0]
            if (not df4.empty and task_no in df4["任务号"].values)
            else pd.Series()
        )

        # 公司名称提取与邮箱污染修正
        s1_company = str(row.get("客户名称 Client Name", "")).strip()
        s2_company = (
            str(row2.get("企业中文名字", row2.get("企业名称", ""))).strip()
            if not row2.empty
            else ""
        )
        s4_company = (
            str(row4.get("Company name", "")).strip() if not row4.empty else ""
        )

        if (
            "@" in s1_company
            or not s1_company
            or s1_company.lower() in ["nan", "none", "null"]
        ):
            company_name = (
                s2_company
                if s2_company and s2_company.lower() != "nan"
                else (
                    s4_company
                    if s4_company and s4_company.lower() != "nan"
                    else "未知企业"
                )
            )
        else:
            company_name = s1_company

        # 英文名称
        company_en = (
            str(
                row2.get(
                    "企业英文名字",
                    s4_company if s4_company != company_name else "",
                )
            ).strip()
            if not row2.empty
            else s4_company
        )
        if company_en.lower() in ["nan", "none", "null"]:
            company_en = ""

        # 审核团队组合
        lead = str(
            row.get("审核组长", row2.get("组长", "") if not row2.empty else "")
        ).strip()
        members = (
            str(row2.get("组员", "")).strip()
            if (not row2.empty and pd.notna(row2.get("组员")))
            else ""
        )
        team_str = (
            f"{lead} (成员: {members})"
            if (members and members.lower() != "nan")
            else lead
        )

        # 审核地址清洗
        address = str(row.get("审核地址", "")).strip()
        if (
            re.match(r"^\d{4}[-/]\d{2}[-/]\d{2}", address)
            or address.lower() in ["nan", "none", "null", ""]
        ):
            real_address = (
                str(row2.get("审核地址", "")).strip() if not row2.empty else ""
            )
            address = (
                real_address if real_address.lower() != "nan" else "未填写"
            )

        # 认证范围与标准
        scope = str(row.get("认证范围", "")).strip()
        if not scope or scope.lower() in ["nan", "none", "null"]:
            scope = (
                str(row2.get("审核范围", "")).strip() if not row2.empty else ""
            )
        if scope.lower() in ["nan", "none", "null"]:
            scope = ""

        standard = (
            str(row2.get("标准", "")).strip() if not row2.empty else "ISO/IATF"
        )
        if standard.lower() in ["nan", "none", "null"]:
            standard = ""

        # 认证结论与日期
        decision = str(row.get("认证决定结论", "")).strip()
        if not decision or decision.lower() in ["nan", "none", "null"]:
            decision = (
                str(
                    row3.get("Sheet3_结论", row4.get("Observations", ""))
                ).strip()
                if not row3.empty
                else ""
            )

        date_val = str(row.get("日期", "")).strip()
        if not date_val or date_val.lower() in ["nan", "0", "none"]:
            date_val = (
                str(
                    row3.get("Sheet3_日期", row4.get("VP pass date", ""))
                ).strip()
                if not row3.empty
                else ""
            )

        master_list.append(
            {
                "序号": row.get("项目序号 No.", idx + 1),
                "合同号": (
                    str(row.get("合同号 Contract No.", "")).strip()
                    if str(row.get("合同号 Contract No.", "")).lower() != "nan"
                    else ""
                ),
                "任务号": task_no,
                "公司中文名": company_name,
                "公司英文名": company_en,
                "审核类型": (
                    str(
                        row.get(
                            "审核类型Audit Type",
                            row2.get("审核类型", "")
                            if not row2.empty
                            else "",
                        )
                    ).strip()
                ),
                "认证标准": standard,
                "审核团队": team_str,
                "评定人员": (
                    str(
                        row.get(
                            "评定人员",
                            row2.get("评定人员", "") if not row2.empty else "",
                        )
                    ).strip()
                ),
                "审核地址": address,
                "认证范围": scope,
                "认证结论": decision,
                "结论日期": (
                    date_val
                    if date_val.lower() not in ["nan", "none", "null", "0"]
                    else ""
                ),
            }
        )

    return pd.DataFrame(master_list)


# ==========================================
# 2. 报告生成引擎
# ==========================================
def fill_word_template_single(data_dict, template_bytes=None):
    """单条记录 Word 模板填充引擎"""
    if template_bytes:
        doc = Document(io.BytesIO(template_bytes))

        def replace_in_paragraphs(paragraphs, data):
            for p in paragraphs:
                for k, v in data.items():
                    tag = f"{{{{{k}}}}}"
                    if tag in p.text:
                        p.text = p.text.replace(
                            tag, str(v) if pd.notna(v) and v != "" else ""
                        )

        replace_in_paragraphs(doc.paragraphs, data_dict)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    replace_in_paragraphs(cell.paragraphs, data_dict)
    else:
        doc = Document()
        doc.add_heading(f"认证评定单项报告 - {data_dict['公司中文名']}", level=1)

        p = doc.add_paragraph()
        p.add_run("• 公司中文名：").bold = True
        p.add_run(f"{data_dict['公司中文名']}\n")

        p.add_run("• 公司英文名：").bold = True
        p.add_run(f"{data_dict['公司英文名']}\n")

        p.add_run("• 任务号：").bold = True
        p.add_run(f"{data_dict['任务号']}   |   ")
        p.add_run("合同号：").bold = True
        p.add_run(f"{data_dict['合同号']}\n")

        p.add_run("• 认证标准：").bold = True
        p.add_run(f"{data_dict['认证标准']}   |   ")
        p.add_run("审核类型：").bold = True
        p.add_run(f"{data_dict['审核类型']}\n")

        p.add_run("• 审核团队：").bold = True
        p.add_run(f"{data_dict['审核团队']}   |   ")
        p.add_run("评定人员：").bold = True
        p.add_run(f"{data_dict['评定人员']}\n")

        p.add_run("• 认证结论：").bold = True
        p.add_run(f"{data_dict['认证结论']}   |   结论日期：{data_dict['结论日期']}\n")

        p.add_run("• 审核地址：").bold = True
        p.add_run(f"{data_dict['审核地址']}\n")

        p.add_run("• 认证范围：").bold = True
        p.add_run(f"{data_dict['认证范围']}")

    target_stream = io.BytesIO()
    doc.save(target_stream)
    target_stream.seek(0)
    return target_stream.getvalue()


def generate_word_zip_batch(df, template_bytes=None):
    """批量独立 Word 报告生成并打包 ZIP"""
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for idx, row in df.iterrows():
            data_dict = row.to_dict()
            doc_bytes = fill_word_template_single(data_dict, template_bytes)

            raw_company = str(data_dict.get("公司中文名", f"企业_{idx + 1}"))
            company_name = re.sub(r'[\\/*?:"<>|]', "_", raw_company)
            task_no = re.sub(
                r'[\\/*?:"<>|]', "_", str(data_dict.get("任务号", ""))
            )

            filename = (
                f"{company_name}_{task_no}_评定报告.docx"
                if task_no
                else f"{company_name}_评定报告.docx"
            )
            zf.writestr(filename, doc_bytes)

    zip_buffer.seek(0)
    return zip_buffer.getvalue()


def generate_excel_bytes(df):
    """批量导出 Excel 表格"""
    target_stream = io.BytesIO()
    with pd.ExcelWriter(target_stream, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="认证评定解析表")
    target_stream.seek(0)
    return target_stream.getvalue()


# ==========================================
# 3. Streamlit 界面布局 (仅保留两个核心模式)
# ==========================================
st.title("🛡️ 认证评定报告生成系统")

# 文件上传区
c1, c2 = st.columns(2)
with c1:
    excel_file = st.file_uploader(
        "1. 上传认证评定 Excel 数据文件 (.xlsx / .xls)",
        type=["xlsx", "xls"],
    )
with c2:
    template_file = st.file_uploader(
        "2. (可选) 上传自定义 Word 模板 (.docx)", type=["docx"]
    )
    with st.expander("💡 占位符提示"):
        st.caption(
            "模板标签：`{{公司中文名}}` `{{公司英文名}}` `{{任务号}}` `{{合同号}}` `{{审核团队}}` `{{评定人员}}` `{{认证标准}}` `{{审核类型}}` `{{审核地址}}` `{{认证范围}}` `{{认证结论}}` `{{结论日期}}`"
        )

template_bytes = template_file.getvalue() if template_file else None

# 仅保留两个核心 Tab 模式
tab_single, tab_batch = st.tabs(
    ["🎯 单条报告生成", "📦 批量导出报告 (ZIP)"]
)

if excel_file is not None:
    try:
        df_master = parse_and_fix_excel(excel_file)

        # 侧边栏搜索与过滤
        st.sidebar.header("🔍 数据检索与筛选")
        search_kw = st.sidebar.text_input("搜索企业名称/任务号:")
        selected_decision = st.sidebar.multiselect(
            "按认证结论筛选:",
            options=df_master["认证结论"].unique().tolist(),
            default=[],
        )
        selected_standard = st.sidebar.multiselect(
            "按认证标准筛选:",
            options=df_master["认证标准"].unique().tolist(),
            default=[],
        )

        filtered_df = df_master.copy()
        if search_kw:
            filtered_df = filtered_df[
                filtered_df["公司中文名"].str.contains(search_kw, na=False)
                | filtered_df["任务号"].str.contains(search_kw, na=False)
                | filtered_df["公司英文名"].str.contains(search_kw, na=False)
            ]
        if selected_decision:
            filtered_df = filtered_df[
                filtered_df["认证结论"].isin(selected_decision)
            ]
        if selected_standard:
            filtered_df = filtered_df[
                filtered_df["认证标准"].isin(selected_standard)
            ]

        # ------------------------------------
        # 模式一：单条报告生成
        # ------------------------------------
        with tab_single:
            st.subheader("🎯 选定企业生成单项报告")
            company_list = filtered_df["公司中文名"].tolist()

            if company_list:
                selected_company = st.selectbox(
                    "选择目标企业:", options=company_list
                )
                single_dict = filtered_df[
                    filtered_df["公司中文名"] == selected_company
                ].iloc[0].to_dict()

                st.info(
                    f"**已选目标**：{single_dict['公司中文名']}（任务号: {single_dict['任务号']}）"
                )

                col_a, col_b = st.columns(2)
                with col_a:
                    st.write(f"**公司英文名**: {single_dict['公司英文名']}")
                    st.write(f"**合同号**: {single_dict['合同号']}")
                    st.write(f"**审核团队**: {single_dict['审核团队']}")
                    st.write(f"**评定人员**: {single_dict['评定人员']}")
                    st.write(f"**认证标准**: {single_dict['认证标准']}")
                with col_b:
                    st.write(f"**审核类型**: {single_dict['审核类型']}")
                    st.write(f"**认证结论**: {single_dict['认证结论']}")
                    st.write(f"**结论日期**: {single_dict['结论日期']}")
                    st.write(f"**审核地址**: {single_dict['审核地址']}")
                    st.write(f"**认证范围**: {single_dict['认证范围']}")

                st.markdown("---")

                s_word_bytes = fill_word_template_single(
                    single_dict, template_bytes
                )
                st.download_button(
                    label="📄 下载该企业 Word 报告",
                    data=s_word_bytes,
                    file_name=f"{single_dict['公司中文名']}_评定报告.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )
            else:
                st.warning("无匹配筛选条件的记录。")

        # ------------------------------------
        # 模式二：批量报告导出 (ZIP)
        # ------------------------------------
        with tab_batch:
            st.subheader("📦 批量导出每个企业的独立 Word 报告")
            st.write(f"当前筛选记录数：**{len(filtered_df)}** 条")

            btn_col1, btn_col2 = st.columns(2)

            # 导出 ZIP 包
            zip_bytes = generate_word_zip_batch(filtered_df, template_bytes)
            btn_col1.download_button(
                label="📦 批量下载所有企业 Word 报告压缩包 (.zip)",
                data=zip_bytes,
                file_name="批量认证评定报告包.zip",
                mime="application/zip",
            )

            # 导出 Excel 汇总
            excel_bytes = generate_excel_bytes(filtered_df)
            btn_col2.download_button(
                label="📊 导出当前筛选数据 Excel (.xlsx)",
                data=excel_bytes,
                file_name="认证评定记录_筛选汇总.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    except Exception as e:
        st.error(f"解析文件出错: {str(e)}")
else:
    with tab_single:
        st.info("👈 请在上方上传 Excel 文件以开启操作。")
    with tab_batch:
        st.info("👈 请在上方上传 Excel 文件以开启操作。")
