import os
import re
import docx
from docx import Document


def process_audit_doc_table(table, data):
    """智能填充审核文档表格：

    1. 第0列严格保留/恢复标签（审核组长、审核地址、审核范围等） 2. 第1列填入具体内容（姓名、地址文本、范围文本） 3. 自动勾选认证标准与审核类型
    """
    for r_idx, row in enumerate(table.rows):
        cells = row.cells
        if len(cells) < 2:
            continue

        c0 = cells[0]
        c1 = cells[1]
        c0_text = (
            c0.text.strip().replace(" ", "").replace("：", "").replace(":", "")
        )
        c1_text = c1.text.strip()

        # 1. 认证标准 (勾选框)
        if (
            "认证标准" in c0_text
            or "IATF16949" in c1_text
            or "ISO9001" in c1_text
        ):
            c0.text = "认证标准"
            std_str = ""
            std_str += (
                "☑ IATF16949:2016"
                if data.get("has_ts")
                else "☐ IATF16949:2016"
            )
            std_str += (
                "    ☑ ISO9001:2015"
                if data.get("has_er")
                else "    ☐ ISO9001:2015"
            )
            std_str += "    ☐ ISO14001:2015\n☐ ISO 45001:2018  ☐ 其他:"
            c1.text = std_str
            continue

        # 2. 审核类型 (勾选框)
        if "审核类型" in c0_text or "初审" in c1_text or "监审" in c1_text:
            c0.text = "审核类型"
            type_str = ""
            type_str += "☑ 初审" if data.get("is_first") else "☐ 初审"
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

        # 3. 审核组长 (第0列为标签，第1列填姓名)
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

        # 4. 审核地址 (第0列为标签，第1列填地址)
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

        # 5. 审核范围 (第0列为标签，第1列填范围)
        if "范围" in c0_text or "{{审核范围}}" in c0.text or r_idx == 5:
            c0.text = "审核范围"
            if data.get("scope") and data["scope"] != "未填写":
                c1.text = str(data["scope"])
            continue

        # 6. 第一行公司名称与任务号
        if "公司" in c0_text or "客户" in c0_text or "{{公司名称}}" in c0.text or r_idx == 0:
            if "{{公司名称}}" in c0.text or c0_text == "":
                c0.text = str(data.get("company_name", ""))
            elif "公司" in c0_text or "客户" in c0_text:
                c1.text = str(data.get("company_name", ""))

            if "{{任务号}}" in c1.text or c1_text == "":
                c1.text = str(data.get("task_no", ""))


def fill_word_template(template_path, output_path, data):
    """读取Word模板，精准填充数据并保存"""
    doc = Document(template_path)

    # 优先处理文档中的所有表格
    for table in doc.tables:
        process_audit_doc_table(table, data)

    # 保存文件
    doc.save(output_path)
    print(f"✅ 处理完成，文件已保存至: {output_path}")


# -------------------- 示例调用 --------------------
if __name__ == "__main__":
    sample_data = {
        "company_name": "无锡永联管路密封紧固技术有限公司",
        "task_no": "2021/1585/TS/01",
        "lead": "姜强",
        "address": (
            "中国江苏省无锡市惠山区藕杨路 157 号 D 栋一层 5 号门、二层西侧"
        ),
        "scope": "非标紧固件的生产",
        "has_ts": True,  # IATF16949:2016
        "has_er": False,  # ISO9001:2015
        "is_first": False,  # 初审
        "is_surveillance": True,  # 监审
        "is_recert": False,  # 再认证
    }

    # 执行替换（请替换为您本地的模板路径与输出路径）
    # fill_word_template("input_template.docx", "output_result.docx", sample_data)
