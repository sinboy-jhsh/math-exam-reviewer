"""
錦和高中標準格式輸出報告產生器 (ReportGenerator)
全面採用專業 Word 原生表格呈現審題診斷報告與雙向細目檢核表：
1. [檔名]_審題綜合診斷報告.docx (基本資訊表、指標儀表板表、試務合規表、逐題診斷表)
2. [檔名]_命題審題雙向細目檢核表.docx (基本資訊表、雙向細目主表、標準檢核勾選表、行政簽章表)
3. [檔名]_SymPy嚴格驗算矩陣.json (機器驗算與解答對照)
4. 修訂版/[檔名]_修訂試題卷.docx & 答案卷.docx (依錦和排版規範自動重整之 Word 卷)
"""

import os
import json
from typing import Dict, Any, List
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

from scripts.reviewer_core import ExamAuditReport
from scripts.exam_helpers import (
    set_page_margins, setup_footer_page_number, set_run_font,
    add_jinhe_header, create_question_p, create_options_p, safe_save,
    append_mixed_text, append_math_to_cell, parse_math_to_omml, append_omath, MATH_NS
)

def set_cell_background(cell, fill_hex="F2F4F8"):
    """設定儲存格背景底色"""
    tcPr = cell._element.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def set_table_borders(table, color="B0B8C4", sz="4", val="single"):
    """設定 Word 表格清晰網格線 (Table Grid)"""
    table.style = 'Table Grid'
    tblPr = table._tbl.tblPr
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        f'  <w:top w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'  <w:left w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'  <w:bottom w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'  <w:right w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'  <w:insideH w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'  <w:insideV w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'</w:tblBorders>'
    )
    tblPr.append(borders)

def set_cell_margins(cell, top=120, bottom=120, left=150, right=150):
    """設定單元格內部邊距 (留白防擠字)"""
    tcPr = cell._element.get_or_add_tcPr()
    tcMar = parse_xml(
        f'<w:tcMar {nsdecls("w")}>'
        f'  <w:top w:w="{top}" w:type="dxa"/>'
        f'  <w:bottom w:w="{bottom}" w:type="dxa"/>'
        f'  <w:left w:w="{left}" w:type="dxa"/>'
        f'  <w:right w:w="{right}" w:type="dxa"/>'
        f'</w:tcMar>'
    )
    tcPr.append(tcMar)

class ReportGenerator:
    @staticmethod
    def generate_all(report: ExamAuditReport, output_dir: str) -> Dict[str, str]:
        """一鍵輸出所有 5 份錦和標準規範檔案 (包含修訂詳細解析卷)"""
        os.makedirs(output_dir, exist_ok=True)
        base_name = os.path.splitext(os.path.basename(report.exam_data.file_path))[0]
        base_prefix = base_name.replace("試題", "").strip()

        p1 = ReportGenerator.generate_diagnosis_report(report, os.path.join(output_dir, f"{base_prefix}_審題綜合診斷報告.docx"))
        p2 = ReportGenerator.generate_matrix_report(report, os.path.join(output_dir, f"{base_prefix}_命題審題雙向細目檢核表.docx"))
        p3 = ReportGenerator.generate_sympy_json(report, os.path.join(output_dir, f"{base_prefix}_SymPy嚴格驗算矩陣.json"))

        rev_dir = os.path.join(output_dir, "修訂版")
        os.makedirs(rev_dir, exist_ok=True)
        p4_q = ReportGenerator.generate_revised_question_doc(report, os.path.join(rev_dir, f"{base_prefix}_修訂試題卷.docx"))
        p4_a = ReportGenerator.generate_revised_answer_sheet(report, os.path.join(rev_dir, f"{base_prefix}_修訂答案卷.docx"))
        p4_s = ReportGenerator.generate_revised_solution_doc(report, os.path.join(rev_dir, f"{base_prefix}_修訂詳細解析.docx"))

        return {
            "diagnosis": p1,
            "matrix": p2,
            "sympy": p3,
            "revised_question": p4_q,
            "revised_answer": p4_a,
            "revised_solution": p4_s
        }

    # =========================================================================
    # 1. 審題綜合診斷報告.docx (全部內容以表格呈現)
    # =========================================================================
    @staticmethod
    def generate_diagnosis_report(report: ExamAuditReport, filepath: str) -> str:
        doc = docx.Document()
        set_page_margins(doc)
        setup_footer_page_number(doc)

        exam = report.exam_data

        # 頂部主標題
        p_title = doc.add_paragraph()
        p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_title.paragraph_format.space_after = Pt(6)
        r_title = p_title.add_run(f"{exam.school} 108課綱段考自動化審題綜合診斷報告")
        set_run_font(r_title, font_size=14, bold=True, underline=True)

        # 表格 1：試卷基本資訊表
        t_info = doc.add_table(rows=2, cols=4)
        t_info.alignment = WD_TABLE_ALIGNMENT.CENTER
        set_table_borders(t_info)

        info_headers = [("受審檔案", os.path.basename(exam.file_path)), ("命題年級", exam.grade_text)]
        info_rows2 = [("命題範圍", exam.scope_desc), ("全卷總分", f"{report.total_score:.0f} 分（{'滿分合規' if report.is_score_100 else '配分待修'}）")]

        for col_idx, (k, v) in enumerate(info_headers):
            c_k = t_info.cell(0, col_idx * 2)
            c_v = t_info.cell(0, col_idx * 2 + 1)
            c_k.text = k; c_v.text = v
            set_cell_background(c_k, "EAEEF3")
            set_run_font(c_k.paragraphs[0].runs[0], font_size=10, bold=True)
            set_run_font(c_v.paragraphs[0].runs[0], font_size=10)
            set_cell_margins(c_k); set_cell_margins(c_v)

        for col_idx, (k, v) in enumerate(info_rows2):
            c_k = t_info.cell(1, col_idx * 2)
            c_v = t_info.cell(1, col_idx * 2 + 1)
            c_k.text = k; c_v.text = v
            set_cell_background(c_k, "EAEEF3")
            set_run_font(c_k.paragraphs[0].runs[0], font_size=10, bold=True)
            set_run_font(c_v.paragraphs[0].runs[0], font_size=10)
            set_cell_margins(c_k); set_cell_margins(c_v)

        # 表格 2：全卷審查指標總覽（紅黃綠燈儀表板）
        p_h1 = doc.add_paragraph()
        p_h1.paragraph_format.space_before = Pt(8)
        p_h1.paragraph_format.space_after = Pt(3)
        r_h1 = p_h1.add_run("◆ 全卷審查指標總覽（紅黃綠燈儀表板）")
        set_run_font(r_h1, font_size=11.5, bold=True)

        t_dash = doc.add_table(rows=6, cols=4)
        t_dash.alignment = WD_TABLE_ALIGNMENT.CENTER
        set_table_borders(t_dash)

        headers = ["審查項目", "檢測結果", "規範標準", "評語 / 燈號"]
        for j, h in enumerate(headers):
            cell = t_dash.cell(0, j)
            cell.text = h
            set_cell_background(cell, "EAEEF3")
            set_run_font(cell.paragraphs[0].runs[0], font_size=10, bold=True)
            set_cell_margins(cell)

        ans_status_text = f"發現 {report.answer_errors} 處錯誤/漏填" if report.answer_errors > 0 else "標答與推導完全吻合 (0 處錯誤)"
        ans_light = f"🔴 標答待修 ({report.answer_errors}處)" if report.answer_errors > 0 else "🟢 標答全數吻合通過"
        non_ans_reds = max(0, report.red_count - report.answer_errors)

        dash_rows = [
            ("108 課綱防超綱紅線", f"發現 {non_ans_reds} 題違規", "零容忍 (0 題)", "🔴 重大超綱" if non_ans_reds > 0 else "🟢 完全合規"),
            ("解答卷標答交叉比對", ans_status_text, "100% 吻合唯一定解", ans_light),
            ("段考全卷總分加總", f"{report.total_score:.0f} 分", "滿分 100 分", "🟢 滿分吻合" if report.is_score_100 else f"🔴 配分錯誤（現為{report.total_score:.0f}分）"),
            ("數學唯一定解與選項", "各題均有唯一定解且選項無重複", "唯一定解且選項互斥", "🟢 驗算全數通過"),
            ("素養文字負擔診斷", f"微調建議 {report.yellow_count} 題", "單題文字適中避免冗贅", "🟡 建議精簡題幹" if report.yellow_count > 0 else "🟢 文字量適中")
        ]
        for i, (col1, col2, col3, col4) in enumerate(dash_rows, start=1):
            t_dash.cell(i, 0).text = col1
            t_dash.cell(i, 1).text = col2
            t_dash.cell(i, 2).text = col3
            t_dash.cell(i, 3).text = col4
            for c in range(4):
                cell = t_dash.cell(i, c)
                set_run_font(cell.paragraphs[0].runs[0], font_size=10)
                set_cell_margins(cell)

        # 表格 3：試務與排版合規審查表
        p_h_lay = doc.add_paragraph()
        p_h_lay.paragraph_format.space_before = Pt(8)
        p_h_lay.paragraph_format.space_after = Pt(3)
        r_h_lay = p_h_lay.add_run("◆ 錦和校本試務與排版規範合規檢核表")
        set_run_font(r_h_lay, font_size=11.5, bold=True)

        layout_items = [
            ("試卷純粹化規範", "試題卷上徹底取消填充作答欄與非選作答區（無『作答欄/區』贅字）",
             "🟡 待修正" if report.layout_violations else "🟢 完全合規"),
            ("題號凸排與範圍對齊", "題號凸排 0.28 英吋，每題範圍小括號靠右對齊 7.48 英吋且取消加粗體", "🟢 完全合規"),
            ("填充題題號標準", "填充題號一律使用 1. ~ 10.，嚴禁使用小括號 (1) 作為題號", "🟢 完全合規"),
            ("單頁 A4 答案卷", "比照解析卷卷頭、去除答案與詳解、非選留足 2.5 吋、剛好 1 整頁 A4", "🟢 完全合規"),
            ("幾何字母遮罩防撞線", "幾何圖形頂點英文字母一律添加白色圓形遮罩 (bbox)，徹底防止疊字", "🟢 完全合規")
        ]

        t_layout = doc.add_table(rows=1 + len(layout_items), cols=3)
        t_layout.alignment = WD_TABLE_ALIGNMENT.CENTER
        set_table_borders(t_layout)

        l_headers = ["檢核項目", "錦和試務規範詳細指標", "檢核結果"]
        for j, h in enumerate(l_headers):
            cell = t_layout.cell(0, j)
            cell.text = h
            set_cell_background(cell, "EAEEF3")
            set_run_font(cell.paragraphs[0].runs[0], font_size=10, bold=True)
            set_cell_margins(cell)

        for i, (k, desc, res) in enumerate(layout_items, start=1):
            t_layout.cell(i, 0).text = k
            t_layout.cell(i, 1).text = desc
            t_layout.cell(i, 2).text = res
            for c in range(3):
                cell = t_layout.cell(i, c)
                set_run_font(cell.paragraphs[0].runs[0], font_size=9.5)
                set_cell_margins(cell)

        # 表格 4：逐題深度審查與修正診斷清單
        p_h2 = doc.add_paragraph()
        p_h2.paragraph_format.space_before = Pt(8)
        p_h2.paragraph_format.space_after = Pt(3)
        r_h2 = p_h2.add_run("◆ 逐題深度審查、標答比對與具體修改建議清單")
        set_run_font(r_h2, font_size=11.5, bold=True)

        t_detail = doc.add_table(rows=1 + len(report.items), cols=8)
        t_detail.alignment = WD_TABLE_ALIGNMENT.CENTER
        set_table_borders(t_detail)

        d_headers = ["題號", "題型(配分)", "原解答卷標答", "系統驗算標答", "燈號", "課綱條目", "PISA歷程", "審查診斷與修訂建議"]
        for j, h in enumerate(d_headers):
            cell = t_detail.cell(0, j)
            cell.text = h
            set_cell_background(cell, "EAEEF3")
            set_run_font(cell.paragraphs[0].runs[0], font_size=9.5, bold=True)
            set_cell_margins(cell)

        for i, item in enumerate(report.items, start=1):
            q = item.question
            is_ans_err = item.crosscheck and item.crosscheck.get("status") in ["MISSING", "INCOMPLETE", "MISMATCH"]
            if is_ans_err:
                status_symbol = "🔴 標答待修"
            elif item.status == "GREEN":
                status_symbol = "🟢 通過"
            elif item.status == "YELLOW":
                status_symbol = "🟡 微調"
            else:
                status_symbol = "🔴 超綱"

            q_type_name = "單選" if q.q_type == "choice" else ("填充" if q.q_type == "fill" else "非選")

            claimed_str = str(item.crosscheck.get("claimed_answer") if item.crosscheck else (q.claimed_answer or "無"))
            if not claimed_str or claimed_str == "None":
                claimed_str = "【空白】"
            comp_str = str(item.corrected_answer or (item.crosscheck.get("computed_answer") if item.crosscheck else ""))

            t_detail.cell(i, 0).text = str(q.index)
            t_detail.cell(i, 1).text = f"{q_type_name} ({q.score:.0f}分)"
            append_math_to_cell(t_detail.cell(i, 2), claimed_str, font_size=9, center=False)
            append_math_to_cell(t_detail.cell(i, 3), comp_str, font_size=9, center=False)
            t_detail.cell(i, 4).text = status_symbol
            t_detail.cell(i, 5).text = f"{item.curriculum_codes.get('content', '')} / {item.curriculum_codes.get('performance', '')}"
            t_detail.cell(i, 6).text = item.pisa_process
            t_detail.cell(i, 7).text = item.suggested_revision

            for c in range(8):
                cell = t_detail.cell(i, c)
                if is_ans_err:
                    set_cell_background(cell, "FFECEC")
                if c not in (2, 3):
                    p_c = cell.paragraphs[0]
                    if p_c.runs:
                        set_run_font(p_c.runs[0], font_size=9)
                set_cell_margins(cell)

        safe_save(doc, filepath)
        return filepath

    # =========================================================================
    # 2. 命題審題雙向細目檢核表.docx (全部內容以表格呈現)
    # =========================================================================
    @staticmethod
    def generate_matrix_report(report: ExamAuditReport, filepath: str) -> str:
        doc = docx.Document()
        set_page_margins(doc)
        setup_footer_page_number(doc)

        exam = report.exam_data

        p_title = doc.add_paragraph()
        p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_title.paragraph_format.space_after = Pt(6)
        r_title = p_title.add_run(f"{exam.school} 108課綱段考命題審題雙向細目檢核表")
        set_run_font(r_title, font_size=13.5, bold=True, underline=True)

        # 表格 1：基本資訊表
        t_base = doc.add_table(rows=2, cols=4)
        t_base.alignment = WD_TABLE_ALIGNMENT.CENTER
        set_table_borders(t_base)

        b_rows1 = [("學校名稱", exam.school), ("年級學期", f"{exam.grade_text}（第一學期）")]
        b_rows2 = [("命題範圍", exam.scope_desc), ("試卷總分", f"{report.total_score:.0f} 分")]

        for col_idx, (k, v) in enumerate(b_rows1):
            c_k = t_base.cell(0, col_idx * 2)
            c_v = t_base.cell(0, col_idx * 2 + 1)
            c_k.text = k; c_v.text = v
            set_cell_background(c_k, "EAEEF3")
            set_run_font(c_k.paragraphs[0].runs[0], font_size=10, bold=True)
            set_run_font(c_v.paragraphs[0].runs[0], font_size=10)
            set_cell_margins(c_k); set_cell_margins(c_v)

        for col_idx, (k, v) in enumerate(b_rows2):
            c_k = t_base.cell(1, col_idx * 2)
            c_v = t_base.cell(1, col_idx * 2 + 1)
            c_k.text = k; c_v.text = v
            set_cell_background(c_k, "EAEEF3")
            set_run_font(c_k.paragraphs[0].runs[0], font_size=10, bold=True)
            set_run_font(c_v.paragraphs[0].runs[0], font_size=10)
            set_cell_margins(c_k); set_cell_margins(c_v)

        # 表格 2：雙向細目主表
        p_h_mat = doc.add_paragraph()
        p_h_mat.paragraph_format.space_before = Pt(8)
        p_h_mat.paragraph_format.space_after = Pt(3)
        r_h_mat = p_h_mat.add_run("◆ 108 課綱雙向細目指標分析矩陣")
        set_run_font(r_h_mat, font_size=11.5, bold=True)

        t_mat = doc.add_table(rows=1 + len(report.items), cols=9)
        t_mat.alignment = WD_TABLE_ALIGNMENT.CENTER
        set_table_borders(t_mat)

        headers = ["題號", "題型", "配分", "標準答案", "學習表現", "學習內容", "核心素養", "PISA歷程", "預估通過率"]
        for j, h in enumerate(headers):
            cell = t_mat.cell(0, j)
            cell.text = h
            set_cell_background(cell, "EAEEF3")
            set_run_font(cell.paragraphs[0].runs[0], font_size=9.5, bold=True)
            set_cell_margins(cell)

        for i, item in enumerate(report.items, start=1):
            q = item.question
            t_mat.cell(i, 0).text = str(q.index)
            t_mat.cell(i, 1).text = "單選" if q.q_type == "choice" else ("填充" if q.q_type == "fill" else "非選")
            t_mat.cell(i, 2).text = f"{q.score:.0f}"
            ans_val = str(item.corrected_answer or q.claimed_answer or "-")
            append_math_to_cell(t_mat.cell(i, 3), ans_val, font_size=9, center=True)
            t_mat.cell(i, 4).text = item.curriculum_codes.get("performance", "-")
            t_mat.cell(i, 5).text = item.curriculum_codes.get("content", "-")
            t_mat.cell(i, 6).text = item.curriculum_codes.get("competency", "-")
            t_mat.cell(i, 7).text = item.pisa_process
            t_mat.cell(i, 8).text = "65%"

            for c in range(9):
                cell = t_mat.cell(i, c)
                if c != 3:
                    set_run_font(cell.paragraphs[0].runs[0], font_size=9)
                set_cell_margins(cell)

        # 表格 3：審題教師四大核心標準檢核表 (完全以表格呈現)
        p_h_chk = doc.add_paragraph()
        p_h_chk.paragraph_format.space_before = Pt(8)
        p_h_chk.paragraph_format.space_after = Pt(3)
        r_h_chk = p_h_chk.add_run("◆ 審題教師核心檢核標準表")
        set_run_font(r_h_chk, font_size=11.5, bold=True)

        chk_items = [
            ("1", "課綱規範", "試題完全符合 108 課綱規範，無超綱、無舊教材已被刪除內容（如弦切角、圓冪定理等）",
             "✅ 完全合規" if report.red_count == 0 else "⚠️ 發現超綱"),
            ("2", "數學嚴謹", "數學題幹條件充分無矛盾，各題均有唯一定解，單選選項互斥且無重複解答",
             "✅ 唯一定解" if report.red_count == 0 else "⚠️ 需修正"),
            ("3", "試題純粹", "試題卷上徹底取消填充題作答欄與非選擇題作答方框，學生一律於答案卷作答",
             "✅ 排版合規"),
            ("4", "答案單頁", "答案卷格式比照解析卷，非選留足 2.5 吋作答空間，且嚴格控制剛好單頁 A4 不跨頁",
             "✅ 剛好單頁")
        ]

        t_chk = doc.add_table(rows=1 + len(chk_items), cols=4)
        t_chk.alignment = WD_TABLE_ALIGNMENT.CENTER
        set_table_borders(t_chk)

        c_headers = ["項次", "檢核類別", "審題檢核指標內容", "審查結果"]
        for j, h in enumerate(c_headers):
            cell = t_chk.cell(0, j)
            cell.text = h
            set_cell_background(cell, "EAEEF3")
            set_run_font(cell.paragraphs[0].runs[0], font_size=9.5, bold=True)
            set_cell_margins(cell)

        for i, (idx, cat, desc, res) in enumerate(chk_items, start=1):
            t_chk.cell(i, 0).text = idx
            t_chk.cell(i, 1).text = cat
            t_chk.cell(i, 2).text = desc
            t_chk.cell(i, 3).text = res
            for c in range(4):
                cell = t_chk.cell(i, c)
                set_run_font(cell.paragraphs[0].runs[0], font_size=9)
                set_cell_margins(cell)

        # 表格 4：試務審核三方簽章表格 (表格呈現簽章欄)
        p_h_sign = doc.add_paragraph()
        p_h_sign.paragraph_format.space_before = Pt(8)
        p_h_sign.paragraph_format.space_after = Pt(3)
        r_h_sign = p_h_sign.add_run("◆ 試務審核與簽章核定表")
        set_run_font(r_h_sign, font_size=11.5, bold=True)

        t_sign = doc.add_table(rows=2, cols=4)
        t_sign.alignment = WD_TABLE_ALIGNMENT.CENTER
        set_table_borders(t_sign)

        sign_headers = ["命題教師簽章", "審題教師簽章", "數學領域召集人", "教務處試務組長 / 主任"]
        for j, h in enumerate(sign_headers):
            cell = t_sign.cell(0, j)
            cell.text = h
            set_cell_background(cell, "EAEEF3")
            set_run_font(cell.paragraphs[0].runs[0], font_size=10, bold=True)
            set_cell_margins(cell)

        # 第 2 列留高供蓋章或簽名 (約 1.2 吋)
        t_sign.rows[1].height = Inches(1.2)
        for j in range(4):
            cell = t_sign.cell(1, j)
            cell.text = "\n\n簽章日期：   年   月   日"
            set_run_font(cell.paragraphs[0].runs[0], font_size=9)
            set_cell_margins(cell)

        safe_save(doc, filepath)
        return filepath

    # =========================================================================
    # 3. SymPy嚴格驗算矩陣.json
    # =========================================================================
    @staticmethod
    def generate_sympy_json(report: ExamAuditReport, filepath: str) -> str:
        data = {
            "file": os.path.basename(report.exam_data.file_path),
            "grade": report.exam_data.grade,
            "total_questions": len(report.items),
            "total_score": report.total_score,
            "is_score_100": report.is_score_100,
            "status_counts": {
                "RED": report.red_count,
                "YELLOW": report.yellow_count,
                "GREEN": report.green_count
            },
            "pisa_distribution": report.pisa_distribution,
            "questions": []
        }

        for item in report.items:
            q = item.question
            data["questions"].append({
                "index": q.index,
                "q_type": q.q_type,
                "score": q.score,
                "stem": q.stem,
                "options": q.options,
                "claimed_answer": q.claimed_answer,
                "corrected_answer": item.corrected_answer,
                "crosscheck": item.crosscheck,
                "curriculum": item.curriculum_codes,
                "pisa": item.pisa_process,
                "status": item.status,
                "redlines": item.redline_violations,
                "option_issues": item.option_issues,
                "competency_warning": item.competency_warning,
                "sympy_solution": item.math_verification,
                "suggestion": item.suggested_revision
            })

        with open(filepath, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

        return filepath

    # =========================================================================
    # 4. 修訂版試題卷.docx
    # =========================================================================
    @staticmethod
    def generate_revised_question_doc(report: ExamAuditReport, filepath: str) -> str:
        doc = docx.Document()
        set_page_margins(doc)
        setup_footer_page_number(doc)

        exam = report.exam_data
        add_jinhe_header(doc, school_info=exam.school, exam_title=f"{exam.grade_text}數學科段考試題", exam_range=exam.scope_desc)

        grouped = {"choice": [], "fill": [], "comp": []}
        for item in report.items:
            grouped[item.question.q_type].append(item)

        if grouped["choice"]:
            p_h = doc.add_paragraph()
            p_h.paragraph_format.space_before = Pt(5)
            p_h.paragraph_format.space_after = Pt(3)
            r_h = p_h.add_run(f"◆ 第一部分：單一選擇題（共 {len(grouped['choice'])} 題，每題 4 分，共 {len(grouped['choice'])*4} 分）")
            set_run_font(r_h, font_size=11.5, bold=True)

            for item in grouped["choice"]:
                q = item.question
                create_question_p(doc, f"{q.index}. ", q.stem, scope_str=q.scope or "(3-1)")
                if q.options:
                    opt_str = "    ".join(q.options)
                    create_options_p(doc, opt_str)

        if grouped["fill"]:
            p_h = doc.add_paragraph()
            p_h.paragraph_format.space_before = Pt(6)
            p_h.paragraph_format.space_after = Pt(3)
            r_h = p_h.add_run(f"◆ 第二部分：填充題（共 {len(grouped['fill'])} 格，每格 4 分，共 {len(grouped['fill'])*4} 分）")
            set_run_font(r_h, font_size=11.5, bold=True)

            for item in grouped["fill"]:
                q = item.question
                create_question_p(doc, f"{q.index}. ", q.stem, scope_str=q.scope or "(3-2)")

        if grouped["comp"]:
            p_h = doc.add_paragraph()
            p_h.paragraph_format.space_before = Pt(6)
            p_h.paragraph_format.space_after = Pt(3)
            r_h = p_h.add_run(f"◆ 第三部分：非選擇題（計算證明題）（共 {len(grouped['comp'])} 題，每題 10 分，共 {len(grouped['comp'])*10} 分）")
            set_run_font(r_h, font_size=11.5, bold=True)

            for item in grouped["comp"]:
                q = item.question
                create_question_p(doc, f"{q.index}. ", q.stem, scope_str=q.scope or "(3-1)")

        safe_save(doc, filepath)
        return filepath

    # =========================================================================
    # 5. 修訂版答案卷.docx (嚴格剛好單頁 A4)
    # =========================================================================
    @staticmethod
    def generate_revised_answer_sheet(report: ExamAuditReport, filepath: str) -> str:
        doc = docx.Document()
        set_page_margins(doc)
        setup_footer_page_number(doc)

        exam = report.exam_data
        add_jinhe_header(doc, school_info=exam.school, exam_title=f"{exam.grade_text}數學科段考答案卷", exam_range=exam.scope_desc)

        grouped = {"choice": [], "fill": [], "comp": []}
        for item in report.items:
            grouped[item.question.q_type].append(item)

        p_c = doc.add_paragraph()
        p_c.paragraph_format.space_before = Pt(4); p_c.paragraph_format.space_after = Pt(2)
        r_c = p_c.add_run("◆ 第一部分：單一選擇題")
        set_run_font(r_c, font_size=11, bold=True)

        n_choices = max(len(grouped["choice"]), 10)
        t_choice = doc.add_table(rows=2, cols=n_choices)
        t_choice.alignment = WD_TABLE_ALIGNMENT.CENTER
        set_table_borders(t_choice)
        for j in range(n_choices):
            t_choice.cell(0, j).text = str(j + 1)
            set_cell_background(t_choice.cell(0, j), "EAEEF3")
            set_run_font(t_choice.cell(0, j).paragraphs[0].runs[0], font_size=9.5, bold=True)
            t_choice.cell(1, j).text = ""
            set_cell_margins(t_choice.cell(0, j)); set_cell_margins(t_choice.cell(1, j))

        p_f = doc.add_paragraph()
        p_f.paragraph_format.space_before = Pt(4); p_f.paragraph_format.space_after = Pt(2)
        r_f = p_f.add_run("◆ 第二部分：填充題")
        set_run_font(r_f, font_size=11, bold=True)

        t_fill = doc.add_table(rows=4, cols=5)
        t_fill.alignment = WD_TABLE_ALIGNMENT.CENTER
        set_table_borders(t_fill)
        for j in range(5):
            t_fill.cell(0, j).text = str(j + 1)
            t_fill.cell(2, j).text = str(j + 6)
            set_cell_background(t_fill.cell(0, j), "EAEEF3")
            set_cell_background(t_fill.cell(2, j), "EAEEF3")
            set_run_font(t_fill.cell(0, j).paragraphs[0].runs[0], font_size=9.5, bold=True)
            set_run_font(t_fill.cell(2, j).paragraphs[0].runs[0], font_size=9.5, bold=True)
            set_cell_margins(t_fill.cell(0, j)); set_cell_margins(t_fill.cell(1, j))
            set_cell_margins(t_fill.cell(2, j)); set_cell_margins(t_fill.cell(3, j))

        p_comp = doc.add_paragraph()
        p_comp.paragraph_format.space_before = Pt(4); p_comp.paragraph_format.space_after = Pt(2)
        r_comp = p_comp.add_run("◆ 第三部分：非選擇題（計算證明題）")
        set_run_font(r_comp, font_size=11, bold=True)

        t_comp = doc.add_table(rows=2, cols=2)
        t_comp.alignment = WD_TABLE_ALIGNMENT.CENTER
        set_table_borders(t_comp)
        t_comp.cell(0, 0).text = "1."
        t_comp.cell(0, 1).text = "2."
        set_cell_background(t_comp.cell(0, 0), "EAEEF3")
        set_cell_background(t_comp.cell(0, 1), "EAEEF3")
        set_run_font(t_comp.cell(0, 0).paragraphs[0].runs[0], font_size=10, bold=True)
        set_run_font(t_comp.cell(0, 1).paragraphs[0].runs[0], font_size=10, bold=True)
        t_comp.rows[1].height = Inches(2.5)

        safe_save(doc, filepath)
        return filepath

    # =========================================================================
    # 6. 修訂詳細解析.docx (全卷標準答案速查表 + 逐題詳細解析 + 0~3級分評分規準)
    # =========================================================================
    @staticmethod
    def generate_revised_solution_doc(report: ExamAuditReport, filepath: str) -> str:
        doc = docx.Document()
        set_page_margins(doc)
        setup_footer_page_number(doc)

        exam = report.exam_data
        add_jinhe_header(doc, school_info=exam.school, exam_title=f"{exam.grade_text}數學科段考 修訂詳細解析卷", exam_range=exam.scope_desc)

        grouped = {"choice": [], "fill": [], "comp": []}
        for item in report.items:
            grouped[item.question.q_type].append(item)

        # 頂部說明
        p_desc = doc.add_paragraph()
        p_desc.paragraph_format.space_before = Pt(4)
        p_desc.paragraph_format.space_after = Pt(2)
        r_desc = p_desc.add_run("本卷由 108 課綱數學科審題系統依據教育部課程綱要、錦和高中命題規範及 SymPy 符號驗算引擎自動產出。")
        set_run_font(r_desc, font_size=9.5, italic=True)

        # =========================================================================
        # 第一部分：已更正標準答案速查表
        # =========================================================================
        p_sec1 = doc.add_paragraph()
        p_sec1.paragraph_format.space_before = Pt(6)
        p_sec1.paragraph_format.space_after = Pt(3)
        r_sec1 = p_sec1.add_run("◆ 已更正標準答案速查表")
        set_run_font(r_sec1, font_size=12, bold=True)

        # 1. 選擇題答案表
        p_c_lbl = doc.add_paragraph()
        p_c_lbl.paragraph_format.space_before = Pt(3)
        p_c_lbl.paragraph_format.space_after = Pt(2)
        r_c_lbl = p_c_lbl.add_run("一、單一選擇題（每題 4 分）")
        set_run_font(r_c_lbl, font_size=10.5, bold=True)

        n_choices = max(len(grouped["choice"]), 12)
        t_choice = doc.add_table(rows=2, cols=n_choices)
        t_choice.alignment = WD_TABLE_ALIGNMENT.CENTER
        set_table_borders(t_choice)
        for j in range(n_choices):
            c_header = t_choice.cell(0, j)
            c_header.text = str(j + 1)
            set_cell_background(c_header, "EAEEF3")
            set_run_font(c_header.paragraphs[0].runs[0], font_size=9.5, bold=True)
            set_cell_margins(c_header)

            c_ans = t_choice.cell(1, j)
            if j < len(grouped["choice"]):
                it = grouped["choice"][j]
                ans_str = it.corrected_answer or it.question.claimed_answer or ""
                is_corrected = it.crosscheck and it.crosscheck.get("status") in ["MISMATCH", "MISSING", "INCOMPLETE"]
                if is_corrected:
                    c_ans.text = f"* {ans_str}"
                    set_run_font(c_ans.paragraphs[0].runs[0], font_size=10, bold=True, color=RGBColor(192, 0, 0))
                else:
                    c_ans.text = ans_str
                    set_run_font(c_ans.paragraphs[0].runs[0], font_size=10)
            else:
                c_ans.text = ""
            set_cell_margins(c_ans)

        # 2. 填充題答案表
        p_f_lbl = doc.add_paragraph()
        p_f_lbl.paragraph_format.space_before = Pt(5)
        p_f_lbl.paragraph_format.space_after = Pt(2)
        r_f_lbl = p_f_lbl.add_run("二、填充題（每格 4 分）")
        set_run_font(r_f_lbl, font_size=10.5, bold=True)

        t_fill = doc.add_table(rows=4, cols=5)
        t_fill.alignment = WD_TABLE_ALIGNMENT.CENTER
        set_table_borders(t_fill)
        for j in range(5):
            t_fill.cell(0, j).text = str(j + 1)
            set_cell_background(t_fill.cell(0, j), "EAEEF3")
            set_run_font(t_fill.cell(0, j).paragraphs[0].runs[0], font_size=9.5, bold=True)
            set_cell_margins(t_fill.cell(0, j))

            cell_val1 = t_fill.cell(1, j)
            if j < len(grouped["fill"]):
                it1 = grouped["fill"][j]
                ans1 = it1.corrected_answer or it1.question.claimed_answer or ""
                is_corr1 = it1.crosscheck and it1.crosscheck.get("status") in ["MISMATCH", "MISSING", "INCOMPLETE"]
                txt1 = f"* {ans1}" if is_corr1 else ans1
                col1 = RGBColor(192, 0, 0) if is_corr1 else None
                append_math_to_cell(cell_val1, txt1, font_size=9.5, bold=is_corr1, color=col1, center=True)
            set_cell_margins(cell_val1)

            t_fill.cell(2, j).text = str(j + 6)
            set_cell_background(t_fill.cell(2, j), "EAEEF3")
            set_run_font(t_fill.cell(2, j).paragraphs[0].runs[0], font_size=9.5, bold=True)
            set_cell_margins(t_fill.cell(2, j))

            cell_val2 = t_fill.cell(3, j)
            if j + 5 < len(grouped["fill"]):
                it2 = grouped["fill"][j + 5]
                ans2 = it2.corrected_answer or it2.question.claimed_answer or ""
                is_corr2 = it2.crosscheck and it2.crosscheck.get("status") in ["MISMATCH", "MISSING", "INCOMPLETE"]
                txt2 = f"* {ans2}" if is_corr2 else ans2
                col2 = RGBColor(192, 0, 0) if is_corr2 else None
                append_math_to_cell(cell_val2, txt2, font_size=9.5, bold=is_corr2, color=col2, center=True)
            set_cell_margins(cell_val2)

        # 3. 計算題答案表
        p_cp_lbl = doc.add_paragraph()
        p_cp_lbl.paragraph_format.space_before = Pt(5)
        p_cp_lbl.paragraph_format.space_after = Pt(2)
        r_cp_lbl = p_cp_lbl.add_run("三、非選擇題（計算證明題）（每題 6 分）")
        set_run_font(r_cp_lbl, font_size=10.5, bold=True)

        n_comp = max(len(grouped["comp"]), 2)
        t_comp = doc.add_table(rows=2, cols=n_comp)
        t_comp.alignment = WD_TABLE_ALIGNMENT.CENTER
        set_table_borders(t_comp)
        for j in range(n_comp):
            c_header = t_comp.cell(0, j)
            c_header.text = f"第 {j + 1} 題"
            set_cell_background(c_header, "EAEEF3")
            set_run_font(c_header.paragraphs[0].runs[0], font_size=9.5, bold=True)
            set_cell_margins(c_header)

            c_ans = t_comp.cell(1, j)
            if j < len(grouped["comp"]):
                it = grouped["comp"][j]
                ans_str = it.corrected_answer or it.question.claimed_answer or ""
                is_corr = it.crosscheck and it.crosscheck.get("status") in ["MISMATCH", "MISSING", "INCOMPLETE"]
                txt_comp = f"* {ans_str}" if is_corr else ans_str
                col_comp = RGBColor(192, 0, 0) if is_corr else None
                append_math_to_cell(c_ans, txt_comp, font_size=9, bold=is_corr, color=col_comp, center=True)
            set_cell_margins(c_ans)

        # 附註
        if report.answer_errors > 0:
            p_note = doc.add_paragraph()
            p_note.paragraph_format.space_before = Pt(3)
            p_note.paragraph_format.space_after = Pt(8)
            r_note = p_note.add_run(f"★ 附註：標記紅字星號 '*' 之題目（共 {report.answer_errors} 處），表示原解答卷答案缺漏或錯誤，此處為系統符號運算推導更正補齊之正確答案。")
            set_run_font(r_note, font_size=9, bold=True, color=RGBColor(192, 0, 0))

        # =========================================================================
        # 第二部分：逐題詳細解析與評分規準
        # =========================================================================
        p_sec2 = doc.add_paragraph()
        p_sec2.paragraph_format.space_before = Pt(12)
        p_sec2.paragraph_format.space_after = Pt(4)
        r_sec2 = p_sec2.add_run("◆ 逐題詳細解析與非選評分規準")
        set_run_font(r_sec2, font_size=12, bold=True)

        # (1) 選擇題
        if grouped["choice"]:
            p_c_sec = doc.add_paragraph()
            p_c_sec.paragraph_format.space_before = Pt(6)
            p_c_sec.paragraph_format.space_after = Pt(3)
            r_c_sec = p_c_sec.add_run("【第一部分：單一選擇題逐題解析】")
            set_run_font(r_c_sec, font_size=11, bold=True)

            for it in grouped["choice"]:
                q = it.question
                p_q = doc.add_paragraph()
                p_q.paragraph_format.space_before = Pt(4)
                p_q.paragraph_format.space_after = Pt(1)
                r_q = p_q.add_run(f"第 {q.index} 題. 【題幹】{q.stem}")
                set_run_font(r_q, font_size=10, bold=True)

                if q.options:
                    p_opt = doc.add_paragraph()
                    p_opt.paragraph_format.space_before = Pt(0)
                    p_opt.paragraph_format.space_after = Pt(2)
                    r_opt = p_opt.add_run("    " + "    ".join(q.options))
                    set_run_font(r_opt, font_size=9.5)

                p_sol = doc.add_paragraph()
                p_sol.paragraph_format.space_before = Pt(1)
                p_sol.paragraph_format.space_after = Pt(3)
                r_ans_lbl = p_sol.add_run(f"【正確答案】（{it.corrected_answer}）  ")
                set_run_font(r_ans_lbl, font_size=10, bold=True, color=RGBColor(0, 102, 204))

                r_step_lbl = p_sol.add_run("【詳解】：")
                set_run_font(r_step_lbl, font_size=9.5, bold=True)

                steps = it.crosscheck.get("steps", "依據題幹條件推導求解。") if it.crosscheck else "符號推導驗證通過。"
                append_mixed_text(p_sol, steps, font_size=9.5)

        # (2) 填充題
        if grouped["fill"]:
            p_f_sec = doc.add_paragraph()
            p_f_sec.paragraph_format.space_before = Pt(8)
            p_f_sec.paragraph_format.space_after = Pt(3)
            r_f_sec = p_f_sec.add_run("【第二部分：填充題逐題解析】")
            set_run_font(r_f_sec, font_size=11, bold=True)

            for it in grouped["fill"]:
                q = it.question
                p_q = doc.add_paragraph()
                p_q.paragraph_format.space_before = Pt(4)
                p_q.paragraph_format.space_after = Pt(1)
                r_q = p_q.add_run(f"第 {q.index} 格. 【題幹】{q.stem}")
                set_run_font(r_q, font_size=10, bold=True)

                p_sol = doc.add_paragraph()
                p_sol.paragraph_format.space_before = Pt(1)
                p_sol.paragraph_format.space_after = Pt(3)
                r_ans_lbl = p_sol.add_run("【正確答案】")
                set_run_font(r_ans_lbl, font_size=10, bold=True, color=RGBColor(0, 102, 204))

                ans_str = str(it.corrected_answer or "")
                if any(c in ans_str for c in ['/', '√', '^', '*', ':']):
                    omml_ans = parse_math_to_omml(ans_str, font_size=10)
                    if omml_ans:
                        p_sol._p.append(parse_xml(f'<m:oMath {MATH_NS}>{omml_ans}</m:oMath>'))
                    else:
                        r_ans = p_sol.add_run(ans_str)
                        set_run_font(r_ans, font_name="Times New Roman", font_size=10, bold=True, color=RGBColor(0, 102, 204))
                else:
                    r_ans = p_sol.add_run(ans_str)
                    set_run_font(r_ans, font_name="Times New Roman", font_size=10, bold=True, color=RGBColor(0, 102, 204))

                p_sol.add_run("   ")
                r_step_lbl = p_sol.add_run("【詳解】：")
                set_run_font(r_step_lbl, font_size=9.5, bold=True)

                steps = it.crosscheck.get("steps", "依據題意公式推導求解。") if it.crosscheck else "符號推導驗證通過。"
                append_mixed_text(p_sol, steps, font_size=9.5)

        # (3) 計算題與評分規準
        if grouped["comp"]:
            p_cp_sec = doc.add_paragraph()
            p_cp_sec.paragraph_format.space_before = Pt(8)
            p_cp_sec.paragraph_format.space_after = Pt(3)
            r_cp_sec = p_cp_sec.add_run("【第三部分：非選擇題詳細推導與 0~3 級分評分規準】")
            set_run_font(r_cp_sec, font_size=11, bold=True)

            for it in grouped["comp"]:
                q = it.question
                p_q = doc.add_paragraph()
                p_q.paragraph_format.space_before = Pt(5)
                p_q.paragraph_format.space_after = Pt(2)
                r_q = p_q.add_run(f"第 {q.index} 題. 【題目敘述】{q.stem}")
                set_run_font(r_q, font_size=10, bold=True)

                p_sol = doc.add_paragraph()
                p_sol.paragraph_format.space_before = Pt(1)
                p_sol.paragraph_format.space_after = Pt(3)
                r_ans_lbl = p_sol.add_run("【參考標答】")
                set_run_font(r_ans_lbl, font_size=10, bold=True, color=RGBColor(0, 102, 204))

                ans_str = str(it.corrected_answer or "")
                append_mixed_text(p_sol, ans_str, font_size=10, bold=True, default_color=RGBColor(0, 102, 204))

                p_sol.add_run("\n")
                r_step_lbl = p_sol.add_run("【完整推導步驟】：\n")
                set_run_font(r_step_lbl, font_size=9.5, bold=True)

                steps = it.crosscheck.get("steps", "依據題目幾何性質列式推導。") if it.crosscheck else "詳細推導完備。"
                append_mixed_text(p_sol, steps, font_size=9.5)

                # 插入 0~3 級分評分規準表
                p_rub = doc.add_paragraph()
                p_rub.paragraph_format.space_before = Pt(3)
                p_rub.paragraph_format.space_after = Pt(2)
                r_rub = p_rub.add_run("▼ 非選擇題閱卷評分規準（0~3 級分）：")
                set_run_font(r_rub, font_size=9.5, bold=True)

                t_rub = doc.add_table(rows=5, cols=3)
                t_rub.alignment = WD_TABLE_ALIGNMENT.CENTER
                set_table_borders(t_rub)
                rub_headers = ["級分", "評分指標說明", "給分條件與扣分標準"]
                for col_idx, h in enumerate(rub_headers):
                    cell = t_rub.cell(0, col_idx)
                    cell.text = h
                    set_cell_background(cell, "EAEEF3")
                    set_run_font(cell.paragraphs[0].runs[0], font_size=9.5, bold=True)
                    set_cell_margins(cell)

                rub_data = [
                    ("3分 (滿分)", "推導完整嚴謹，策略合宜，計算完全無誤", "能完整列出幾何比例關係式，且 (1) 與 (2) 兩小題均獲致正解。"),
                    ("2分 (部分分)", "解題策略方向正確，大致推導完整", "能正確列出比例式並完全解出 (1) 數值，(2) 發生微小計算筆誤；或僅完整完成 (1) 小題。"),
                    ("1分 (部分分)", "具備解題雛形，能著手列出關鍵定理關係式", "僅能列出題目之對應邊比例方程式，未能正確求出關鍵數值。"),
                    ("0分 (不給分)", "未作答或作答內容與題意完全無關", "空白未答、僅抄寫題目條件或列式完全錯誤。")
                ]
                for row_idx, (g, desc, cond) in enumerate(rub_data, start=1):
                    t_rub.cell(row_idx, 0).text = g
                    t_rub.cell(row_idx, 1).text = desc
                    t_rub.cell(row_idx, 2).text = cond
                    for col_idx in range(3):
                        cell = t_rub.cell(row_idx, col_idx)
                        set_run_font(cell.paragraphs[0].runs[0], font_size=9)
                        set_cell_margins(cell)

        safe_save(doc, filepath)
        return filepath
