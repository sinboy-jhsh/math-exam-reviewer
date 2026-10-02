"""
產生測試用待審考卷：input/115上國九數學段三試題.docx
內含：
1. 正常題目 (符合 108 課綱外心、內心、重心、相似形)
2. 故意包含的超綱違規題 (例如：考弦切角定理、考圓冪定理)
3. 故意包含的選項重複瑕疵題
4. 試題卷上刻意殘留作答欄（測試錦和排版純粹化審查）
"""

import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

from scripts.exam_helpers import (
    set_page_margins, setup_footer_page_number, set_run_font,
    add_jinhe_header, create_question_p, create_options_p, safe_save
)

def create_sample_exam():
    output_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "input", "115上國九數學段三試題.docx")
    os.makedirs(os.path.dirname(output_path), exist_ok=True)

    doc = docx.Document()
    set_page_margins(doc)
    setup_footer_page_number(doc)

    add_jinhe_header(
        doc,
        school_info="新北市立錦和高級中學",
        exam_title="115學年度第一學期國中部九年級數學科第三次段考試題",
        exam_range="九上 第 3 章 幾何與證明（3-1證明與推理、3-2三角形的外心、內心與重心）"
    )

    # 第一部分：單一選擇題
    p1 = doc.add_paragraph()
    p1.paragraph_format.space_before = Pt(6)
    r1 = p1.add_run("◆ 第一部分：單一選擇題（共 5 題，每題 4 分，共 20 分）")
    set_run_font(r1, font_size=11.5, bold=True)

    # 第 1 題：正常題 (外心性質)
    create_question_p(doc, "1. ", "若點 O 為直角三角形 ABC 的外心，且斜邊長為 10，則外接圓半徑為何？", scope_str="(3-2)")
    create_options_p(doc, "(A) 5    (B) 10    (C) 15    (D) 20")

    # 第 2 題：故意違規超綱題 (弦切角定理)
    create_question_p(doc, "2. ", "如圖，直線 L 切圓 O 於點 A，弦 AB 與直線 L 夾角為 40°，則利用弦切角定理可知劣弧 AB 的度數為何？", scope_str="(3-1)")
    create_options_p(doc, "(A) 40°    (B) 80°    (C) 120°    (D) 160°")

    # 第 3 題：故意選項瑕疵題 (選項重複)
    create_question_p(doc, "3. ", "在坐標平面上，已知三角形重心坐標為 G(2, 4)，頂點 A(0, 0)，B(3, 3)，求 C 點坐標？", scope_str="(3-2)")
    create_options_p(doc, "(A) (3, 9)    (B) (3, 9)    (C) (6, 9)    (D) (2, 5)")

    # 第 4 題：故意違規超綱題 (圓冪定理/相交弦定理)
    create_question_p(doc, "4. ", "已知圓內兩弦 AB 與 CD 交於 P 點，若 PA=3, PB=4, PC=2，利用相交弦定理求 PD 長度為何？", scope_str="(3-1)")
    create_options_p(doc, "(A) 6    (B) 8    (C) 10    (D) 12")

    # 第 5 題：正常題 (內心性質)
    create_question_p(doc, "5. ", "等腰直角三角形 ABC 中，兩股長皆為 6，求其內切圓半徑為何？", scope_str="(3-2)")
    create_options_p(doc, "(A) 6-3√2    (B) 3√2    (C) 6-√2    (D) 2")

    # 第二部分：填充題 (故意題號使用小括號 (1) 且留作答欄)
    p2 = doc.add_paragraph()
    p2.paragraph_format.space_before = Pt(8)
    r2 = p2.add_run("◆ 第二部分：填充題（共 5 格，每格 4 分，共 20 分）（請在作答欄作答）")
    set_run_font(r2, font_size=11.5, bold=True)

    create_question_p(doc, "1. ", "正三角形的邊長為 6，則其外接圓面積與內切圓面積之比值為 ____________。", scope_str="(3-2)")
    create_question_p(doc, "2. ", "若 G 為直角三角形之重心，兩股為 6 與 8，則斜邊中線長度為 ____________。", scope_str="(3-2)")
    create_question_p(doc, "3. ", "在鈍角三角形 ABC 中，其外心 O 位於三角形的 ____________ 部。（填內部、外部或邊上）", scope_str="(3-2)")
    create_question_p(doc, "4. ", "已知小明和小華在超商買飲料，小明買了 3 瓶可樂付了 90 元，設每瓶 x 元，求 x 之值為 ____________。（偽素養示範）", scope_str="(3-1)")
    create_question_p(doc, "5. ", "直角坐標系中，點 P(-2, 3) 到 x 軸的距離為 ____________。", scope_str="(3-1)")

    # 第三部分：非選擇題
    p3 = doc.add_paragraph()
    p3.paragraph_format.space_before = Pt(8)
    r3 = p3.add_run("◆ 第三部分：非選擇題（計算證明題）（共 2 題，每題 10 分，共 20 分）")
    set_run_font(r3, font_size=11.5, bold=True)

    create_question_p(doc, "1. ", "如圖，三角形 ABC 中，AB=AC，D、E 分別為 AB、AC 之中點。請證明：BE = CD。", scope_str="(3-1)")
    create_question_p(doc, "2. ", "已知直角三角形兩股分別為 5 與 12，求其內切圓半徑與外接圓半徑之和。", scope_str="(3-2)")

    safe_save(doc, output_path)
    print(f"範例待審考卷已生成：{output_path}")

if __name__ == "__main__":
    create_sample_exam()
