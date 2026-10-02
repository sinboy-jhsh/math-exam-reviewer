"""
國高中數學科 108 課綱素養命題與自動化審題主程式 (review.py)
一鍵掃描 input/ 資料夾中的待審考卷 (.docx 或 .pdf)，
自動執行課綱紅線檢核、SymPy 驗算與錦和排版審查，並輸出 4 份標準文件至 output/。
"""

import os
import sys
import argparse
from typing import List

# 加入當前路徑
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# 確保 Windows 命令提示字元與 PowerShell 正確支援 UTF-8 輸出
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

from scripts.exam_parser import ExamParser
from scripts.reviewer_core import ReviewerCore
from scripts.report_generator import ReportGenerator

def ensure_doc_converted(input_dir: str):
    """若發現舊版 Word 97-2003 (.doc)，自動調用 PowerShell Word COM 轉成 .docx"""
    doc_files = [f for f in os.listdir(input_dir) if f.lower().endswith(".doc") and not f.startswith("~$")]
    if doc_files:
        cmd = '''powershell -NoProfile -Command "$w = New-Object -ComObject Word.Application; $w.Visible = $false; try { Get-ChildItem '%s' -Filter *.doc | ForEach-Object { $src = $_.FullName; $dest = [System.IO.Path]::ChangeExtension($src, '.docx'); if (-not (Test-Path $dest)) { $d = $w.Documents.Open($src); $d.SaveAs2($dest, 16); $d.Close() } } } finally { $w.Quit() }"''' % input_dir
        os.system(cmd)

def main():
    parser = argparse.ArgumentParser(description="108 課綱數學科自動化審題引擎")
    parser.add_argument("--file", "-f", type=str, help="指定單一待審考卷路徑 (若不指定則自動掃描 input/ 目錄)")
    parser.add_argument("--input_dir", "-i", type=str, default="input", help="輸入資料夾路徑 (預設: input/)")
    parser.add_argument("--output_dir", "-o", type=str, default="output", help="輸出資料夾路徑 (預設: output/)")
    parser.add_argument("--scope", "-s", type=str, default=None, help="手動指定本次段考範圍 (可選)")
    args = parser.parse_args()

    input_dir = os.path.abspath(args.input_dir)
    output_dir = os.path.abspath(args.output_dir)
    os.makedirs(input_dir, exist_ok=True)
    os.makedirs(output_dir, exist_ok=True)

    files_to_review = []
    if args.file:
        if os.path.exists(args.file):
            files_to_review.append(os.path.abspath(args.file))
        else:
            print(f"[錯誤] 找不到指定的考卷檔案：{args.file}")
            sys.exit(1)
    else:
        # 掃描 input 目錄
        ensure_doc_converted(input_dir)
        for fname in os.listdir(input_dir):
            if fname.startswith("~$") or fname.startswith("."):
                continue
            # 若為解答卷、答案卷或解析卷，作為配對參考，不作為獨立待審試卷重複審核
            if any(k in fname for k in ["解答", "答案卷", "詳細解析"]):
                continue
            ext = os.path.splitext(fname)[1].lower()
            if ext in [".docx", ".pdf", ".txt", ".md"]:
                files_to_review.append(os.path.join(input_dir, fname))

    if not files_to_review:
        print("=" * 65)
        print("【提示】input/ 資料夾目前沒有任何待審考卷！")
        print("請將欲審核的 Word (.docx) 或 PDF (.pdf) 考卷放入以下目錄：")
        print(f"  -> {input_dir}")
        print("然後再次執行：python review.py")
        print("=" * 65)
        sys.exit(0)

    print("=" * 65)
    print(f"【開始 108 課綱數學科自動化審題】共發現 {len(files_to_review)} 份待審檔案")
    print("=" * 65)

    for idx, filepath in enumerate(files_to_review, start=1):
        filename = os.path.basename(filepath)
        print(f"\n[{idx}/{len(files_to_review)}] 正在解析並審核：{filename} ...")

        try:
            # 1. 解析試卷
            exam_data = ExamParser.parse(filepath)
            print(f"  -> 成功識別：{exam_data.school} {exam_data.grade_text}，共 {len(exam_data.questions)} 題，總分 {exam_data.total_score:.0f} 分")

            # 2. 核心審核
            report = ReviewerCore.audit_exam(exam_data, scope_hint=args.scope)
            ans_err_msg = f" (含 {report.answer_errors} 處標答瑕疵/漏填已校正)" if report.answer_errors > 0 else ""
            print(f"  -> 審查完成：[紅燈:超綱/標答瑕疵] {report.red_count} 題{ans_err_msg} | [黃燈:微調] {report.yellow_count} 題 | [綠燈:合規通過] {report.green_count} 題")

            # 3. 輸出報告與修訂卷
            out_files = ReportGenerator.generate_all(report, output_dir)
            print("  -> 成功產生 5 份錦和標準格式輸出文件：")
            print(f"     1. 審題綜合診斷報告：{os.path.basename(out_files['diagnosis'])}")
            print(f"     2. 雙向細目檢核表  ：{os.path.basename(out_files['matrix'])}")
            print(f"     3. SymPy驗算矩陣   ：{os.path.basename(out_files['sympy'])}")
            print(f"     4. 修訂版試題/答案卷：{os.path.basename(out_files['revised_question'])} & 答案卷")
            print(f"     5. 修訂詳細解析卷  ：{os.path.basename(out_files['revised_solution'])}")

        except Exception as e:
            print(f"  [錯誤] 處理 {filename} 時發生例外：{e}")
            import traceback
            traceback.print_exc()

    print("\n" + "=" * 65)
    print(f"【審題作業全數完成】所有審查報告與修訂考卷已存放於：\n  -> {output_dir}")
    print("=" * 65)

if __name__ == "__main__":
    main()
