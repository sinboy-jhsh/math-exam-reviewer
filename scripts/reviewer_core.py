"""
審題核心控制管線 (ReviewerCore)
串接：
- ExamParser (解析考卷)
- CurriculumAuditor (108課綱防超綱與雙向細目)
- SympyVerifier (代數與幾何驗算、選項互斥與答案分佈)
- CompetencyEvaluator (PISA 歷程與偽素養診斷)
"""

from dataclasses import dataclass, field
from typing import List, Dict, Any, Optional
import os
import re

from scripts.exam_parser import ExamData, ExamQuestion, ExamParser
from scripts.curriculum_rules import CurriculumAuditor
from scripts.sympy_verifier import SympyVerifier
from scripts.competency_evaluator import CompetencyEvaluator

@dataclass
class QuestionAuditItem:
    question: ExamQuestion
    curriculum_codes: Dict[str, str]       # content, performance, competency
    pisa_process: str                      # M1, M2, M3, M4
    pisa_label: str
    redline_violations: List[Dict[str, Any]]
    option_issues: List[str]
    competency_warning: Optional[str]
    math_verification: Optional[Dict[str, Any]]
    crosscheck: Optional[Dict[str, Any]] = None
    corrected_answer: Optional[str] = None
    suggested_revision: Optional[str] = None
    status: str = "GREEN"                  # "RED", "YELLOW", "GREEN"

@dataclass
class ExamAuditReport:
    exam_data: ExamData
    items: List[QuestionAuditItem] = field(default_factory=list)
    total_score: float = 0.0
    is_score_100: bool = True
    red_count: int = 0
    yellow_count: int = 0
    green_count: int = 0
    answer_errors: int = 0
    pisa_distribution: Dict[str, int] = field(default_factory=dict)
    choice_distribution: Dict[str, Any] = field(default_factory=dict)
    layout_violations: List[str] = field(default_factory=list)
    overall_summary: str = ""

class ReviewerCore:
    @staticmethod
    def audit_exam(exam: ExamData, scope_hint: Optional[str] = None) -> ExamAuditReport:
        report = ExamAuditReport(exam_data=exam)
        pisa_dist = {"M1": 0, "M2": 0, "M3": 0, "M4": 0}
        choice_answers = []

        # 1. 錦和校本排版與全卷通則檢核
        layout_violations = []

        # 檢查題幹中是否殘留「作答欄」或「作答區」
        for q in exam.questions:
            if "作答欄" in q.stem or "作答區" in q.stem:
                layout_violations.append(f"題號 {q.index}：試題卷殘留『作答欄/作答區』字樣，不符錦和試題卷純粹化規範。")
            if q.q_type == "fill":
                if f"({q.index})" in q.stem:
                    layout_violations.append(f"填充第 {q.index} 題：題號使用小括號，依規範應使用一般數字 '1.' ~ '10.'。")

        # 檢查總分是否為 100 分
        report.total_score = exam.total_score
        if abs(exam.total_score - 100.0) > 1e-3 and exam.total_score > 0:
            report.is_score_100 = False
            layout_violations.append(f"全卷配分加總為 {exam.total_score:.1f} 分，不符合正式段考滿分 100 分標準！")

        # 2. 逐題進行深度課綱、數學、素養與解答交叉審核
        for q in exam.questions:
            # (1) 108 課綱防超綱檢核
            redlines = CurriculumAuditor.audit_question(q.stem, grade=exam.grade, options=q.options)

            # (2) 雙向細目課綱編碼推論
            codes = CurriculumAuditor.infer_curriculum_codes(q.stem, default_scope=scope_hint or exam.scope_desc)

            # (3) PISA 四歷程判定
            pisa_info = CompetencyEvaluator.evaluate_pisa_process(q.stem)
            p_code = pisa_info["primary_process"]
            pisa_dist[p_code] = pisa_dist.get(p_code, 0) + 1

            # (4) 偽素養與文字負擔診斷
            comp_eval = CompetencyEvaluator.analyze_pseudo_competency(q.stem)
            comp_warn = comp_eval["warning"]

            # (5) SymPy 數學驗算、選項互斥與解答交叉比對
            opt_issues = []
            if q.q_type == "choice" and q.options:
                opt_check = SympyVerifier.verify_options_distinctness(q.options)
                if not opt_check["valid"]:
                    opt_issues.extend(opt_check["issues"])

            math_sol = SympyVerifier.solve_simple_stem(q.stem)

            # 核心亮點：推導答案與原卷標答交叉比對
            cross_res = SympyVerifier.verify_and_crosscheck_answer(q)
            corrected_ans = cross_res.get("corrected_answer") or cross_res.get("computed_answer") or q.claimed_answer

            if q.q_type == "choice" and corrected_ans:
                choice_answers.append(corrected_ans)

            has_answer_error = False
            if cross_res["status"] in ["MISSING", "INCOMPLETE", "MISMATCH"]:
                has_answer_error = True
                report.answer_errors += 1
                redlines.append({
                    "rule": cross_res["status"],
                    "level": "FATAL",
                    "name": f"【標答檢核】{cross_res.get('error', '標答有誤')}",
                    "suggestion": cross_res.get("reason", f"原解答卷標答【{q.claimed_answer}】與推導答案【{cross_res['computed_answer']}】不符。已更正為【{corrected_ans}】。")
                })

            # (6) 綜合狀態判定 (紅燈 FATAL / 黃燈 WARNING / 綠燈 PASS)
            status = "GREEN"
            has_fatal = any(r["level"] == "FATAL" for r in redlines)
            has_opt_conflict = len(opt_issues) > 0

            if has_fatal or has_opt_conflict or has_answer_error:
                status = "RED"
                report.red_count += 1
            elif redlines or comp_warn:
                status = "YELLOW"
                report.yellow_count += 1
            else:
                report.green_count += 1

            # 產生修正建議
            suggestions = []
            for r in redlines:
                suggestions.append(f"【{r['name']}】：{r['suggestion']}")
            if opt_issues:
                suggestions.append("【選項矛盾】：修正重複或實質相等的選項數值。")
            if comp_warn:
                suggestions.append(f"【文字負擔】：{comp_warn}")

            item = QuestionAuditItem(
                question=q,
                curriculum_codes=codes,
                pisa_process=p_code,
                pisa_label=pisa_info["label"],
                redline_violations=redlines,
                option_issues=opt_issues,
                competency_warning=comp_warn,
                math_verification=math_sol,
                crosscheck=cross_res,
                corrected_answer=corrected_ans,
                suggested_revision="\n".join(suggestions) if suggestions else "題意清晰，符應課綱指標，建議維持原題。",
                status=status
            )
            report.items.append(item)

        report.pisa_distribution = pisa_dist
        report.layout_violations = layout_violations

        # 診斷選擇題答案分佈均勻度
        if choice_answers:
            report.choice_distribution = SympyVerifier.audit_choice_distribution(choice_answers)

        # 3. 診斷總結
        q_count = len(exam.questions)
        summary = (
            f"本卷共審查 {q_count} 題，總分 {exam.total_score:.0f} 分。"
            f"合格題目 (綠燈) 共 {report.green_count} 題；"
            f"微調建議 (黃燈) 共 {report.yellow_count} 題；"
            f"重大違規/超綱/標答瑕疵 (紅燈) 共 {report.red_count} 題"
            f"（其中發現 {report.answer_errors} 處解答卷錯誤或漏填，已於修訂詳細解析卷全數更正補齊）。"
        )
        report.overall_summary = summary
        return report
