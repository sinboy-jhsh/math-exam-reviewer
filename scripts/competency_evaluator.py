"""
素養導向評量與 PISA 四大數學歷程評鑑模組
包含：
1. PISA 四大數學歷程判定 (M1推論, M2形成問題, M3運用程序, M4詮釋評鑑)
2. 題目情境文字負擔度與「偽素養」診斷 (Contextual Relevance Index)
3. 非選擇題 0~3 級分評分規準檢測
"""

import re
from typing import Dict, Any, List

class CompetencyEvaluator:
    # 歷程關鍵詞權重庫
    PROCESS_KEYWORDS = {
        "M1": ["證明", "推導", "推理", "外心", "內心", "重心", "全等", "相似", "性質", "說明理由", "真偽", "判定"],
        "M2": ["生活", "情境", "設計", "模型", "某公司", "規劃", "超市", "折扣", "購買", "圖形如下", "觀察下圖", "小明", "小華"],
        "M3": ["計算", "求值", "化簡", "方程式", "解出", "坐標", "斜率", "面積", "周長", "體積", "展開", "因式分解"],
        "M4": ["合理", "評鑑", "比較", "最佳", "利潤最大", "最划算", "檢驗", "詮釋", "是否可行", "哪一種方案"]
    }

    @staticmethod
    def evaluate_pisa_process(stem: str) -> Dict[str, Any]:
        """
        自動判斷題目主要的 PISA 數學推論歷程 (M1~M4)
        """
        scores = {"M1": 0, "M2": 0, "M3": 0, "M4": 0}
        for process, kw_list in CompetencyEvaluator.PROCESS_KEYWORDS.items():
            for kw in kw_list:
                if kw in stem:
                    scores[process] += 1

        # 若有包含生活文字且有計算
        if scores["M2"] > 0 and scores["M4"] > 0:
            primary = "M4"
        elif scores["M1"] > 0 and ("證明" in stem or "理由" in stem):
            primary = "M1"
        elif scores["M2"] > 0:
            primary = "M2"
        elif scores["M1"] > 0:
            primary = "M1"
        else:
            primary = "M3"

        labels = {
            "M1": "M1 (數學推論與論證)",
            "M2": "M2 (形成數學問題/建立模型)",
            "M3": "M3 (運用概念、符號與程序)",
            "M4": "M4 (詮釋、應用與評鑑)"
        }

        return {
            "primary_process": primary,
            "label": labels[primary],
            "scores": scores
        }

    @staticmethod
    def analyze_pseudo_competency(stem: str) -> Dict[str, Any]:
        """
        診斷「包裝型偽素養題」與文字負擔度：
        若題幹文字很長 (>150字)，但情境只是硬套名字（如小明小華買飲料），
        本質僅為極簡單之一步代數運算，且文字去除後完全不影響解題，
        則判定為文字冗餘度過高。
        """
        clean_text = re.sub(r"\s+", "", stem)
        total_chars = len(clean_text)

        # 生活情境詞彙統計
        context_words = ["小明", "小華", "阿姨", "超商", "咖啡", "特賣會", "搭車", "出遊", "遊樂園", "排隊", "飲料", "手機"]
        context_count = sum(1 for w in context_words if w in stem)

        is_verbose = total_chars > 160
        has_context = context_count > 0

        # 計算簡易核心算式數量
        eq_count = len(re.findall(r"[=><\+\-\*\/]", stem))

        is_pseudo = False
        warning = None

        if is_verbose and has_context and eq_count <= 2:
            is_pseudo = True
            warning = f"題幹字數達 {total_chars} 字，生活情境包裝厚重，但數學運算核心單純，疑似『偽素養包裝題』，恐徒增學生無效閱讀負擔。"
        elif total_chars > 220:
            warning = f"題幹文字量偏高（{total_chars} 字），建議精簡題目描述以符合段考時間分配。"

        return {
            "char_count": total_chars,
            "is_pseudo": is_pseudo,
            "warning": warning,
            "is_competency_oriented": has_context and not is_pseudo
        }

    @staticmethod
    def check_non_choice_rubric(claimed_solution: Optional[str]) -> Dict[str, Any]:
        """
        檢核非選擇題是否提供清晰的 0~3 級分標準規準：
        - 3 級分：策略嚴謹、計算無誤、幾何理由完整
        - 2 級分：核心關係列出但計算失誤或未化簡
        - 1 級分：僅列出單一中間量或已知條件
        - 0 級分：空白或完全無關
        """
        if not claimed_solution:
            return {"has_rubric": False, "msg": "未提供非選題評分規準"}

        has_levels = any(k in claimed_solution for k in ["3級分", "2級分", "1級分", "3 分", "2 分", "1 分", "評分規準", "給分標準"])
        return {
            "has_rubric": has_levels,
            "msg": "符合 0~3 級分階梯評分規準" if has_levels else "缺少明確的 0~3 級分規準，建議補充以利閱卷公平性"
        }
