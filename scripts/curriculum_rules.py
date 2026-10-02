"""
國高中數學科 108 課綱防超綱紅線與雙向細目審查資料庫
包含：
1. 國中七至九年級各階段嚴格超綱紅線 (FATAL / WARNING)
2. 高中數 A / 數 B 分流防越界邊界
3. 108 課綱學習表現 (n/s/g/a/f/d)、學習內容 (N/S/G/A/F/D)、核心素養指標 (數-J / 數S-U)
4. PISA 四大數學歷程 (M1~M4) 判定映射
"""

import re
from typing import List, Dict, Any, Optional

REDLINE_RULES = [
    # ==========================================
    # 國中 7 年級紅線
    # ==========================================
    {
        "id": "G7-ABS-EQ",
        "grade": "G7",
        "name": "含未知數絕對值方程式",
        "level": "FATAL",
        "pattern": r"\|[0-9a-zA-Z\s\+\-\*\/]+\|\s*=\s*[0-9a-zA-Z\s\+\-\*\/]+",
        "keywords": ["|x", "|y", "絕對值方程式"],
        "desc": "108課綱國七絕對值僅限於數線距離與具體數值化簡，嚴禁考未知數絕對值方程 (如 |x-3|=5 或 |2x-1|=|x+3|)。",
        "suggestion": "改為純數值之絕對值計算，或以數線上兩點距離表示。"
    },
    {
        "id": "G7-NEG-EXP",
        "grade": "G7",
        "name": "負整數指數（負次方）",
        "level": "FATAL",
        "pattern": r"(\^\{-?[1-9]\d*\}|(?<![\w\-])[a-zA-Z0-9]+\^-\d+|負\s*[1-9]\d*\s*次方)",
        "keywords": ["負次方", "^-1", "^-2", "^-3"],
        "desc": "國七指數律嚴禁出現負整數次方 (如 2^-3 或 a^-1)，負指數為高一數與式主題。",
        "suggestion": "改以正整數次方運算，或以分數表示。"
    },
    {
        "id": "G7-DETACH-COEFF",
        "grade": "G7",
        "name": "分離係數法",
        "level": "WARNING",
        "pattern": r"分離係數",
        "keywords": ["分離係數法"],
        "desc": "108 課綱已弱化分離係數法特殊解題技巧，不應列為段考強制考核題型。",
        "suggestion": "以一般長除法或直式計算呈現。"
    },
    {
        "id": "G7-RECUR-DEC",
        "grade": "G7",
        "name": "循環小數化分數",
        "level": "FATAL",
        "pattern": r"循環小數|0\.\\overline|\.\.\.",
        "keywords": ["循環小數"],
        "desc": "循環小數轉分數公式已移至高中數與式，國中僅處理有限小數與分數轉換。",
        "suggestion": "改為有限小數或直接以最簡分數命題。"
    },

    # ==========================================
    # 國中 8 年級紅線
    # ==========================================
    {
        "id": "G8-FUNC-NOTATION",
        "grade": "G8",
        "name": "抽象函數符號 f(x)",
        "level": "FATAL",
        "pattern": r"f\s*\(\s*[xX]\s*\)|g\s*\(\s*[xX]\s*\)|f\s*\(\s*[0-9a-zA-Z\+\-]+\s*\)",
        "keywords": ["f(x)", "g(x)", "函數值 f("],
        "desc": "國八線型函數僅限使用 y=ax+b 表示，嚴禁出現高中抽象函數代號 f(x) 或 f(a+b)。",
        "suggestion": "將題幹中的 f(x) 改為 y 或以變數關係「當 x=... 時 y=...」呈現。"
    },
    {
        "id": "G8-CUBIC-FORMULA",
        "grade": "G8",
        "name": "三次乘法公式與三項和平方",
        "level": "FATAL",
        "pattern": r"a\^3\s*[\+\-]\s*b\^3|\(a\s*[\+\-]\s*b\)\^3|\(a\s*\+\s*b\s*\+\s*c\)\^2",
        "keywords": ["立方和", "立方差", "三項和平方"],
        "desc": "國中乘法公式僅限「和平方、差平方、平方差」，嚴禁出現三次展開或三項和公式。",
        "suggestion": "改用二次多項式或透過分組提公因式求解。"
    },
    {
        "id": "G8-DOUBLE-SQRT",
        "grade": "G8",
        "name": "雙重根號化簡",
        "level": "FATAL",
        "pattern": r"\\sqrt\{[^\}]*\\sqrt\{|\u221a[^\u221a]*\u221a",
        "keywords": ["雙重根號"],
        "desc": "雙重根號 √(a±2√b) 化簡已完全移至高一數與式，國中嚴禁命題。",
        "suggestion": "拆解為單一根式計算，或給予幾何圖形直角三角形斜邊計算。"
    },
    {
        "id": "G8-GEO-SERIES",
        "grade": "G8",
        "name": "等比級數求和公式",
        "level": "FATAL",
        "pattern": r"等比級數.*[總和|和|公式]|S_n\s*=.*r\^n",
        "keywords": ["等比級數和", "等比級數公式"],
        "desc": "國中等比僅限觀察數列前幾項之規則，嚴禁考等比級數總和公式 S_n=a1(1-r^n)/(1-r)。",
        "suggestion": "若有求和需求，限等差級數公式，等比僅能考前 3~4 項之逐項相加。"
    },

    # ==========================================
    # 國中 9 年級紅線
    # ==========================================
    {
        "id": "G9-QUAD-VERTEX",
        "grade": "G9",
        "name": "二次函數一般式配方法求極值",
        "level": "FATAL",
        "pattern": r"ax\^2\s*\+\s*bx\s*\+\s*c.*最[大小]值|一般式配方",
        "keywords": ["配方求最大值", "配方求最小值", "二次函數極值應用"],
        "desc": "108 課綱已將二次函數一般式 y=ax²+bx+c 配方法與極值應用移至高一；國九紙筆段考求極值題目必須直接給定「頂點式 y=a(x-h)²+k」或圖形頂點/對稱軸。",
        "suggestion": "將題幹改為頂點式 y=a(x-h)²+k，或明確給予頂點坐標及對稱點。"
    },
    {
        "id": "G9-CIRCLE-TANGENT-CHORD",
        "grade": "G9",
        "name": "弦切角定理（已廢除）",
        "level": "FATAL",
        "pattern": r"弦切角",
        "keywords": ["弦切角", "弦切角定理"],
        "desc": "108 課綱國中幾何已明確廢除「弦切角定理」，段考試題絕不可出現此名詞或需直接代公式之題型。",
        "suggestion": "改考圓心角、圓周角與切線垂直半徑之性質。"
    },
    {
        "id": "G9-CIRCLE-INTERIOR-EXTERIOR",
        "grade": "G9",
        "name": "圓內角/圓外角公式（已廢除）",
        "level": "FATAL",
        "pattern": r"圓內角|圓外角",
        "keywords": ["圓內角", "圓外角"],
        "desc": "108 課綱已刪除圓內角 (1/2(弧+弧)) 與圓外角 (1/2(大弧-小弧)) 公式，學生應使用三角形內角和與外角定理推導，嚴禁考套公式死背題。",
        "suggestion": "提供連心線或輔助線，引導利用圓周角與外角定理逐步解題。"
    },
    {
        "id": "G9-CIRCLE-POWER",
        "grade": "G9",
        "name": "圓冪定理/切割線定理（已廢除）",
        "level": "FATAL",
        "pattern": r"圓冪定理|切割線定理|相交弦定理",
        "keywords": ["圓冪定理", "切割線定理", "相交弦定理", "割線定理"],
        "desc": "108 課綱已完全刪除相交弦定理、割線定理與切割線定理（圓冪定理），嚴禁在國中命題。",
        "suggestion": "改以相似三角形對應邊成比例的方式，逐步建立推論。"
    },
    {
        "id": "G9-TRIG-GENERAL-ANGLE",
        "grade": "G9",
        "name": "廣義角與非特殊角三角比",
        "level": "FATAL",
        "pattern": r"sin\s*(?!30|45|60)\d{2,3}|cos\s*(?!30|45|60)\d{2,3}|tan\s*(?!30|45|60)\d{2,3}|正弦定理|餘弦定理",
        "keywords": ["正弦定理", "餘弦定理", "廣義角", "鈍角三角比"],
        "desc": "國中三角比嚴格限定於銳角 30°、45°、60° 之直角三角形邊長比，不考鈍角/廣義角、正弦定理與餘弦定理。",
        "suggestion": "改用特殊直角三角形 (1:1:√2 或 1:√3:2) 或給定直角邊長直接求比值。"
    },

    # ==========================================
    # 高中 10~12 年級與數 A / 數 B 分流紅線
    # ==========================================
    {
        "id": "G10-NEWTON-FACTOR",
        "grade": "G10",
        "name": "牛頓一次因式檢驗法",
        "level": "FATAL",
        "pattern": r"牛頓一次因式|牛頓檢驗法",
        "keywords": ["牛頓一次因式"],
        "desc": "108 課綱高一已刪除「牛頓一次因式檢驗法」，整係數一次因式檢驗改採有理根直接代入檢驗。",
        "suggestion": "提供因式分解提示或設計容易看出公因式的整數根。"
    },
    {
        "id": "G10-LAGRANGE-INTERPOLATION",
        "grade": "G10",
        "name": "拉格朗日插值法",
        "level": "FATAL",
        "pattern": r"拉格朗日插值",
        "keywords": ["拉格朗日插值"],
        "desc": "高一多項式已刪除拉格朗日插值公式，僅保留一般求多項式聯立方程或牛頓插值（若需要）。",
        "suggestion": "以代入點坐標解多項式係數。"
    },
    {
        "id": "G11-MATHB-CROSS-PRODUCT",
        "grade": "G11_MATHB",
        "name": "數 B 考空間向量外積與平面方程式",
        "level": "FATAL",
        "pattern": r"外積|\\times|法向量|平面方程式",
        "keywords": ["向量外積", "平面方程式", "法向量", "三階行列式"],
        "desc": "高中數學 B 嚴格禁止命題「空間向量外積 (a × b)」、「空間平面方程式 ax+by+cz=d」與「三階行列式幾何意義」，此為數 A 專屬內容。",
        "suggestion": "數 B 空間題僅限平面投影、長方體坐標化與兩點距離、兩向量內積及夾角。"
    }
]

# 108 課綱常見國中單元代碼對照表
CURRICULUM_CODE_MAP = {
    # 國七
    "數與數線": {"content": "N-7-1", "performance": "n-IV-1", "competency": "數-J-A1"},
    "整數運算": {"content": "N-7-2", "performance": "n-IV-2", "competency": "數-J-A1"},
    "分數運算": {"content": "N-7-3", "performance": "n-IV-2", "competency": "數-J-A1"},
    "一元一次方程式": {"content": "A-7-1", "performance": "a-IV-1", "competency": "數-J-A2"},
    "二元一次聯立方程式": {"content": "A-7-2", "performance": "a-IV-2", "competency": "數-J-A2"},
    "直角坐標與函數圖形": {"content": "G-7-1", "performance": "g-IV-1", "competency": "數-J-B1"},
    "比與比例式": {"content": "N-7-6", "performance": "n-IV-4", "competency": "數-J-A3"},
    "一元一次不等式": {"content": "A-7-3", "performance": "a-IV-3", "competency": "數-J-A2"},
    # 國八
    "乘法公式與多項式": {"content": "A-8-1", "performance": "a-IV-4", "competency": "數-J-A1"},
    "平方根與畢氏定理": {"content": "N-8-1", "performance": "n-IV-3", "competency": "數-J-B1"},
    "因式分解": {"content": "A-8-2", "performance": "a-IV-4", "competency": "數-J-A1"},
    "一元二次方程式": {"content": "A-8-3", "performance": "a-IV-5", "competency": "數-J-A2"},
    "等差數列與等差級數": {"content": "N-8-3", "performance": "n-IV-5", "competency": "數-J-A1"},
    "線型函數": {"content": "F-8-1", "performance": "f-IV-1", "competency": "數-J-B1"},
    "三角形與多邊形內外角": {"content": "S-8-1", "performance": "s-IV-1", "competency": "數-J-B1"},
    "平行與四邊形": {"content": "S-8-4", "performance": "s-IV-4", "competency": "數-J-B1"},
    # 國九
    "比例線段與相似形": {"content": "S-9-1", "performance": "s-IV-5", "competency": "數-J-B1"},
    "圓形性質與圓周角": {"content": "S-9-2", "performance": "s-IV-6", "competency": "數-J-B1"},
    "幾何證明與推理": {"content": "S-9-3", "performance": "s-IV-7", "competency": "數-J-A2"},
    "外心、內心與重心": {"content": "S-9-4", "performance": "s-IV-8", "competency": "數-J-B1"},
    "二次函數圖形與極值": {"content": "F-9-1", "performance": "f-IV-2", "competency": "數-J-B2"},
    "統計與機率": {"content": "D-9-1", "performance": "d-IV-1", "competency": "數-J-B3"}
}

class CurriculumAuditor:
    @staticmethod
    def audit_question(stem: str, grade: str = "G9", options: Optional[List[str]] = None) -> List[Dict[str, Any]]:
        """針對單一題目檢核是否觸犯 108 課綱超綱紅線"""
        violations = []
        full_text = stem + " " + " ".join(options if options else [])

        for rule in REDLINE_RULES:
            # 年級過濾：若指定 G7，則檢查 G7；若未指定，全卷通用檢驗
            rule_grade = rule["grade"]
            if grade and rule_grade != grade and rule_grade not in ["ALL", "G10_G12"]:
                # 如果是國九，也會繼承國七、國八嚴禁項（例如國九也不可考三次乘法公式或f(x)若未進入函數單元）
                if grade == "G9" and rule_grade in ["G7", "G8"]:
                    pass
                elif grade in ["G7", "G8"] and rule_grade == "G9":
                    continue
                elif grade != rule_grade:
                    continue

            # 關鍵字與正規式檢測
            matched = False
            for kw in rule["keywords"]:
                if kw in full_text:
                    matched = True
                    break

            if not matched and rule.get("pattern"):
                if re.search(rule["pattern"], full_text, re.IGNORECASE):
                    matched = True

            if matched:
                violations.append({
                    "rule_id": rule["id"],
                    "name": rule["name"],
                    "level": rule["level"],
                    "desc": rule["desc"],
                    "suggestion": rule["suggestion"]
                })

        return violations

    @staticmethod
    def infer_curriculum_codes(text: str, default_scope: str = "幾何與證明") -> Dict[str, str]:
        """依題目文字自動推論 108 課綱雙向細目代碼"""
        for topic, codes in CURRICULUM_CODE_MAP.items():
            if topic in text or topic in default_scope:
                return codes

        # 針對常見關鍵詞推論
        if any(w in text for w in ["外心", "內心", "重心", "三心"]):
            return {"content": "S-9-4", "performance": "s-IV-8", "competency": "數-J-B1"}
        elif any(w in text for w in ["證明", "已知", "求證", "全等"]):
            return {"content": "S-9-3", "performance": "s-IV-7", "competency": "數-J-A2"}
        elif any(w in text for w in ["圓", "半徑", "切線", "圓心角", "圓周角"]):
            return {"content": "S-9-2", "performance": "s-IV-6", "competency": "數-J-B1"}
        elif any(w in text for w in ["相似", "比例線段"]):
            return {"content": "S-9-1", "performance": "s-IV-5", "competency": "數-J-B1"}
        elif any(w in text for w in ["二次函數", "拋物線", "對稱軸"]):
            return {"content": "F-9-1", "performance": "f-IV-2", "competency": "數-J-B2"}
        elif any(w in text for w in ["等差", "公差"]):
            return {"content": "N-8-3", "performance": "n-IV-5", "competency": "數-J-A1"}
        elif any(w in text for w in ["畢氏", "勾股", "直角三角形"]):
            return {"content": "N-8-1", "performance": "n-IV-3", "competency": "數-J-B1"}

        # 預設落點
        return {"content": "S-9-3", "performance": "s-IV-7", "competency": "數-J-A1"}
