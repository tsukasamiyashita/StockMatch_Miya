"""
StockMatch_Miya 日付・文字列処理ユーティリティモジュール
日付パース、曖昧期間変換、文字列正規化、類似度計算、除外判定などの純粋ロジックを集約
"""

import re
import calendar
import datetime
import unicodedata
import pandas as pd


class AnalysisCancelledException(Exception):
    """ユーザーによって照合処理が中断された場合のカスタム例外"""
    pass


def add_months(d: datetime.date, months: int) -> datetime.date:
    """指定された日付 d に months ヶ月を加算した日付を返す（月末日補正あり）"""
    total_months = (d.year * 12 + (d.month - 1)) + months
    target_year = total_months // 12
    target_month = (total_months % 12) + 1
    max_day = calendar.monthrange(target_year, target_month)[1]
    return datetime.date(target_year, target_month, min(d.day, max_day))


def is_order_date_expired(
    order_date: datetime.date | None,
    base_date: datetime.date | None = None,
    exclude_months: int | None = 12,
) -> bool:
    """注文日(order_date)から指定ヶ月数(exclude_months)以上が経過しているかを判定"""
    if order_date is None:
        return False
    if exclude_months is None or exclude_months <= 0:
        return False
    if base_date is None:
        base_date = datetime.date.today()
    cutoff_date = add_months(order_date, exclude_months)
    return base_date >= cutoff_date


def classify_instruction_col(raw_val: str) -> str:
    """
    注文品管理表の指示書列(J列)の値を判定する:
      - "BLANK": 空白（未記入）
      - "HAS": 「有り」「有」「あり」「○」「〇」「丸」「Yes」「1」等
      - "NONE": 「無し」「無」「なし」「不要」「-」「ー」「×」「No」「0」等
      - "OTHER": その他の文字記載がある場合は基本「HAS (有り)」扱い
    """
    s = unicodedata.normalize("NFKC", str(raw_val or "")).strip().lower()
    if not s or s in ("-", "ー", "―", "‐", "none", "nan", "nat", "(空白)", "空白", "未記入"):
        return "BLANK"

    has_tokens = ("有", "あり", "アリ", "○", "〇", "◯", "丸", "yes", "true", "1", "済", "要", "指示")
    none_tokens = ("無", "なし", "ナシ", "不要", "×", "x", "no", "false", "0", "非", "否")

    for tok in none_tokens:
        if tok in s:
            return "NONE"

    for tok in has_tokens:
        if tok in s:
            return "HAS"

    return "HAS"


def is_instruction_matched(instruction_raw: str, instruction_mode: str) -> bool:
    """
    選択されたフィルターモード (HAS_ONLY / HAS_OR_BLANK / BLANK_ONLY / ALL) に合致するか判定
    """
    if instruction_mode == "ALL":
        return True
    state = classify_instruction_col(instruction_raw)
    if instruction_mode == "HAS_ONLY":
        return state == "HAS"
    elif instruction_mode == "HAS_OR_BLANK":
        return state in ("HAS", "BLANK")
    elif instruction_mode == "BLANK_ONLY":
        return state == "BLANK"
    return True


def parse_flexible_date(
    val, default_year: int | None = None, allow_period: bool = True
) -> datetime.date | None:
    """
    Excelの生データ（日付型、数値シリアル、文字列）から柔軟かつ高精度に datetime.date を復元する。
    上旬・中旬・下旬・月末、和暦表記（令和/R）、年末年始の年跨ぎ、月日のみ表記を自動補完。
    """
    if val is None or pd.isna(val):
        return None

    if isinstance(val, datetime.datetime):
        return val.date()
    if isinstance(val, datetime.date):
        return val

    if isinstance(val, (int, float)):
        if 35000 <= val <= 60000:
            try:
                return datetime.date(1899, 12, 30) + datetime.timedelta(days=int(val))
            except Exception:
                pass

    s = str(val).strip()
    if not s:
        return None

    s = unicodedata.normalize("NFKC", s).strip()
    if not s or s.lower() in (
        "-",
        "ー",
        "―",
        "‐",
        "未",
        "未定",
        "未確認",
        "未着",
        "別途",
        "相談",
        "要相談",
        "調整中",
        "none",
        "nan",
        "nat",
    ):
        return None

    def _make_valid_date(y: int, m: int, d: int) -> datetime.date | None:
        if not (1990 <= y <= 2099 and 1 <= m <= 12 and 1 <= d <= 31):
            return None
        max_d = calendar.monthrange(y, m)[1]
        return datetime.date(y, m, min(d, max_d))

    def _period_to_day(y: int, m: int, period_token: str) -> int | None:
        if not (1990 <= y <= 2099 and 1 <= m <= 12):
            return None
        p = period_token.strip()
        if any(k in p for k in ("下旬", "下", "末", "いっぱい", "一杯", "月内", "後半", "最終")):
            return calendar.monthrange(y, m)[1]
        if any(k in p for k in ("中旬", "中頃", "中ごろ", "中", "半ば", "なかば", "央")):
            return 20
        if "前半" in p:
            return 15
        if any(k in p for k in ("上旬", "上", "初旬", "初", "はじめ", "頭", "あたま")):
            return 10
        return None

    # 曜日表記 (月), (火) 等を除去
    s = re.sub(r"\([月火水木金土日祝・]+\)", " ", s)
    s = s.replace("(", " ").replace(")", " ")

    # 和暦（令和/R・平成/H）を西暦に変換
    s = re.sub(r"令和\s*元\s*年", "2019年", s)
    m_reiwa = re.search(r"令和\s*(\d{1,2})\s*年", s)
    if m_reiwa:
        ry = int(m_reiwa.group(1))
        s = s[: m_reiwa.start()] + f"{2018 + ry}年" + s[m_reiwa.end() :]
    m_r = re.search(r"(?:^|[^a-zA-Z0-9])(?:r|R)\s*(\d{1,2})[\.\/\-年]", s)
    if m_r:
        ry = int(m_r.group(1))
        start_idx = m_r.start() if s[m_r.start()].upper() == "R" else m_r.start() + 1
        s = s[:start_idx] + f"{2018 + ry}-" + s[m_r.end() :]
    m_heisei = re.search(r"平成\s*(\d{1,2})\s*年", s)
    if m_heisei:
        hy = int(m_heisei.group(1))
        s = s[: m_heisei.start()] + f"{1988 + hy}年" + s[m_heisei.end() :]

    # 範囲指定（例: 11/10〜11/15, 11/上〜中）は終了側の期日を採用
    if "〜" in s or "~" in s or "–" in s:
        parts = re.split(r"[〜~–]", s)
        if len(parts) >= 2 and parts[-1].strip():
            right_part = parts[-1].strip()
            # 右側が日・期間表記のみ（例: 11/10〜15）の場合、左側の年月を結合
            if not re.search(r"[\/\-年]", right_part):
                left_match = re.search(r"(\d{4}[年\/\-])?(\d{1,2})[月\/\-]", parts[0])
                if left_match:
                    prefix = left_match.group(0)
                    right_part = prefix + right_part
            s = right_part

    # 1. 完全表記: YYYY年MM月DD日 または YYYY/MM/DD
    m_full = re.search(r"(20\d{2})[\.\/\-年](\d{1,2})[\.\/\-月](\d{1,2})", s)
    if m_full:
        y, m, d = int(m_full.group(1)), int(m_full.group(2)), int(m_full.group(3))
        dt = _make_valid_date(y, m, d)
        if dt:
            return dt

    # 2. 年＋月＋旬・期間表記: YYYY年MM月(下旬|末等)
    if allow_period:
        m_ym_period = re.search(
            r"(20\d{2})[\.\/\-年](\d{1,2})[\.\/\-月]?\s*([上中下初末半][旬半ば頃ごろ末日内杯いっぱい前半後半最終]*)",
            s,
        )
        if m_ym_period:
            y, m = int(m_ym_period.group(1)), int(m_ym_period.group(2))
            token = m_ym_period.group(3)
            d = _period_to_day(y, m, token)
            if d:
                dt = _make_valid_date(y, m, d)
                if dt:
                    return dt

    cur_year = default_year if default_year else datetime.date.today().year

    # 3. 月日表記: MM/DD または MM月DD日
    m_md = re.search(r"(?:^|[^\d])(\d{1,2})[\.\/\-月](\d{1,2})(?:日|$|[^\d])", s)
    if m_md:
        m, d = int(m_md.group(1)), int(m_md.group(2))
        if 1 <= m <= 12 and 1 <= d <= 31:
            # 年末年始の年跨ぎ推論
            y = cur_year
            if default_year:
                ref_month = 11
                if ref_month >= 11 and m <= 2:
                    y = default_year + 1
                elif ref_month <= 2 and m >= 11:
                    y = default_year - 1
            dt = _make_valid_date(y, m, d)
            if dt:
                return dt

    # 4. 月＋旬・期間表記: MM月(上旬|中旬|下旬|月末等)
    if allow_period:
        m_m_period = re.search(
            r"(?:^|[^\d])(\d{1,2})[\.\/\-月]\s*([上中下初末半][旬半ば頃ごろ末日内杯いっぱい前半後半最終]*)",
            s,
        )
        if m_m_period:
            m = int(m_m_period.group(1))
            token = m_m_period.group(2)
            if 1 <= m <= 12:
                y = cur_year
                if default_year and default_year >= 11 and m <= 2:
                    y = default_year + 1
                d = _period_to_day(y, m, token)
                if d:
                    dt = _make_valid_date(y, m, d)
                    if dt:
                        return dt

    # 5. 年月のみ表記: YYYY年MM月 または YYYY/MM
    if allow_period:
        m_ym_only = re.search(r"(20\d{2})[\.\/\-年](\d{1,2})月?$", s)
        if m_ym_only:
            y, m = int(m_ym_only.group(1)), int(m_ym_only.group(2))
            if 1 <= m <= 12:
                max_d = calendar.monthrange(y, m)[1]
                return datetime.date(y, m, max_d)

    # 6. 8桁連続数値: YYYYMMDD
    m_compact = re.search(r"(20\d{2})(0[1-9]|1[0-2])(0[1-9]|[12]\d|3[01])", s)
    if m_compact:
        y, m, d = int(m_compact.group(1)), int(m_compact.group(2)), int(m_compact.group(3))
        dt = _make_valid_date(y, m, d)
        if dt:
            return dt

    return None


def is_filled_confirmation(val) -> bool:
    """管理表の現品確認日・請求日に文字やマーク（済・○・確認等）が入力されているか判定"""
    if val is None:
        return False
    if isinstance(val, (datetime.date, datetime.datetime)):
        return True
    s = unicodedata.normalize("NFKC", str(val)).strip()
    if not s or s.lower() in ("-", "ー", "―", "‐", "none", "nan", "nat", "未", "未確認", "未済", "未着"):
        return False
    return True


def normalize_key(s: str) -> str:
    """照合用のキー文字列正規化（全角半角統一、記号除去、大文字化）"""
    if not s:
        return ""
    s_norm = unicodedata.normalize("NFKC", str(s)).upper()
    s_norm = re.sub(r"[\s\-_/\\:：・()（）\[\]【】「」]+", "", s_norm)
    return s_norm


def clean_company_name(s: str) -> str:
    """会社名・業者名から（株）や株式会社などの法人格を除去して正規化"""
    if not s:
        return ""
    norm = unicodedata.normalize("NFKC", str(s)).upper()
    norm = re.sub(r"[\(（]?(?:株|株式|有限|合同|有)[\)）]?", "", norm)
    norm = re.sub(r"(?:株式会社|有限会社|合同会社|合資会社)", "", norm)
    return normalize_key(norm)


def make_bigrams(s: str) -> set[str]:
    """2文字ずつのバイグラム集合を作成"""
    if not s:
        return set()
    if len(s) == 1:
        return {s}
    return {s[i : i + 2] for i in range(len(s) - 1)}


def fast_similarity(s1: str, s2: str, bg1: set[str], bg2: set[str]) -> float:
    """Jaccard類似度を用いた高速文字類似度計算"""
    if not s1 or not s2:
        return 0.0
    if s1 == s2:
        return 1.0
    if s1 in s2 or s2 in s1:
        return min(len(s1), len(s2)) / max(len(s1), len(s2))
    if not bg1 or not bg2:
        return 0.0
    intersection_len = len(bg1 & bg2)
    union_len = len(bg1 | bg2)
    return intersection_len / union_len if union_len else 0.0


def extract_leading_number(s: str) -> str:
    """文字列の先頭にある数値を抽出（数量比較用）"""
    if not s:
        return ""
    s_norm = unicodedata.normalize("NFKC", str(s)).strip()
    m = re.match(r"^(\d+(?:\.\d+)?)", s_norm)
    return m.group(1) if m else ""


def is_zero_quantity(s: str) -> bool:
    """数量が 0（または 0個・0台 等）であるかを判定"""
    if not s:
        return False
    norm = unicodedata.normalize("NFKC", str(s)).strip()
    if not norm or norm in ("-", "ー", "―", "‐"):
        return False
    if norm == "0":
        return True
    m = re.match(r"^0(?:\.0+)?(?:\s*([個本台枚組個組set式箱袋巻包]|$))", norm, re.IGNORECASE)
    return bool(m)


def clean_cell_str(val) -> str:
    """セルの値を文字列化し、不要な空白や改行を除去"""
    if val is None or pd.isna(val):
        return ""
    s = str(val).strip()
    return "" if s.lower() in ("none", "nan", "nat") else s


def extract_8digit_no(val) -> str | None:
    """文字列から8桁の管理No（注文番号）を抽出"""
    if not val:
        return None
    s = unicodedata.normalize("NFKC", str(val)).strip()
    m = re.search(r"(?<!\d)(\d{8})(?!\d)", s)
    return m.group(1) if m else None


def parse_lines_to_order_nos(raw_text: str) -> list[str]:
    """テキスト（改行・カンマ・タブ区切り）から有効な管理Noリストを抽出・重複排除"""
    if not raw_text:
        return []
    result = []
    seen = set()
    tokens = re.split(r"[\r\n,;\t]+", raw_text)
    for tok in tokens:
        cand = unicodedata.normalize("NFKC", tok).strip()
        if not cand:
            continue
        no8 = extract_8digit_no(cand)
        final_val = no8 if no8 else cand
        if final_val and final_val not in seen:
            seen.add(final_val)
            result.append(final_val)
    return result


def is_order_no_excluded(order_no: str, excluded_nos_set: set[str]) -> bool:
    """指定された管理Noが除外セットに含まれているかを判定"""
    if not excluded_nos_set or not order_no:
        return False
    norm_no = unicodedata.normalize("NFKC", str(order_no)).strip()
    if not norm_no:
        return False
    if norm_no in excluded_nos_set:
        return True
    no8 = extract_8digit_no(norm_no)
    if no8 and no8 in excluded_nos_set:
        return True
    return False


def rgb_to_bgr_int(hex_str: str) -> int:
    """RGB 16進数文字列 (例: '1E293B') を Excel COM 用の BGR 整数に変換"""
    clean_hex = hex_str.strip().lstrip("#")
    if len(clean_hex) != 6:
        return 0
    r = int(clean_hex[0:2], 16)
    g = int(clean_hex[2:4], 16)
    b = int(clean_hex[4:6], 16)
    return (b << 16) | (g << 8) | r
