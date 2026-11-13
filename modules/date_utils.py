# -*- coding: utf-8 -*-
"""
日付解析・文字列正規化・指示書判定・管理No抽出などの共通ユーティリティモジュール。
"""

import calendar
import datetime
import math
import re
import unicodedata


class AnalysisCancelledException(Exception):
    """ユーザーによる解析中断時に送出される例外。"""


def is_na(val) -> bool:
    """None / NaN / NaT / 空文字系欠損値を統一判定する。"""
    if val is None:
        return True
    if isinstance(val, float) and math.isnan(val):
        return True
    # pandas.NaT 等の欠損型を pandas 非依存で安全に判定
    cls_name = type(val).__name__
    if cls_name in ("NaTType", "NAType"):
        return True
    if isinstance(val, str):
        s = val.strip().lower()
        return s in ("", "nan", "nat", "none")
    return False


def add_months(dt_date: datetime.date, months: int) -> datetime.date:
    """指定日付に月数を加減算する（月末日は対象月の末日に丸める）。"""
    if isinstance(dt_date, datetime.datetime):
        dt_date = dt_date.date()
    zero_based_month = (dt_date.month - 1) + months
    year = dt_date.year + zero_based_month // 12
    month = zero_based_month % 12 + 1
    last_day = calendar.monthrange(year, month)[1]
    day = min(dt_date.day, last_day)
    return datetime.date(year, month, day)


def is_order_date_expired(
    order_date_val,
    base_date: datetime.date | None = None,
    exclude_months: int = 0,
) -> bool:
    """発注日が基準日 (base_date) から exclude_months ヶ月以上経過しているかを判定する。"""
    # 万一旧シグネチャ順 (order_date_val, months_int, today_ref) で呼ばれた場合の互換補正
    if isinstance(base_date, (int, float)) and not isinstance(base_date, bool):
        actual_months = int(base_date)
        actual_base = exclude_months if isinstance(exclude_months, (datetime.date, datetime.datetime)) else datetime.date.today()
    else:
        actual_base = base_date or datetime.date.today()
        try:
            actual_months = int(exclude_months)
        except Exception:
            return False

    if is_na(order_date_val) or actual_months <= 0:
        return False

    parsed_date: datetime.date | None = None
    if isinstance(order_date_val, datetime.datetime):
        parsed_date = order_date_val.date()
    elif isinstance(order_date_val, datetime.date):
        parsed_date = order_date_val
    else:
        parsed_date = parse_flexible_date(order_date_val)

    if parsed_date is None:
        return False

    if isinstance(actual_base, datetime.datetime):
        actual_base = actual_base.date()
    cutoff_date = add_months(parsed_date, actual_months)
    return actual_base >= cutoff_date


# --- 指示書列（管理表 J列）の判定 ---

_INST_EXACT_NONE = frozenset(
    ("-", "ー", "―", "‐", "–", "−", "NO", "NONE", "N/A", "NA", "0", "FALSE", "X")
)
_INST_SUBSTR_NONE = ("無", "なし", "ナシ", "不要", "対象外", "否", "×", "✕", "✗", "✖")
_INST_EXACT_HAS = frozenset(("YES", "TRUE", "1", "OK", "O", "要"))
_INST_SUBSTR_HAS = ("有", "あり", "アリ", "○", "〇", "◯", "◎", "済", "完", "指示")


def classify_instruction_col(raw_val) -> str:
    """
    管理表の指示書列の値を 'HAS' / 'NONE' / 'BLANK' の3状態に分類する。
    - BLANK: 空欄・未記入・空白系文字列・None
    - NONE: ハイフンや「なし」「無」「不要」「×」などの否定・無指示表記
            (FAX, No.1, 10 等の英数中の文字を誤検出しないよう英字記号は完全一致で判定)
    - HAS: 「あり」「有」「○」「済」や指示書番号・ファイル参照など実質的な指示記述
    """
    if raw_val is None or (isinstance(raw_val, float) and math.isnan(raw_val)):
        return "BLANK"
    if type(raw_val).__name__ in ("NaTType", "NAType"):
        return "BLANK"

    norm = unicodedata.normalize("NFKC", str(raw_val)).strip()
    if not norm:
        return "BLANK"

    lower = norm.lower()
    if lower in ("nan", "nat", "空白", "(空白)", "空欄", "(空欄)", "未記入", "未入力"):
        return "BLANK"

    upper = norm.upper()
    # 事前分類済みステータスが直接渡された場合も維持
    if upper in ("HAS", "NONE", "BLANK"):
        return upper

    # 完全一致での「なし」判定 ("-", "ー", "NO", "NONE", "N/A", "0", "FALSE", "X")
    if upper in _INST_EXACT_NONE:
        return "NONE"

    # 日本語明確否定語・バツ記号の部分一致判定
    for tok in _INST_SUBSTR_NONE:
        if tok in norm:
            return "NONE"

    # 完全一致・部分一致での「あり」判定
    if upper in _INST_EXACT_HAS:
        return "HAS"
    for tok in _INST_SUBSTR_HAS:
        if tok in norm:
            return "HAS"

    # 上記の空白・なし系に該当しない文字（指示書番号・記号等）が入っている場合は「あり」とみなす
    return "HAS"


def is_instruction_matched(raw_val_or_state, filter_mode: str) -> bool:
    """
    指示書フィルター条件に合致するかを判定する。
    constants.INSTRUCTION_FILTER_OPTIONS のコード値:
      - 'HAS_ONLY'     : 「有り」のみ
      - 'HAS_OR_BLANK' : 有り または 空白 のみ
      - 'BLANK_ONLY'   : 「空白」のみ
      - 'ALL'          : すべて表示 (フィルタなし)
    および対応するラベル文字列・状態コードの両方に一貫して対応する。
    """
    raw_mode = unicodedata.normalize("NFKC", str(filter_mode or "")).strip()
    upper_mode = raw_mode.upper()

    if not upper_mode or upper_mode == "ALL" or raw_mode in ("すべて", "全て", "すべて表示 (フィルタなし)"):
        return True

    state = classify_instruction_col(raw_val_or_state)

    if upper_mode in ("HAS_ONLY", "HAS") or raw_mode in ("「有り」のみ", "有りのみ", "指示書ありのみ", "指示書あり", "ありのみ"):
        return state == "HAS"
    if upper_mode in ("HAS_OR_BLANK",) or ("有り" in raw_mode and "空白" in raw_mode):
        return state in ("HAS", "BLANK")
    if upper_mode in ("BLANK_ONLY", "BLANK") or raw_mode in ("「空白」のみ", "空白のみ"):
        return state == "BLANK"
    if upper_mode in ("NONE_ONLY", "NONE") or raw_mode in ("「無し」のみ", "なしのみ"):
        return state == "NONE"
    if upper_mode == "NONE_OR_BLANK" or ("なし" in raw_mode and "フィルタなし" not in raw_mode):
        return state in ("NONE", "BLANK")

    return state == "HAS"


# --- 柔軟な日付パーサ ---

_UNIT_FOLLOWED_RE = re.compile(
    r"^\s*\d+\s*/\s*\d+\s*(?:個|本|枚|台|式|セット|set|箱|袋|巻|缶|瓶|組|対|足|着|冊|部|件|回|人|名|m|mm|cm|km|kg|g|mg|l|ml|ℓ|㎡)\b",
    re.IGNORECASE,
)


def _safe_make_date(year: int, month: int, day: int) -> datetime.date | None:
    """妥当範囲 (1990〜2100) かつ実在するカレンダー日付を生成する。"""
    if not (1990 <= year <= 2100 and 1 <= month <= 12):
        return None
    last_day = calendar.monthrange(year, month)[1]
    if not (1 <= day <= last_day):
        return None
    try:
        return datetime.date(year, month, day)
    except ValueError:
        return None


def _resolve_cross_year(
    month: int,
    default_year: int | None,
    ref_month: int | None,
) -> int:
    """年が省略された月・月日について、発注基準月から年跨ぎを自動判定する。"""
    today = datetime.date.today()
    base_year = default_year if default_year and 1990 <= default_year <= 2100 else today.year
    base_month = ref_month if ref_month and 1 <= ref_month <= 12 else today.month

    if base_month >= 11 and month <= 2:
        return base_year + 1
    if base_month <= 2 and month >= 11:
        return base_year - 1
    return base_year


def parse_flexible_date(
    val,
    default_year: int | None = None,
    ref_month: int | None = None,
    allow_period: bool = True,
) -> datetime.date | None:
    """
    Excelセル値・日付文字列・日本語曖昧表現（前半/後半/頭/上旬/中旬/下旬/月末/月のみ 等）
    および期間範囲（11/10〜11/15 等）を柔軟に datetime.date へ変換する。
    allow_period=False の場合は月のみ・曖昧時期表現・範囲表現を除外し確定日付のみ返す。
    数量・分数（例: '1/2個'）や型番などは誤変換しないよう厳格にガードする。
    """
    if is_na(val):
        return None

    if isinstance(val, datetime.datetime):
        return val.date()
    if isinstance(val, datetime.date):
        return val

    # Excel シリアル値の数値入力に対応（範囲外の数値 1, 2, 100 等は日付と誤認しない）
    if isinstance(val, (int, float)) and not isinstance(val, bool):
        num = float(val)
        if 32874 <= num <= 73050:  # 1990-01-01 〜 2099-12-31
            try:
                base = datetime.date(1899, 12, 30)
                return base + datetime.timedelta(days=int(num))
            except Exception:
                return None
        return None

    raw_s = unicodedata.normalize("NFKC", str(val)).strip()
    if not raw_s:
        return None

    # 明確に日付ではない状態・未入力語句
    low = raw_s.lower()
    if low in ("-", "ー", "―", "‐", "未", "未定", "未確認", "未着", "未入力", "未済", "随時", "別途", "tbd"):
        return None

    # '1/2個', '3/4本' などの分数+単位表現を除外
    if _UNIT_FOLLOWED_RE.match(raw_s):
        return None

    # 曜日表記・補足カッコ ('(水)', '（木）', '頃', '予定', '必着', '着', '納', '出荷', 'AM', 'PM' など) を整理
    cleaned = re.sub(r"[\(（][月火水木金土日祝][\)）]", " ", raw_s)
    # 範囲表記 (〜, ~, から, －等) が複数日付を繋いでいる場合、後半(納期の後ろ側)を優先採用
    range_parts = re.split(r"[〜～~]|から", cleaned)
    if len(range_parts) >= 2:
        if not allow_period:
            return None
        left_part = range_parts[0].strip()
        right_part = range_parts[-1].strip()
        d_right = parse_flexible_date(
            right_part,
            default_year=default_year,
            ref_month=ref_month,
            allow_period=allow_period,
        )
        if d_right is not None:
            return d_right
        m_left_ym = re.search(r"(?:(\d{2,4})\s*[年/\-\.])?\s*(\d{1,2})\s*[月/\-\.]\s*\d{1,2}", left_part)
        m_right_d = re.search(r"^\s*(\d{1,2})\s*日?", right_part)
        if m_left_ym and m_right_d:
            y_raw = m_left_ym.group(1)
            m_val = int(m_left_ym.group(2))
            d_val = int(m_right_d.group(1))
            if y_raw:
                y_val = int(y_raw)
                if y_val < 100:
                    y_val += 2000
            else:
                y_val = _resolve_cross_year(m_val, default_year, ref_month)
            dt = _safe_make_date(y_val, m_val, d_val)
            if dt is not None:
                return dt

    # "11/10-15" のように M/D-D 形式で後ろに日が続くケース
    m_md_range = re.search(
        r"(?:(\d{4}|\d{2})\s*[年/\-\.])?\s*(\d{1,2})\s*[月/\.]\s*(\d{1,2})\s*[-‐−]\s*(\d{1,2})(?!\d)",
        cleaned,
    )
    if m_md_range:
        if not allow_period:
            return None
        y_raw, m_str, _, d2_str = m_md_range.groups()
        m_val, d2_val = int(m_str), int(d2_str)
        if y_raw:
            y_val = int(y_raw) + (2000 if len(y_raw) == 2 else 0)
        else:
            y_val = _resolve_cross_year(m_val, default_year, ref_month)
        dt = _safe_make_date(y_val, m_val, d2_val)
        if dt is not None:
            return dt

    # 1. 和暦 (R8.11.15, 令和8年11月15日 など)
    m_wa = re.search(
        r"(?:令和|R)\s*(\d{1,2})\s*[年/\-\.]\s*(\d{1,2})\s*[月/\-\.]\s*(\d{1,2})\s*日?",
        cleaned,
        re.IGNORECASE,
    )
    if m_wa:
        ry, m_val, d_val = int(m_wa.group(1)), int(m_wa.group(2)), int(m_wa.group(3))
        dt = _safe_make_date(2018 + ry, m_val, d_val)
        if dt is not None:
            return dt

    # 2. 4桁年 + 月 + 日 (YYYY/M/D, YYYY-M-D, YYYY.M.D, YYYY年M月D日)
    m_ymd4 = re.search(
        r"(?<!\d)(\d{4})\s*[年/\-\.]\s*(\d{1,2})\s*[月/\-\.]\s*(\d{1,2})(?:\s*日)?(?!\d)",
        cleaned,
    )
    if m_ymd4:
        y_val, m_val, d_val = int(m_ymd4.group(1)), int(m_ymd4.group(2)), int(m_ymd4.group(3))
        dt = _safe_make_date(y_val, m_val, d_val)
        if dt is not None:
            return dt

    # 3. 2桁年 + 月 + 日 (YY/M/D, YY-M-D, YY.M.D, YY年M月D日)
    m_ymd2 = re.search(
        r"(?<!\d)(\d{2})\s*[年/\-\.]\s*(\d{1,2})\s*[月/\-\.]\s*(\d{1,2})(?:\s*日)?(?!\d)",
        cleaned,
    )
    if m_ymd2:
        y_val = 2000 + int(m_ymd2.group(1))
        m_val, d_val = int(m_ymd2.group(2)), int(m_ymd2.group(3))
        dt = _safe_make_date(y_val, m_val, d_val)
        if dt is not None:
            return dt

    # 4. 年 + 月 (日なし: YYYY/MM, YYYY年MM月)
    if allow_period:
        m_ym = re.search(
            r"(?<!\d)(\d{4})\s*[年/\-\.]\s*(\d{1,2})\s*月?(?!\s*[/\-\.\d])",
            cleaned,
        )
        if m_ym:
            y_val, m_val = int(m_ym.group(1)), int(m_ym.group(2))
            if 1 <= m_val <= 12:
                last_d = calendar.monthrange(y_val, m_val)[1]
                if any(k in cleaned for k in ("頭", "あたま", "月初")):
                    d_val = 10
                elif any(k in cleaned for k in ("前半", "上旬")):
                    d_val = 15
                elif "中旬" in cleaned:
                    d_val = 20
                else:
                    d_val = last_d
                dt = _safe_make_date(y_val, m_val, d_val)
                if dt is not None:
                    return dt

    # 5. 年省略の月 + 日 (M月D日, M/D, M-D, M.D)
    m_md_kanji = re.search(r"(?<!\d)(\d{1,2})\s*月\s*(\d{1,2})\s*日?(?!\d)", cleaned)
    if m_md_kanji:
        m_val, d_val = int(m_md_kanji.group(1)), int(m_md_kanji.group(2))
        y_val = _resolve_cross_year(m_val, default_year, ref_month)
        dt = _safe_make_date(y_val, m_val, d_val)
        if dt is not None:
            return dt

    m_md_sep = re.search(
        r"^(?:[^\d]*?)(?<!\d)(\d{1,2})\s*[/\.]\s*(\d{1,2})(?!\s*[/\.\d])(?:\s*(?:頃|予定|必着|納|出荷|着|マデ|まで|AM|PM|以降|中|前後|内))?",
        cleaned,
        re.IGNORECASE,
    )
    if m_md_sep:
        tail = cleaned[m_md_sep.end():].strip()
        if not tail or re.match(r"^(?:[\(\s]|頃|予定|必着|納|着|出荷|まで|マデ|以降|前後|中|$)", tail):
            m_val, d_val = int(m_md_sep.group(1)), int(m_md_sep.group(2))
            y_val = _resolve_cross_year(m_val, default_year, ref_month)
            dt = _safe_make_date(y_val, m_val, d_val)
            if dt is not None:
                return dt

    m_md_dash = re.match(
        r"^(\d{1,2})-(\d{1,2})(?:\s*(?:頃|予定|必着|納|着|出荷|まで|マデ))?$",
        cleaned,
    )
    if m_md_dash:
        m_val, d_val = int(m_md_dash.group(1)), int(m_md_dash.group(2))
        y_val = _resolve_cross_year(m_val, default_year, ref_month)
        dt = _safe_make_date(y_val, m_val, d_val)
        if dt is not None:
            return dt

    # 6. 8桁連結数値日付 (YYYYMMDD) 例: 20261115
    m_8 = re.match(r"^(20\d{2})(0[1-9]|1[0-2])(0[1-9]|[12]\d|3[01])$", cleaned)
    if m_8:
        dt = _safe_make_date(int(m_8.group(1)), int(m_8.group(2)), int(m_8.group(3)))
        if dt is not None:
            return dt

    # 7. 日本語月のみ / 時期表現 ('11月', '11月予定', '11月前半', '11月後半', '11月頭', '11月いっぱい', '11月最終' 等)
    if allow_period:
        m_month_only = re.search(
            r"(?<!\d)(\d{1,2})\s*月(?:\s*(?:前半|後半|上旬|中旬|下旬|頭|あたま|月初|末|月末|いっぱい|最終|中|頃|予定|納|出荷|着|内))?",
            cleaned,
        )
        if m_month_only:
            m_val = int(m_month_only.group(1))
            if 1 <= m_val <= 12:
                y_val = _resolve_cross_year(m_val, default_year, ref_month)
                last_d = calendar.monthrange(y_val, m_val)[1]
                if any(k in cleaned for k in ("頭", "あたま", "月初")):
                    d_val = 10
                elif any(k in cleaned for k in ("前半", "上旬")):
                    d_val = 15
                elif "中旬" in cleaned:
                    d_val = 20
                else:
                    d_val = last_d
                dt = _safe_make_date(y_val, m_val, d_val)
                if dt is not None:
                    return dt

    return None


# --- 現品確認・請求日の入力済み判定 ---

_UNFILLED_TOKENS = frozenset(
    (
        "-",
        "ー",
        "―",
        "‐",
        "–",
        "−",
        "none",
        "nan",
        "nat",
        "未",
        "未確認",
        "未済",
        "未着",
        "未入力",
        "未定",
        "×",
        "✕",
        "✗",
        "✖",
        "なし",
        "無",
        "不要",
        "空欄",
        "(空欄)",
        "(空白)",
        "空白",
        "0",
        "false",
    )
)


def is_filled_confirmation(val) -> bool:
    """現品確認日や請求日が入力済み（完了扱い）かどうかを判定する。"""
    if is_na(val):
        return False
    if isinstance(val, bool):
        return False
    if isinstance(val, (int, float)) and val == 0:
        return False
    if isinstance(val, (datetime.datetime, datetime.date)):
        return True
    s = unicodedata.normalize("NFKC", str(val)).strip()
    if not s:
        return False
    if s.lower() in _UNFILLED_TOKENS:
        return False
    return True


# --- 社名・比較キー正規化 ---

_FULL_CORP_RE = re.compile(
    r"(?:一般社団法人|公益社団法人|一般財団法人|公益財団法人|"
    r"株式会社|有限会社|合同会社|合資会社|合名会社|"
    r"Co\.,?\s*Ltd\.?|Inc\.?|Ltd\.?|Corp\.?)",
    re.IGNORECASE,
)
_BRACKET_CORP_RE = re.compile(r"[\(（\[【]\s*(?:株|有|同|資|名|社|財)\s*[\)）\]】]")
_HONORIFIC_RE = re.compile(r"(?:御中|様|殿)$")


def clean_company_name(s: str) -> str:
    """
    会社名から法人格（「株式会社」「有限会社」「(株)」「(有)」「㈱」「㈲」等）および敬称を安全に除去する。
    「有明」「三株」など社名本体に含まれる漢字一文字の「有」「株」は誤削除しない。
    """
    if not s:
        return ""
    # NFKC前の旧環境依存文字を除去（NFKC後は(株)/(有)になるためどちらでも対応）
    raw = str(s).replace("㈱", "").replace("㈲", "").replace("㈳", "").replace("㈵", "")
    norm = unicodedata.normalize("NFKC", raw).strip()
    norm = _BRACKET_CORP_RE.sub("", norm)
    norm = _FULL_CORP_RE.sub("", norm)
    norm = norm.strip()
    norm = _HONORIFIC_RE.sub("", norm).strip()
    # 末尾に全角/半角スペースを挟んで付く法人記号以外の空白も詰めすぎない範囲で整理
    norm = re.sub(r"\s+", " ", norm).strip()
    return norm


def normalize_key(s: str) -> str:
    """型番・品番・会社名などを照合比較用の正規化キーに変換する。"""
    if not s:
        return ""
    raw = str(s).replace("㈱", "").replace("㈲", "").replace("㈳", "").replace("㈵", "")
    norm = unicodedata.normalize("NFKC", raw).strip()
    norm = _BRACKET_CORP_RE.sub("", norm)
    norm = _FULL_CORP_RE.sub("", norm)
    norm = norm.upper()
    norm = norm.replace("×", "X").replace("✕", "X").replace("✖", "X")
    # 空白および照合上の区切り記号を除去
    norm = re.sub(r"[\s\-_ー―‐/\\.,、。:：;；・\(\)（）\[\]【】\"'`]+", "", norm)
    return norm


def make_bigrams(s: str) -> set[str]:
    """文字列から Dice / Jaccard 計算用の 2-gram 集合を作る。"""
    if not s:
        return set()
    if len(s) == 1:
        return {s}
    return {s[i : i + 2] for i in range(len(s) - 1)}


def fast_similarity(
    s1: str,
    s2: str,
    bg1: set[str] | None = None,
    bg2: set[str] | None = None,
) -> float:
    """2つの文字列（または事前計算済み 2-gram 集合）の高速類似度 (0.0〜1.0) を算出する。"""
    if bg1 is not None and bg2 is not None:
        k1 = str(s1 or "")
        k2 = str(s2 or "")
        b1, b2 = bg1, bg2
    else:
        k1 = normalize_key(s1) if s1 else ""
        k2 = normalize_key(s2) if s2 else ""
        b1 = make_bigrams(k1)
        b2 = make_bigrams(k2)

    if not k1 or not k2:
        return 0.0
    if k1 == k2:
        return 1.0
    if k1 in k2 or k2 in k1:
        shorter, longer = min(len(k1), len(k2)), max(len(k1), len(k2))
        if shorter >= 3:
            return max(0.85, shorter / longer)
    if not b1 or not b2:
        return 0.0
    inter = len(b1 & b2)
    return (2.0 * inter) / (len(b1) + len(b2))


# --- 数量・セル文字列・管理No ---


def extract_leading_number(val) -> float | None:
    """先頭または数量文字列から数値を取り出す。取れない場合は None。"""
    if is_na(val):
        return None
    if isinstance(val, (int, float)) and not isinstance(val, bool):
        return float(val)
    s = unicodedata.normalize("NFKC", str(val)).strip().replace(",", "")
    m = re.search(r"[-+]?\d+(?:\.\d+)?", s)
    if not m:
        return None
    try:
        return float(m.group(0))
    except ValueError:
        return None


def is_zero_quantity(val) -> bool:
    """数量が実質 0（'0', '0個', 0.0 など）かを判定する。空欄は False。"""
    num = extract_leading_number(val)
    if num is None:
        return False
    return abs(num) < 1e-9


def clean_cell_str(val) -> str:
    """Excel セル値を UI・照合用の表示文字列へ安全に変換する（浮動小数 '.0' の不要末尾除去）。"""
    if is_na(val):
        return ""
    if isinstance(val, datetime.datetime):
        return val.strftime("%Y-%m-%d")
    if isinstance(val, datetime.date):
        return val.strftime("%Y-%m-%d")
    if isinstance(val, float):
        if val.is_integer():
            return str(int(val))
        return str(val).rstrip("0").rstrip(".")
    s = str(val).strip()
    if s.lower() in ("nan", "nat", "none"):
        return ""
    return s


_EIGHT_DIGIT_RE = re.compile(r"(?<!\d)(\d{8})(?!\d)")


def extract_all_8digit_nos(text: str) -> list[str]:
    """テキスト内に含まれるすべての独立した8桁数字（前後に数字が続かない8桁）を順序・重複排除で返す。"""
    if not text:
        return []
    norm = unicodedata.normalize("NFKC", str(text))
    found = _EIGHT_DIGIT_RE.findall(norm)
    result: list[str] = []
    seen: set[str] = set()
    for d in found:
        if d not in seen:
            seen.add(d)
            result.append(d)
    return result


def extract_8digit_no(text: str) -> str:
    """文字列から先頭の独立8桁数字を抽出する。見つからない場合は空文字。"""
    cands = extract_all_8digit_nos(text)
    return cands[0] if cands else ""


def parse_lines_to_order_nos(raw_text: str) -> list[str]:
    """
    テキスト（改行・カンマ・全角読点・セミコロン・タブ・スペース区切り）から
    除外対象の管理Noリストを抽出・正規化する。
    8桁数字が含まれる場合は8桁を抜き出し、接頭英字やその他形式でもトークン全体を維持可能とする。
    """
    if not raw_text:
        return []
    norm = unicodedata.normalize("NFKC", str(raw_text))
    tokens = re.split(r"[\r\n,、;\t\s]+", norm)
    result: list[str] = []
    seen: set[str] = set()
    for tok in tokens:
        t = tok.strip()
        if not t:
            continue
        digits = extract_all_8digit_nos(t)
        if digits:
            for d in digits:
                if d not in seen:
                    seen.add(d)
                    result.append(d)
        else:
            key = t.upper()
            if key not in seen:
                seen.add(key)
                result.append(key)
    return result


def is_order_no_excluded(order_no: str, excluded_nos) -> bool:
    """指定管理Noが除外リスト（8桁または正規化文字列）に一致するか判定する。"""
    if not order_no or not excluded_nos:
        return False
    excl_set = {unicodedata.normalize("NFKC", str(x)).strip().upper() for x in excluded_nos if str(x).strip()}
    if not excl_set:
        return False

    norm_no = unicodedata.normalize("NFKC", str(order_no)).strip().upper()
    if norm_no in excl_set:
        return True

    for d8 in extract_all_8digit_nos(norm_no):
        if d8 in excl_set:
            return True

    norm_compact = normalize_key(norm_no)
    if norm_compact and norm_compact in excl_set:
        return True
    return False


def get_record_display_value(rec: dict, col_id: str) -> str:
    """
    レコードと列IDからテーブル表示・列フィルター・検索で一貫した表示文字列を取得する。
    order_date / due_date など特殊列の表示値（'-' 等）とフィルタ値との食い違いを防止する。
    """
    if not isinstance(rec, dict):
        return ""
    if col_id == "order_date":
        val = rec.get("order_date_str")
        if val:
            return str(val)
        od = rec.get("order_date")
        if isinstance(od, (datetime.datetime, datetime.date)):
            return od.strftime("%Y-%m-%d")
        return "-"
    if col_id == "due_date":
        val = rec.get("due_date_str")
        if val:
            return str(val)
        raw = rec.get("due_date")
        if isinstance(raw, (datetime.datetime, datetime.date)):
            return raw.strftime("%Y-%m-%d")
        return str(raw) if raw not in (None, "") else ""
    v = rec.get(col_id, "")
    if v is None:
        return ""
    if isinstance(v, (datetime.datetime, datetime.date)):
        return v.strftime("%Y-%m-%d")
    return str(v)


def rgb_to_bgr_int(*args) -> int:
    """RGB成分（'#RRGGBB' 文字列 または r, g, b の3整数）から Excel COM API 用の BGR 24bit 整数を返す。"""
    if len(args) == 1:
        val = str(args[0]).strip().lstrip("#")
        if len(val) == 3:
            val = "".join(ch * 2 for ch in val)
        if len(val) == 6:
            r = int(val[0:2], 16)
            g = int(val[2:4], 16)
            b = int(val[4:6], 16)
            return (b << 16) | (g << 8) | r
        raise ValueError(f"Invalid hex color: {args[0]}")
    if len(args) == 3:
        r, g, b = args
        return (int(b) << 16) | (int(g) << 8) | int(r)
    raise TypeError(f"rgb_to_bgr_int expected 1 or 3 arguments, got {len(args)}")
