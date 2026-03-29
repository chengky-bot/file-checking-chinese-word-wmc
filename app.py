# -*- coding: utf-8 -*-
"""
教材/學行報告審核工具

請先安裝相依套件：
    pip install streamlit python-docx

執行方式：
    streamlit run app.py

說明：
- 紅色標記：凡經自動修正的字元（RGB 255,0,0），便於與原文對照。
- 專名號：自訂清單中的名稱套用底線（underline=True）。
- 排位（不可分割詞）：若使用者像測試稿一樣把詞寫成「廣 / 場」「落 / 淚」（空白、斜線斷開），
  會先合併成連續詞再於字間插入 Unicode U+2060（WORD JOINER，零寬不換行字元）。
  WORD JOINER 讓 Word 盡量不在詞內斷行，整詞較易一起移到下一行，避免行尾「斷詞」；
  合併與插入後的整段詞（含 joiner）標為紅色。
- 含圖段落：以私人使用區字元作為「圖片占位符」參與文字規則，寫回時還原內嵌圖 w:r，
  並保留 w:pPr，以降低表格／段落樣式走位；表格本身不刪除，只處理儲存格內文字。
"""

from __future__ import annotations

import io
import os
import re
from copy import deepcopy
from collections import defaultdict
from typing import Any, Callable, Dict, List, Optional, Set, Tuple

import streamlit as st
from docx import Document
from docx.enum.text import WD_UNDERLINE
from docx.oxml.ns import qn
from docx.shared import RGBColor

# ---------------------------------------------------------------------------
# 內建錯別字與用字對照（可擴充；鍵值長度將於套用時依長度優先處理，減少子字串誤替）
# ---------------------------------------------------------------------------
BUILTIN_REPLACEMENTS: Dict[str, str] = {
    "遺規": "違規",
    "冲動": "衝動",
    "散慢": "散漫",
    "違尿": "遺尿",
    "暴燥": "暴躁",
    "暴騷": "暴躁",
    "書藉": "書籍",
    "加許": "嘉許",
    "座標": "坐標",
    "裡": "裏",
    "溫度": "温度",
    "部份": "部分",
    "盡量": "儘量",
    "掃瞄二維碼": "掃描二維碼",
    "越來越多": "愈來愈多",
    "找贖": "找續",
    "計畫": "計劃",
}

# WORD JOINER（U+2060）：零寬不換行。實務上部分 Word／字型對中文仍可能在單一 WJ 處斷行，
# 故程式在「字與字之間」使用「雙重 WJ」加強黏合；必要時在詞首前再加一個 WJ，與前一字黏合，
# 讓整詞較容易一起移到下一行（仍無寬度、不影響版面對齊）。
WORD_JOINER = "\u2060"
# 字間黏合強度：雙重 WJ（仍為零寬字元，一般不會出現可見空隙）
WORD_GLUE_INNER = WORD_JOINER * 2

# 合併「人工斷詞」：兩字之間可能出現空白、半形/全形斜線（如排版誤加「 / 」）
_SPLIT_GAP_RE = r"(?:\s*[/／]?\s*)"

# 句尾已有標點則不再補句號
_SENTENCE_END_CHARS = set("。！？!?…」』\"'）)】］＞>、，,；;：:")

# 記錄／紀錄：後接此類字樣時，較可能為動詞「記錄」
_JILU_VERB_HINTS = (
    "了",
    "著",
    "過",
    "在",
    "中",
    "時",
    "一下",
    "起來",
    "成",
    "到",
    "給",
    "好",
    "完",
    "下",
    "出",
)

# 預設不可分割詞（可在 UI 修改）
DEFAULT_INDIVISIBLE = "廣場\n落淚\n靈魂"

# 內嵌圖／物件占位：私人使用區字元（U+FDD0），不應出於一般文稿；規則字典不會替換此字。
# 含圖段落先改為「文字 + 占位 + 文字…」參與修正，再依序插回複製的 w:r，以保留圖像。
IMAGE_PLACEHOLDER = "\ufdd0"


def _parse_custom_rules(text: str) -> Dict[str, str]:
    """解析「舊詞:新詞」每行一組；忽略空行與格式錯誤行。"""
    out: Dict[str, str] = {}
    for line in text.splitlines():
        line = line.strip()
        if not line or ":" not in line:
            continue
        old, new = line.split(":", 1)
        old, new = old.strip(), new.strip()
        if old:
            out[old] = new
    return out


def _parse_lines(text: str) -> List[str]:
    return [ln.strip() for ln in text.splitlines() if ln.strip()]


def _merge_intervals(intervals: List[Tuple[int, int]]) -> List[Tuple[int, int]]:
    if not intervals:
        return []
    intervals = sorted(intervals)
    merged = [intervals[0]]
    for a, b in intervals[1:]:
        la, lb = merged[-1]
        if a <= lb:
            merged[-1] = (la, max(lb, b))
        else:
            merged.append((a, b))
    return merged


def _intervals_to_set(intervals: List[Tuple[int, int]]) -> Set[int]:
    s: Set[int] = set()
    for a, b in intervals:
        s.update(range(a, b))
    return s


def _replace_non_overlapping(
    text: str,
    pattern_to_repl: List[Tuple[str, str]],
    on_each: Optional[Callable[[int, int, str], None]] = None,
) -> str:
    """
    由左而右：每次從當前位置找「最早出現」的 old；同位置取較長鍵。
    """
    if not text:
        return text
    result: List[str] = []
    i = 0
    n = len(text)
    while i < n:
        best_j: Optional[int] = None
        best_old = ""
        best_new = ""
        for old, new in pattern_to_repl:
            j = text.find(old, i)
            if j == -1:
                continue
            if best_j is None or j < best_j or (j == best_j and len(old) > len(best_old)):
                best_j = j
                best_old = old
                best_new = new
        if best_j is None:
            result.append(text[i:])
            break
        result.append(text[i:best_j])
        start = len("".join(result))
        result.append(best_new)
        end = start + len(best_new)
        if on_each:
            on_each(start, end, f"{best_old}→{best_new}")
        i = best_j + len(best_old)
    return "".join(result)


def _apply_sorted_replacements(
    text: str,
    replacements: Dict[str, str],
    stats: defaultdict,
    stat_prefix: str,
) -> Tuple[str, List[Tuple[int, int]]]:
    """依鍵長度由長到短非重疊取代；回傳新字串與紅字區間。"""
    if not replacements:
        return text, []
    pairs = sorted(replacements.items(), key=lambda x: len(x[0]), reverse=True)
    pattern_to_repl = [(a, b) for a, b in pairs if a]

    red_intervals: List[Tuple[int, int]] = []

    def on_rep(start: int, end: int, label: str) -> None:
        red_intervals.append((start, end))
        stats[f"{stat_prefix}:{label}"] += 1

    new_text = _replace_non_overlapping(text, pattern_to_repl, on_each=on_rep)
    return new_text, _merge_intervals(red_intervals)


def _apply_single_char_replace(
    text: str,
    old_ch: str,
    new_ch: str,
    stats: defaultdict,
    stat_key: str,
) -> Tuple[str, List[Tuple[int, int]]]:
    """單一字元全域替換並記錄紅字區間。"""
    if old_ch not in text:
        return text, []
    parts: List[str] = []
    red: List[Tuple[int, int]] = []
    pos = 0
    for i, ch in enumerate(text):
        if ch == old_ch:
            parts.append(text[pos:i])
            base = len("".join(parts))
            parts.append(new_ch)
            red.append((base, base + len(new_ch)))
            stats[stat_key] += 1
            pos = i + 1
    parts.append(text[pos:])
    return "".join(parts), _merge_intervals(red)


def _apply_zhuo_to_zhe(
    text: str,
    stats: defaultdict,
) -> Tuple[str, List[Tuple[int, int]]]:
    """僅「著名」「著作」保留「著」，其餘「著」→「着」。"""
    red: List[Tuple[int, int]] = []
    if "著" not in text:
        return text, red
    out: List[str] = []
    i = 0
    n = len(text)
    while i < n:
        if text[i] != "著":
            out.append(text[i])
            i += 1
            continue
        if text.startswith("著名", i):
            out.append("著名")
            i += 2
            continue
        if text.startswith("著作", i):
            out.append("著作")
            i += 2
            continue
        start = len("".join(out))
        out.append("着")
        red.append((start, start + 1))
        stats["著→着"] += 1
        i += 1
    return "".join(out), _merge_intervals(red)


def _should_be_jilu_verb(text: str, idx: int) -> Optional[bool]:
    """
    應為動詞「記錄」→ True；名詞「紀錄」→ False；無法判斷 → None。
    idx 指向「記」或「紀」。
    """
    if idx + 2 > len(text):
        return None
    c1 = text[idx + 2] if idx + 2 < len(text) else ""
    if not c1:
        return None
    if c1 in ("了", "著", "過", "在", "中", "時", "成", "到", "給", "好", "完", "下", "出"):
        return True
    if c1 in ("的", "和", "與", "及", "等", "、", "，"):
        return False
    rest = text[idx + 2 : idx + 6]
    for hint in _JILU_VERB_HINTS:
        if len(hint) >= 2 and rest.startswith(hint):
            return True
    return None


def _apply_jilu_jilu(
    text: str,
    stats: defaultdict,
) -> Tuple[str, List[Tuple[int, int]]]:
    """「記錄／紀錄」簡易區分；無法判斷則保留並計入待確認。"""
    red: List[Tuple[int, int]] = []
    if "記錄" not in text and "紀錄" not in text:
        return text, red

    out: List[str] = []
    i = 0
    n = len(text)
    while i < n:
        if i + 1 < n and text[i : i + 2] in ("記錄", "紀錄"):
            pair = text[i : i + 2]
            decision = _should_be_jilu_verb(text, i)
            if decision is None:
                out.append(pair)
                stats["記錄／紀錄待確認"] += 1
                i += 2
                continue
            want = "記錄" if decision else "紀錄"
            if pair == want:
                out.append(pair)
                i += 2
                continue
            start = len("".join(out))
            out.append(want)
            red.append((start, start + 2))
            stats[f"記錄／紀錄→{want}"] += 1
            i += 2
            continue
        out.append(text[i])
        i += 1
    return "".join(out), _merge_intervals(red)


def _apply_wei_to_wei_transition(text: str, stats: defaultdict) -> Tuple[str, List[Tuple[int, int]]]:
    """保守：「，唯」且下一字非「一」→「，惟」；紅字標在「惟」。"""
    red: List[Tuple[int, int]] = []
    pattern = re.compile(r"，唯(?![一])")
    out: List[str] = []
    last = 0
    for m in pattern.finditer(text):
        out.append(text[last : m.start()])
        start = len("".join(out))
        chunk = "，惟"
        out.append(chunk)
        red.append((start + 1, start + len(chunk)))
        stats["，唯→，惟"] += 1
        last = m.end()
    out.append(text[last:])
    return "".join(out), _merge_intervals(red)


def _should_add_period_at_end(text: str) -> bool:
    """段落結尾若像完整陳述但缺句號，則補上（略過極短行）。"""
    s = text.rstrip()
    if len(s) < 8:
        return False
    last = s[-1]
    if last in _SENTENCE_END_CHARS or last.isspace():
        return False
    if last.isalnum() or "\u4e00" <= last <= "\u9fff":
        return True
    return False


def _merge_red(*interval_lists: List[Tuple[int, int]]) -> List[Tuple[int, int]]:
    flat: List[Tuple[int, int]] = []
    for lst in interval_lists:
        flat.extend(lst)
    return _merge_intervals(flat)


def _flex_pattern_for_word(w: str) -> str:
    """
    組出可匹配「連續詞」或「中間被空白／斜線打斷」的規則運算式。
    例：廣場、廣/場、廣 / 場、廣　/　場 皆可匹配並還原為同一詞。
    """
    if len(w) < 2:
        return re.escape(w)
    parts: List[str] = [re.escape(w[0])]
    for ch in w[1:]:
        parts.append(_SPLIT_GAP_RE)
        parts.append(re.escape(ch))
    return "".join(parts)


def _normalize_indivisible_splits(text: str, words: List[str], stats: defaultdict) -> str:
    """
    將清單中詞彙的「人工斷行」還原為連續字（以便後續插入 WORD JOINER）。
    僅在長度>=2 且實際有斷開（含 / 或多餘空白）時計入「不可分割_合併人工斷字」。
    """
    words_sorted = sorted(set(w for w in words if len(w) >= 2), key=len, reverse=True)
    if not words_sorted:
        return text
    out: List[str] = []
    i = 0
    n = len(text)
    while i < n:
        matched = False
        for w in words_sorted:
            pat = _flex_pattern_for_word(w)
            m = re.match(pat, text[i:])
            if not m:
                continue
            raw = m.group(0)
            out.append(w)
            if raw != w:
                stats["不可分割_合併人工斷字"] += 1
            i += len(raw)
            matched = True
            break
        if not matched:
            out.append(text[i])
            i += 1
    return "".join(out)


def _apply_proper_nouns(
    text: str,
    names: List[str],
    stats: defaultdict,
) -> Tuple[str, List[Tuple[int, int]]]:
    """專有名詞底線區間（長詞優先、不重疊）。"""
    names_sorted = sorted(set(names), key=len, reverse=True)
    underline: List[Tuple[int, int]] = []
    occupied: List[Tuple[int, int]] = []

    def overlaps(a: int, b: int) -> bool:
        for x, y in occupied:
            if not (b <= x or a >= y):
                return True
        return False

    for name in names_sorted:
        if not name:
            continue
        start = 0
        while True:
            j = text.find(name, start)
            if j == -1:
                break
            end = j + len(name)
            if not overlaps(j, end):
                underline.append((j, end))
                occupied.append((j, end))
                stats["專名號底線"] += 1
            start = j + 1

    return text, _merge_intervals(underline)


def _apply_word_joiners(
    text: str,
    words: List[str],
    stats: defaultdict,
) -> Tuple[str, List[Tuple[int, int]]]:
    """
    不可分割詞：字間插入「雙重 WORD JOINER」；若詞前還有字元，詞首前再插一個 WJ。
    整段替換結果（含 joiner）標紅。
    掃描：由左而右，每步嘗試最長詞；已匹配區間不重疊。
    """
    words_sorted = sorted(set(w for w in words if len(w) >= 2), key=len, reverse=True)
    if not words_sorted:
        return text, []

    occupied: List[Tuple[int, int]] = []

    def overlaps(a: int, b: int) -> bool:
        for x, y in occupied:
            if not (b <= x or a >= y):
                return True
        return False

    matches: List[Tuple[int, int, str]] = []
    i = 0
    n = len(text)
    while i < n:
        found: Optional[str] = None
        for w in words_sorted:
            if i + len(w) <= n and text[i : i + len(w)] == w and not overlaps(i, i + len(w)):
                found = w
                break
        if found:
            matches.append((i, i + len(found), found))
            occupied.append((i, i + len(found)))
            i += len(found)
        else:
            i += 1

    if not matches:
        return text, []

    red: List[Tuple[int, int]] = []
    out: List[str] = []
    last = 0
    for start, end, w in matches:
        out.append(text[last:start])
        pos = len("".join(out))
        # 字間：雙重 WJ，加強 Word 對中文「不斷詞」的黏著（單一 WJ 有時仍會斷行）
        core = WORD_GLUE_INNER.join(list(w))
        # 詞首前若還有前綴字元，多插一個 WJ，與前一字黏合，利於整詞一起換行
        if start > 0:
            piece = WORD_JOINER + core
        else:
            piece = core
        out.append(piece)
        red.append((pos, pos + len(piece)))
        stats["排位調整_WORD_JOINER"] += 1
        last = end
    out.append(text[last:])
    return "".join(out), _merge_intervals(red)


def _map_span_before_to_after(
    before: str,
    after: str,
    bs: int,
    be: int,
) -> Optional[Tuple[int, int]]:
    """
    將 before 中半開區間 [bs, be) 對應到 after 中的半開區間 [a0, a1)。
    假設 after 僅比 before「在詞內多插入了 WORD_JOINER」，其餘字元順序一致。
    """
    if bs >= be:
        return (0, 0)
    ib, ia = 0, 0
    nb, na = len(before), len(after)
    a0: Optional[int] = None
    while ib < nb and ia < na:
        # 連續的 WJ（含雙重黏合、詞首黏合）一律略過，只對齊「真實文字」
        while ia < na and after[ia] == WORD_JOINER:
            ia += 1
        if ia >= na:
            break
        if ib == bs:
            a0 = ia
        if ib == be - 1:
            return (a0 if a0 is not None else ia, ia + 1)
        ib += 1
        ia += 1
    return None


def _map_underline_after_joiners(
    text_before: str,
    text_after: str,
    underline_before: List[Tuple[int, int]],
) -> List[Tuple[int, int]]:
    """專名底線區間在插入 WORD JOINER 後，映射到最終字串索引。"""
    if not underline_before:
        return []
    spans: List[Tuple[int, int]] = []
    for bs, be in _merge_intervals(underline_before):
        m = _map_span_before_to_after(text_before, text_after, bs, be)
        if m:
            spans.append(m)
    return _merge_intervals(spans)


def process_paragraph_plain_text(
    raw: str,
    custom_rules: Dict[str, str],
    proper_names: List[str],
    indivisible: List[str],
    stats: defaultdict,
    is_paragraph_for_period: bool,
) -> Tuple[str, Set[int], Set[int]]:
    """
    對單一段落純文字套用全部規則。
    回傳：(最終文字, 紅字 index 集合, 底線 index 集合)
    """
    t = raw

    # 1) 內建字典
    t, r1 = _apply_sorted_replacements(t, BUILTIN_REPLACEMENTS, stats, "內建")

    # 2) 自訂規則（覆寫／額外；同樣長度優先）
    t, r2 = _apply_sorted_replacements(t, custom_rules, stats, "自訂")

    # 3) 「佈」→「布」
    t, r3 = _apply_single_char_replace(t, "佈", "布", stats, "佈→布")

    # 4) 「著」/「着」
    t, r4 = _apply_zhuo_to_zhe(t, stats)

    # 5) 「記錄」/「紀錄」
    t, r5 = _apply_jilu_jilu(t, stats)

    # 6) 「唯」→「惟」（轉折）
    t, r6 = _apply_wei_to_wei_transition(t, stats)

    # 7) 「情景」→「情境」
    qingjing: Dict[str, str] = {"情景": "情境"}
    t, r7 = _apply_sorted_replacements(t, qingjing, stats, "情景")

    red = _merge_red(r1, r2, r3, r4, r5, r6, r7)

    # 8) 句尾標點（僅針對「段落」層級呼叫時 is_paragraph_for_period=True）
    if is_paragraph_for_period and _should_add_period_at_end(t):
        t = t.rstrip()
        t = t + "。"
        pos = len(t) - 1
        red.append((pos, pos + 1))
        stats["句尾標點補上"] += 1

    red = _merge_intervals(red)

    # 8.5) 不可分割詞：先合併「廣 / 場」類人工斷詞（與測試圖一致），再插入 WORD JOINER
    t = _normalize_indivisible_splits(t, indivisible, stats)

    # 9) 專有名詞底線（在 joiner 前先於目前字串上找詞）
    t, underline_before = _apply_proper_nouns(t, proper_names, stats)

    # 10) WORD JOINER（可能拉長字串；專名底線需映射到新索引）
    t_before_join = t
    t, r_join = _apply_word_joiners(t, indivisible, stats)
    red = _merge_red(red, r_join)

    underline_after = _map_underline_after_joiners(t_before_join, t, underline_before)

    red_set = _intervals_to_set(red)
    under_set = _intervals_to_set(underline_after)
    return t, red_set, under_set


def _paragraph_plain(paragraph) -> str:
    return "".join(run.text for run in paragraph.runs)


def _run_has_drawing(run) -> bool:
    """偵測 run 是否含內嵌圖／繪圖（w:drawing、舊版 pict 等）。"""
    xml = run._element  # lxml
    for el in xml.iter():
        tag = el.tag
        if tag.endswith("drawing") or tag.endswith("pict") or tag.endswith("binaryData"):
            return True
    return False


def _paragraph_text_and_drawings(paragraph) -> Tuple[str, List[Any]]:
    """
    串接段落可見文字；遇含圖 run 則插入 IMAGE_PLACEHOLDER 並保存該 w:r 的深拷貝（供還原）。
    順序與 paragraph.runs 一致。
    """
    drawings: List[Any] = []
    parts: List[str] = []
    for run in paragraph.runs:
        if _run_has_drawing(run):
            parts.append(IMAGE_PLACEHOLDER)
            drawings.append(deepcopy(run._element))
        else:
            parts.append(run.text or "")
    return "".join(parts), drawings


def _clear_paragraph_content_keep_ppr(paragraph) -> None:
    """刪除段落內容但保留 w:pPr（對齊、間距、大綱層級等），避免表格／樣式走位。"""
    p = paragraph._p
    for child in list(p):
        if child.tag == qn("w:pPr"):
            continue
        p.remove(child)


def _append_formatted_runs(
    paragraph,
    text: str,
    red_indices: Set[int],
    underline_indices: Set[int],
) -> None:
    """在段落末尾追加具紅字／底線的 runs（不清空段落）。"""
    if not text:
        return
    n = len(text)
    i = 0
    while i < n:
        is_red = i in red_indices
        is_ul = i in underline_indices
        j = i + 1
        while j < n:
            jr = j in red_indices
            ju = j in underline_indices
            if jr != is_red or ju != is_ul:
                break
            j += 1
        run = paragraph.add_run(text[i:j])
        if is_red:
            run.font.color.rgb = RGBColor(255, 0, 0)
        if is_ul:
            run.font.underline = WD_UNDERLINE.SINGLE
        i = j


def _write_formatted_runs_full(
    paragraph,
    text: str,
    red_indices: Set[int],
    underline_indices: Set[int],
) -> None:
    """清空段落內容（保留 pPr）後寫入紅字／底線 runs。"""
    _clear_paragraph_content_keep_ppr(paragraph)
    _append_formatted_runs(paragraph, text, red_indices, underline_indices)


def _rebuild_paragraph_with_image_placeholders(
    paragraph,
    new_full: str,
    drawing_elements: List[Any],
    red_indices: Set[int],
    underline_indices: Set[int],
) -> bool:
    """
    將含 IMAGE_PLACEHOLDER 的字串還原為：文字 runs + 依序插入的圖像 w:r。
    各文字片段的紅／底線索引為 new_full 中的絕對位置換算為區段相對位置。
    """
    ph = IMAGE_PLACEHOLDER
    n_ph = new_full.count(ph)
    if n_ph != len(drawing_elements):
        # 規則誤動到占位符：不覆寫本段，避免圖片遺失（統計於外層）
        return False
    parts = new_full.split(ph)
    _clear_paragraph_content_keep_ppr(paragraph)
    pos = 0
    ph_len = len(ph)
    for i, part in enumerate(parts):
        seg_start = pos
        seg_end = pos + len(part)
        rs = {x - seg_start for x in red_indices if seg_start <= x < seg_end}
        us = {x - seg_start for x in underline_indices if seg_start <= x < seg_end}
        _append_formatted_runs(paragraph, part, rs, us)
        pos = seg_end
        if i < len(drawing_elements):
            paragraph._p.append(drawing_elements[i])
            pos += ph_len
    return True


def _apply_paragraph_format(
    paragraph,
    text: str,
    red_indices: Set[int],
    underline_indices: Set[int],
) -> None:
    """依字元索引重建 run（純文字段落，無內嵌圖）。"""
    _write_formatted_runs_full(paragraph, text, red_indices, underline_indices)


def iter_all_paragraphs(doc: Document):
    """本文、表格（含巢狀）、頁首、頁尾。"""
    for p in doc.paragraphs:
        yield p

    def walk_table(table):
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    yield p
                for nt in cell.tables:
                    yield from walk_table(nt)

    for table in doc.tables:
        yield from walk_table(table)

    for section in doc.sections:
        for part in (section.header, section.footer):
            try:
                for p in part.paragraphs:
                    yield p
                for table in part.tables:
                    yield from walk_table(table)
            except Exception:
                pass


def process_document(
    doc: Document,
    custom_rules: Dict[str, str],
    proper_names: List[str],
    indivisible: List[str],
) -> defaultdict:
    """就地修改 doc，並回傳統計。含內嵌圖之段落以占位符保留圖片 run，其餘保留 w:pPr。"""
    stats: defaultdict = defaultdict(int)
    for p in iter_all_paragraphs(doc):
        text_ph, drs = _paragraph_text_and_drawings(p)
        new_text, red, under = process_paragraph_plain_text(
            text_ph,
            custom_rules,
            proper_names,
            indivisible,
            stats,
            is_paragraph_for_period=True,
        )
        if not drs:
            _apply_paragraph_format(p, new_text, red, under)
        else:
            ok = _rebuild_paragraph_with_image_placeholders(
                p, new_text, drs, red, under
            )
            if not ok:
                stats["含圖段落占位異常_未改寫"] += 1
    return stats


def _total_changes(stats: defaultdict) -> int:
    """估算「修正處數」：各統計鍵加總（待確認亦計入提醒）。"""
    return int(sum(stats.values()))


def main() -> None:
    st.set_page_config(page_title="教材/單元報告審核工具", layout="wide")
    st.title("教材/單元報告審核工具")

    if "processed_bytes" not in st.session_state:
        st.session_state.processed_bytes = None
    if "last_stats" not in st.session_state:
        st.session_state.last_stats = None
    if "last_file_id" not in st.session_state:
        st.session_state.last_file_id = None
    if "download_filename" not in st.session_state:
        st.session_state.download_filename = "document_fixed.docx"

    col_main, col_side = st.columns([2, 1])

    with col_side:
        st.subheader("規則與清單")
        custom_rules_text = st.text_area(
            "自訂額外規則（每行：舊詞:新詞）",
            height=120,
            placeholder="遺規:違規",
        )
        proper_text = st.text_area(
            "自訂專有名詞清單（每行一個）",
            height=120,
        )
        indiv_text = st.text_area(
            "不可分割詞彙清單（每行一個）",
            value=DEFAULT_INDIVISIBLE,
            height=120,
        )

    with col_main:
        up = st.file_uploader("上傳 Word 檔（僅 .docx）", type=["docx"])
        run_btn = st.button("開始審核並套用修正", type="primary")

    if up is not None:
        file_id = (up.name, up.size)
        new_upload = st.session_state.last_file_id != file_id
        if run_btn or new_upload:
            st.session_state.last_file_id = file_id
            data = up.getvalue()
            custom = _parse_custom_rules(custom_rules_text)
            names = _parse_lines(proper_text)
            indiv = _parse_lines(indiv_text)
            doc = Document(io.BytesIO(data))
            stats = process_document(doc, custom, names, indiv)
            buf = io.BytesIO()
            doc.save(buf)
            buf.seek(0)
            st.session_state.processed_bytes = buf.getvalue()
            st.session_state.last_stats = dict(stats)
            base, _ = os.path.splitext(up.name)
            st.session_state.download_filename = f"{base}_fixed.docx"

    if st.session_state.processed_bytes is not None:
        stats = st.session_state.last_stats or {}
        total = _total_changes(defaultdict(int, stats))
        st.success(f"本次自動修正（統計項目加總）共 **{total}** 項；請見下方分類。")
        with st.expander("詳細摘要", expanded=True):
            # 分類顯示
            builtin_lines = [f"- {k}: {v}" for k, v in sorted(stats.items()) if k.startswith("內建:")]
            custom_lines = [f"- {k}: {v}" for k, v in sorted(stats.items()) if k.startswith("自訂:")]
            other_keys = [
                k
                for k in sorted(stats.keys())
                if not k.startswith("內建:") and not k.startswith("自訂:")
            ]
            if builtin_lines:
                st.markdown("**內建規則**")
                st.markdown("\n".join(builtin_lines))
            if custom_lines:
                st.markdown("**自訂規則**")
                st.markdown("\n".join(custom_lines))
            st.markdown("**其他**")
            for k in other_keys:
                st.markdown(f"- {k}: {stats[k]}")

        st.download_button(
            label="下載已修改並紅色標記＋底線＋排位優化的 DOCX",
            data=st.session_state.processed_bytes,
            file_name=st.session_state.download_filename,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )


if __name__ == "__main__":
    main()
