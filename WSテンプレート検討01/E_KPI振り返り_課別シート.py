# -*- coding: utf-8 -*-
from __future__ import annotations

import os
import re
import shutil
import time
from typing import List, Dict

import win32com.client as win32


# =========================
# 入出力
# =========================
SRC_XLSX = r"C:\Users\SujiT\OneDrive\ドキュメント\HiproBiz わかもと製薬\KPI振り返りシート\20研修_KPI振り返りシート_最終案\WSテンプレート\KPI振り返り_WSシート01.xlsx"

OUT_DIR = r"C:\Users\SujiT\OneDrive\ドキュメント\HiproBiz わかもと製薬\KPI振り返りシート\20研修_KPI振り返りシート_最終案\WSテンプレート\課別配布"

# ベースシート候補（テンプレは通常こちら）
BASE_SHEET_CANDIDATES = [
    "00_全課_①②③検討",
    "00_全課_①②③検討_サンプル",
    "00_全課_①②③検討_サンプルシート",
]

# 北→南（No順）
ORDER_KA: List[str] = [
    "札幌",
    "仙台",
    "東京1課",
    "東京2課",
    "東京3課",
    "東京4課",
    "東京5課",
    "名古屋1課",
    "名古屋2課",
    "大阪1課",
    "大阪2課",
    "広島課",
    "福岡1課",
    "福岡2課",
]

# 追加する「次フェーズKPI」行数
ADD_NEXT_KPI_ROWS = 3

# 追加行に入れる値
NEXT_PHASE_KPI_KUBUN = "発売後"
NEXT_PHASE_PERIOD = "～6月末"

# 入力列（指定列のみ折り返しON）
WRAP_COLS = [5, 13, 15, 16]  # E, M, O, P

# N列の幅（狭くて表示されない対策）
N_COL = 14
N_COL_WIDTH = 30  # 初期幅
N_COL_MIN_WIDTH = 30  # 最低幅保証


# =========================
# Excel定数（win32com）
# =========================
xlPasteFormats = -4122
xlPasteValidation = 6
xlCenter = -4108


# =========================
# util
# =========================
def norm(v) -> str:
    if v is None:
        return ""
    return str(v).replace(" ", "").replace("　", "").replace("\n", "").strip()

def safe_name(s: str) -> str:
    return re.sub(r'[\\/:*?"<>|]', "_", s).strip()

def safe_sheet_title(title: str, max_len: int = 31) -> str:
    t = re.sub(r'[:\\/?*\[\]]', "_", title).strip()
    return t[:max_len] if len(t) > max_len else t

def ka_key(s: str) -> str:
    t = norm(s)
    if t.endswith("課"):
        t = t[:-1]
    return t

def sheet_names(wb) -> List[str]:
    return [wb.Worksheets(i).Name for i in range(1, wb.Worksheets.Count + 1)]

def pick_base_sheet(wb) -> str:
    names = sheet_names(wb)
    for n in BASE_SHEET_CANDIDATES:
        if n in names:
            return n
    for n in names:
        if ("全課" in n) and ("①②③" in n):
            return n
    raise RuntimeError(f"ベースシートが見つかりません。存在するシート: {names}")

def find_header_row(ws, max_scan: int = 120) -> int:
    """
    見出し行を探す（「課」「KPI区分」「評価対象期間」）
    ※セル内改行を含むので部分一致で判定
    """
    for r in range(1, max_scan + 1):
        row_texts = []
        for c in range(1, 61):
            s = norm(ws.Cells(r, c).Value)
            if s:
                row_texts.append(s.replace("\n", ""))
        joined = " ".join(row_texts)
        if ("課" in joined) and ("KPI" in joined and "区分" in joined) and ("評価" in joined and "対象" in joined and "期間" in joined):
            return r
    return 4

def detect_columns(ws, header_row: int) -> Dict[str, int]:
    """
    ヘッダー行から列位置を自動検出
    keys: no, ka, kubun, period, kpi
    """
    found: Dict[str, int] = {}
    for c in range(1, 61):
        s = norm(ws.Cells(header_row, c).Value).replace("\n", "")
        if not s:
            continue

        if s in ("No", "NO", "番号"):
            found["no"] = c
        if s == "課":
            found["ka"] = c
        if ("KPI" in s) and ("区分" in s):
            found["kubun"] = c
        if ("評価" in s) and ("対象" in s) and ("期間" in s):
            found["period"] = c
        if s == "KPI":
            found["kpi"] = c

    found.setdefault("no", 1)
    found.setdefault("ka", 2)
    found.setdefault("kubun", 3)
    found.setdefault("period", 4)
    found.setdefault("kpi", 5)
    return found

def last_used_row(ws) -> int:
    ur = ws.UsedRange
    return ur.Row + ur.Rows.Count - 1

def row_has_any_value(ws, r: int, c_from: int, c_to: int) -> bool:
    for c in range(c_from, c_to + 1):
        if norm(ws.Cells(r, c).Value) != "":
            return True
    return False

def is_true_template_row(ws, r: int, col_ka: int, col_kpi: int, header_row: int) -> bool:
    """
    “全課に残す” 真テンプレ行：
    - 課が空
    - KPIが空
    - かつ B〜S（2〜19）に一切値が無い（＝枠/空行のみ）
    """
    if r <= header_row:
        return False

    if norm(ws.Cells(r, col_ka).Value) != "":
        return False
    if norm(ws.Cells(r, col_kpi).Value) != "":
        return False

    if row_has_any_value(ws, r, 2, 19):
        return False

    return True

def delete_other_sheets(wb, keep_names: List[str]) -> None:
    keep = set(keep_names)
    for s in list(sheet_names(wb)):
        if s not in keep:
            wb.Worksheets(s).Delete()

def clear_row_values(ws, r: int, col_from: int, col_to: int) -> None:
    ws.Range(ws.Cells(r, col_from), ws.Cells(r, col_to)).ClearContents()

def copy_formats_and_validation_from_row(ws, src_row: int, dst_row: int) -> None:
    """
    罫線含む書式＋入力規則（DV）を、src_row → dst_row へコピー
    """
    ws.Rows(src_row).Copy()
    ws.Rows(dst_row).PasteSpecial(Paste=xlPasteFormats)
    ws.Rows(dst_row).PasteSpecial(Paste=xlPasteValidation)

def safe_save_close(wb) -> None:
    last_err = None
    for _ in range(6):
        try:
            wb.Save()
            wb.Close(SaveChanges=False)
            return
        except Exception as e:
            last_err = e
            time.sleep(0.8)
    try:
        wb.Close(SaveChanges=False)
    except Exception:
        pass
    if last_err:
        raise last_err

def set_wrap_for_specific_columns(ws, header_row: int, cols: List[int]) -> None:
    """
    指定列のみ（入力列）折り返しON
    """
    last = last_used_row(ws)
    if last <= header_row:
        return
    for c in cols:
        rng = ws.Range(ws.Cells(header_row + 1, c), ws.Cells(last, c))
        rng.WrapText = True

def set_column_width(ws, col: int, width: float) -> None:
    ws.Columns(col).ColumnWidth = width

def enforce_min_col_width(ws, col: int, min_width: float) -> None:
    if ws.Columns(col).ColumnWidth < min_width:
        ws.Columns(col).ColumnWidth = min_width

def align_no_column_center(ws, header_row: int, col_no: int = 1) -> None:
    """
    A列（No）を縦横中央に
    """
    last = last_used_row(ws)
    if last <= header_row:
        return
    rng = ws.Range(ws.Cells(header_row + 1, col_no), ws.Cells(last, col_no))
    rng.HorizontalAlignment = xlCenter
    rng.VerticalAlignment = xlCenter

def autofit_rows(ws) -> None:
    ws.Rows.AutoFit()

def fix_merged_title_rows(ws, title_rows=(4, 5), col_from=1, col_to=19, row_height=36) -> None:
    """
    結合セルのタイトル行が AutoFit で潰れるのを防止：
    - 折り返しON
    - 行高を固定で復元
    """
    r1, r2 = title_rows
    rng = ws.Range(ws.Cells(r1, col_from), ws.Cells(r2, col_to))
    rng.WrapText = True
    ws.Rows(f"{r1}:{r2}").RowHeight = row_height


# =========================
# main
# =========================
def main():
    os.makedirs(OUT_DIR, exist_ok=True)

    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    try:
        for i, ka in enumerate(ORDER_KA, start=1):
            out_path = os.path.join(
                OUT_DIR,
                f"{i:02d}_{safe_name(ka)}_①②③振り返り（研修用）.xlsx"
            )

            # SaveAsせず「物理コピー→直接編集→Save」
            if os.path.exists(out_path):
                os.remove(out_path)
            shutil.copyfile(SRC_XLSX, out_path)

            wb = excel.Workbooks.Open(out_path, ReadOnly=False)
            base_sheet_name = pick_base_sheet(wb)

            # DVリスト保持のため「_」から始まるシートは残す
            keeps = [base_sheet_name]
            for s in sheet_names(wb):
                if s.startswith("_"):
                    keeps.append(s)
            delete_other_sheets(wb, keeps)

            ws = wb.Worksheets(base_sheet_name)

            # 見出し行・列検出
            header_row = find_header_row(ws)
            cols = detect_columns(ws, header_row)

            col_no = cols["no"]
            col_ka = cols["ka"]
            col_kubun = cols["kubun"]
            col_period = cols["period"]
            col_kpi = cols["kpi"]

            target_key = ka_key(ka)

            # ---------- 課別に切り分け（下から削除） ----------
            last_row = last_used_row(ws)
            for r in range(last_row, header_row, -1):
                if is_true_template_row(ws, r, col_ka, col_kpi, header_row):
                    continue
                ka_in_sheet = ka_key(ws.Cells(r, col_ka).Value)
                if ka_in_sheet != target_key:
                    ws.Rows(r).Delete()

            # ---------- No/課名の再付与（テンプレ行は触らない） ----------
            last_row2 = last_used_row(ws)
            for r in range(header_row + 1, last_row2 + 1):
                if is_true_template_row(ws, r, col_ka, col_kpi, header_row):
                    continue
                if row_has_any_value(ws, r, 2, 19):
                    ws.Cells(r, col_no).Value = i
                    ws.Cells(r, col_ka).Value = ka

            # ---------- 次フェーズKPI入力行（3行）追加 ----------
            last_row3 = last_used_row(ws)
            insert_after = header_row
            for r in range(header_row + 1, last_row3 + 1):
                if is_true_template_row(ws, r, col_ka, col_kpi, header_row):
                    continue
                if ka_key(ws.Cells(r, col_ka).Value) == target_key and row_has_any_value(ws, r, 2, 19):
                    insert_after = r

            insert_at = insert_after + 1
            ws.Rows(insert_at).Resize(ADD_NEXT_KPI_ROWS).Insert()

            src_format_row = insert_after if insert_after > header_row else (header_row + 1)
            for rr in range(insert_at, insert_at + ADD_NEXT_KPI_ROWS):
                copy_formats_and_validation_from_row(ws, src_format_row, rr)
                clear_row_values(ws, rr, 1, 19)

                ws.Cells(rr, col_no).Value = i
                ws.Cells(rr, col_ka).Value = ka
                ws.Cells(rr, col_kubun).Value = NEXT_PHASE_KPI_KUBUN
                ws.Cells(rr, col_period).Value = NEXT_PHASE_PERIOD
                ws.Cells(rr, col_kpi).Value = ""

            # ---------- 表示調整（ご要望分） ----------
            # 1) 入力列のみ折り返しON（E/M/O/P）
            set_wrap_for_specific_columns(ws, header_row, WRAP_COLS)

            # 2) N列の列幅を広げる（先に広げる）
            set_column_width(ws, N_COL, N_COL_WIDTH)

            # 3) A列（No）を中央寄せ（縦横）
            align_no_column_center(ws, header_row, col_no=1)

            # 4) 折り返し後に行高を自動調整
            excel.CalculateFull()
            time.sleep(0.2)
            autofit_rows(ws)

            # 5) ★結合タイトル（4-5行目）を復元（AutoFitで潰れるため）
            fix_merged_title_rows(ws, title_rows=(4, 5), col_from=1, col_to=19, row_height=36)

            # 6) ★N列は AutoFit で縮むことがあるので “最低幅” を強制
            enforce_min_col_width(ws, N_COL, N_COL_MIN_WIDTH)

            # ---------- シート名をファイル名に一致 ----------
            new_sheet_name = safe_sheet_title(f"{ka}_①②③振り返り（研修用）")
            ws.Name = new_sheet_name

            safe_save_close(wb)
            print("作成完了:", out_path, " / base_sheet:", base_sheet_name, " / sheet:", new_sheet_name)

    finally:
        excel.Quit()


if __name__ == "__main__":
    main()
