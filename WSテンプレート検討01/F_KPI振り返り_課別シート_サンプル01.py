# -*- coding: utf-8 -*-
from __future__ import annotations

import os
import re
import shutil
import time
from typing import List, Optional, Dict, Tuple

import win32com.client as win32


# =========================
# 入出力
# =========================
SRC_XLSX = r"C:\Users\SujiT\OneDrive\ドキュメント\HiproBiz わかもと製薬\KPI振り返りシート\20研修_KPI振り返りシート_最終案\WSテンプレート\KPI振り返り_WSシート01.xlsx"

OUT_DIR = r"C:\Users\SujiT\OneDrive\ドキュメント\HiproBiz わかもと製薬\KPI振り返りシート\20研修_KPI振り返りシート_最終案\WSテンプレート\課別配布サンプル"

BASE_SHEET_CANDIDATES = [
    "00_全課_①②③検討_サンプル",
    "00_全課_①②③検討_サンプルシート",
    "00_全課_①②③検討",
]

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


# =========================
# Excel定数（win32com）
# =========================
xlPasteFormats = -4122
xlPasteValidation = 6


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
    """
    課名照合キー：末尾「課」除去 + 空白除去
    """
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
    # 最後の保険：部分一致
    for n in names:
        if ("全課" in n) and ("①②③" in n):
            return n
    raise RuntimeError(f"ベースシートが見つかりません。存在するシート: {names}")

def find_header_row(ws, max_scan: int = 80) -> int:
    """
    見出し行を探す（「課」「KPI区分」「評価対象期間」）
    ※セル内改行を含むので部分一致で判定
    """
    for r in range(1, max_scan + 1):
        row_texts = []
        for c in range(1, 61):
            v = ws.Cells(r, c).Value
            s = norm(v)
            if s:
                row_texts.append(s)
        joined = " ".join(row_texts)
        if ("課" in joined) and ("KPI" in joined and "区分" in joined) and ("評価" in joined and "対象" in joined):
            return r
    # ダメなら4（添付は4）
    return 4

def detect_columns(ws, header_row: int) -> Dict[str, int]:
    """
    ヘッダー行から列位置を自動検出
    keys: ka, kubun, period, kpi
    """
    def cell_s(c: int) -> str:
        return norm(ws.Cells(header_row, c).Value)

    found = {}
    for c in range(1, 61):
        s = cell_s(c)
        if not s:
            continue

        s2 = s.replace("\n", "")
        # 課
        if (s2 == "課") or (s2.endswith("課") and "課" == s2):
            found["ka"] = c

        # KPI区分（"KPI\n区分"など）
        if ("KPI" in s2) and ("区分" in s2):
            found["kubun"] = c

        # 評価対象期間（"評価\n対象期間"など）
        if ("評価" in s2) and ("対象" in s2) and ("期間" in s2):
            found["period"] = c

        # KPI（単独KPIの列）
        # "KPI活動結果" などは除外したいので、完全一致寄り
        if s2 == "KPI":
            found["kpi"] = c

    # 添付ファイルの既定（保険）
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
    - 見出し行より下
    - 課(B相当)が空
    - KPI列が空
    - かつ 入力領域（B〜S相当：2〜19）のどこにも値がない
      → 枠だけ/空行だけ残す

    ※ここが「名古屋行が混ざる」を防ぐ肝
    """
    if r <= header_row:
        return False

    if norm(ws.Cells(r, col_ka).Value) != "":
        return False
    if norm(ws.Cells(r, col_kpi).Value) != "":
        return False

    # 2〜19（B〜S）をチェック（列ズレがあっても入力領域として大抵この範囲）
    # もし将来ズレるならここをヘッダー検出ベースで拡張します
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
    # OneDrive配下で Save が一瞬失敗することがあるため軽くリトライ
    last_err = None
    for _ in range(5):
        try:
            wb.Save()
            wb.Close(SaveChanges=False)
            return
        except Exception as e:
            last_err = e
            time.sleep(0.8)
    # 最後にCloseだけは試す
    try:
        wb.Close(SaveChanges=False)
    except Exception:
        pass
    if last_err:
        raise last_err


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

            # 物理コピーで必ず1ファイル1回（SaveAsしない）
            if os.path.exists(out_path):
                os.remove(out_path)
            shutil.copyfile(SRC_XLSX, out_path)

            wb = excel.Workbooks.Open(out_path, ReadOnly=False)
            base_sheet_name = pick_base_sheet(wb)

            # DVリスト保持のため _ から始まるシートは残す
            keeps = [base_sheet_name]
            for s in sheet_names(wb):
                if s.startswith("_"):
                    keeps.append(s)
            delete_other_sheets(wb, keeps)

            ws = wb.Worksheets(base_sheet_name)

            # ヘッダと列検出
            header_row = find_header_row(ws)
            cols = detect_columns(ws, header_row)
            col_ka = cols["ka"]
            col_kubun = cols["kubun"]
            col_period = cols["period"]
            col_kpi = cols["kpi"]

            target_key = ka_key(ka)

            # --- 下から切り分け（課一致の行だけ残す／真テンプレ行は残す） ---
            last_row = last_used_row(ws)
            for r in range(last_row, header_row, -1):
                if is_true_template_row(ws, r, col_ka, col_kpi, header_row):
                    continue

                # 入力領域に何も無い（空行）の場合は残しても害は少ないが、ここではテンプレ以外は削除寄りにする
                # ※混入事故を防ぐため
                ka_in_sheet = ka_key(ws.Cells(r, col_ka).Value)
                if ka_in_sheet != target_key:
                    ws.Rows(r).Delete()

            # --- 切り分け後：No/課名の再付与（真テンプレ行は触らない） ---
            last_row2 = last_used_row(ws)
            for r in range(header_row + 1, last_row2 + 1):
                if is_true_template_row(ws, r, col_ka, col_kpi, header_row):
                    continue
                # B〜Sに何かある行のみ対象
                if row_has_any_value(ws, r, 2, 19):
                    ws.Cells(r, 1).Value = i
                    ws.Cells(r, col_ka).Value = ka

            # --- 次フェーズKPI入力行（3行）を追加 ---
            # “この課の最終入力行”を探す：課一致で B〜S に何か入っている最後の行
            last_row3 = last_used_row(ws)
            insert_after = header_row
            for r in range(header_row + 1, last_row3 + 1):
                if is_true_template_row(ws, r, col_ka, col_kpi, header_row):
                    continue
                if ka_key(ws.Cells(r, col_ka).Value) == target_key and row_has_any_value(ws, r, 2, 19):
                    insert_after = r

            insert_at = insert_after + 1
            ws.Rows(insert_at).Resize(ADD_NEXT_KPI_ROWS).Insert()

            # 罫線・書式・入力規則を “挿入直前の行” からコピーして揃える
            # （insert_after が header の場合もあるので、その場合は header+1 を参照）
            src_format_row = insert_after if insert_after > header_row else (header_row + 1)

            for rr in range(insert_at, insert_at + ADD_NEXT_KPI_ROWS):
                copy_formats_and_validation_from_row(ws, src_format_row, rr)

                # 値は一旦クリア（コピーで元行の値が入る可能性があるため）
                clear_row_values(ws, rr, 1, 19)

                # 必須項目をセット
                ws.Cells(rr, 1).Value = i
                ws.Cells(rr, col_ka).Value = ka
                ws.Cells(rr, col_kubun).Value = NEXT_PHASE_KPI_KUBUN
                ws.Cells(rr, col_period).Value = NEXT_PHASE_PERIOD

                # KPI列は空欄（ここに次フェーズKPIを記入）
                ws.Cells(rr, col_kpi).Value = ""

            # --- シート名をファイル名と一致させる ---
            new_sheet_name = safe_sheet_title(f"{ka}_①②③振り返り（研修用）")
            ws.Name = new_sheet_name

            safe_save_close(wb)
            print("作成完了:", out_path, " / sheet:", new_sheet_name)

    finally:
        excel.Quit()


if __name__ == "__main__":
    main()
