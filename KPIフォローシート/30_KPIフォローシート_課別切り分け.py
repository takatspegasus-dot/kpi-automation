# -*- coding: utf-8 -*-
from __future__ import annotations

import os
import re
import shutil

import pythoncom
import win32com.client as win32


# ===== 入力（元ブック）=====
INPUT_XLSX = r"C:\Users\SujiT\OneDrive\ドキュメント\HiproBiz わかもと製薬\わかもと製薬_KPIフォローシート\KPIフォローシート\KPIフォローシート04.xlsx"

# ===== 出力（課別配布：テンプレート）=====
OUT_DIR = r"C:\Users\SujiT\OneDrive\ドキュメント\HiproBiz わかもと製薬\わかもと製薬_KPIフォローシート\10_課別配布_テンプレート"

# ===== メイン2シート（表示する）=====
SHEET_FOLLOW = "月次KPIフォロー_現場入力"
SHEET_HQ     = "推進部_月次管理"

# ===== 必ず隠したい補助シート（表示されて困っているもの）=====
AUX_SHEETS_MUST_HIDE = [
    "Sheet1",
    "_lists",
    "_dv_lists",
    "_tpl_lists",
]

# ===== 追加で「参照維持のため残すかもしれない」補助シート候補（存在するものだけ残して隠す）=====
AUX_SHEETS_CANDIDATES = [
    "_lists", "_dv_lists", "_tpl_lists", "Sheet1",
    "月次KPIフォロー",
    "WS_サンプル回答",
]

# ===== データ開始行 =====
DATA_START_ROW = 3

# ===== 課の表示順（ベース名）=====
DEPT_ORDER = [
    "札幌", "仙台",
    "東京1", "東京2", "東京3", "東京4", "東京5",
    "名古屋1", "名古屋2",
    "大阪1", "大阪2",
    "広島",
    "福岡1", "福岡2",
]


# Excel定数（COMで数値指定する）
XL_SHEET_VISIBLE    = -1  # xlSheetVisible
XL_SHEET_HIDDEN     = 0   # xlSheetHidden
XL_SHEET_VERYHIDDEN = 2   # xlSheetVeryHidden

XL_TO_LEFT = -4159
XL_UP      = -4162


def safe_filename(name: str) -> str:
    name = str(name).strip()
    name = re.sub(r'[\\/:*?"<>|]', "_", name)
    return name[:120]


def normalize_dept(s: str) -> str:
    s = str(s).strip().replace("　", "").replace(" ", "")
    if s.endswith("課"):
        s = s[:-1]
    return s


def get_last_row_by_col(ws, col: int) -> int:
    return ws.Cells(ws.Rows.Count, col).End(XL_UP).Row


def get_cell_value_safe(ws, r: int, c: int):
    try:
        return ws.Cells(r, c).Value
    except Exception:
        return None


def find_dept_col_by_header(ws, header_row: int = 1, header_text: str = "課", fallback: int = 2) -> int:
    try:
        last_col = ws.Cells(header_row, ws.Columns.Count).End(XL_TO_LEFT).Column
    except Exception:
        last_col = 50

    for c in range(1, last_col + 1):
        v = get_cell_value_safe(ws, header_row, c)
        if v is None:
            continue
        if str(v).strip() == header_text:
            return c
    return fallback


def delete_rows_not_matching_dept(ws, dept_name_in_sheet: str, dept_col: int, start_row: int) -> None:
    last_row = get_last_row_by_col(ws, dept_col)
    for r in range(last_row, start_row - 1, -1):
        v = get_cell_value_safe(ws, r, dept_col)
        if v is None:
            continue
        s = str(v).strip()
        if not s:
            continue
        if s != dept_name_in_sheet:
            ws.Rows(r).Delete()


def sheet_exists(wb, sheet_name: str) -> bool:
    for i in range(1, wb.Worksheets.Count + 1):
        if str(wb.Worksheets(i).Name) == sheet_name:
            return True
    return False


def delete_other_sheets_keep(wb, keep_names: list[str]) -> None:
    keep = set(keep_names)
    for i in range(wb.Worksheets.Count, 0, -1):
        ws = wb.Worksheets(i)
        if ws.Name not in keep:
            ws.Delete()


def set_sheet_visible(wb, sheet_name: str) -> None:
    ws = wb.Worksheets(sheet_name)
    ws.Visible = XL_SHEET_VISIBLE


def set_sheet_very_hidden_safe(wb, sheet_name: str) -> None:
    """
    VeryHidden化。失敗した場合は Hidden(0) にフォールバック。
    """
    ws = wb.Worksheets(sheet_name)
    try:
        ws.Visible = XL_SHEET_VERYHIDDEN
    except Exception:
        # どうしてもダメなら通常Hiddenへ
        ws.Visible = XL_SHEET_HIDDEN


def main() -> None:
    if not os.path.exists(INPUT_XLSX):
        raise FileNotFoundError(f"入力ファイルが見つかりません: {INPUT_XLSX}")

    os.makedirs(OUT_DIR, exist_ok=True)

    pythoncom.CoInitialize()
    excel = None

    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = 0
        excel.DisplayAlerts = 0

        # ===== ① 元データから「課一覧」を抽出 =====
        wb_src = excel.Workbooks.Open(INPUT_XLSX, ReadOnly=True, UpdateLinks=0)

        for s in [SHEET_FOLLOW, SHEET_HQ]:
            if not sheet_exists(wb_src, s):
                names = [str(wb_src.Worksheets(i).Name) for i in range(1, wb_src.Worksheets.Count + 1)]
                wb_src.Close(SaveChanges=False)
                raise ValueError(
                    f"必要シートが見つかりません: {s}\n\nブック内のシート一覧:\n- " + "\n- ".join(names)
                )

        ws_src = wb_src.Worksheets(SHEET_FOLLOW)
        dept_col = find_dept_col_by_header(ws_src, header_row=1, header_text="課", fallback=2)
        last_row = get_last_row_by_col(ws_src, dept_col)

        dept_map_norm_to_raw = {}
        for r in range(DATA_START_ROW, last_row + 1):
            v = get_cell_value_safe(ws_src, r, dept_col)
            if v is None:
                continue
            raw = str(v).strip()
            if not raw:
                continue
            dept_map_norm_to_raw.setdefault(normalize_dept(raw), raw)

        wb_src.Close(SaveChanges=False)

        if not dept_map_norm_to_raw:
            raise ValueError("課のデータが見つかりません（課列の値が空の可能性）。")

        # ===== ② 母艦に存在する補助シート（実在のみ）を収集 =====
        wb_check = excel.Workbooks.Open(INPUT_XLSX, ReadOnly=True, UpdateLinks=0)
        existing_aux = []
        for s in (AUX_SHEETS_MUST_HIDE + AUX_SHEETS_CANDIDATES):
            if sheet_exists(wb_check, s):
                if s not in existing_aux:
                    existing_aux.append(s)
        wb_check.Close(SaveChanges=False)

        KEEP_SHEETS = [SHEET_FOLLOW, SHEET_HQ] + existing_aux

        # ===== ③ 課別ファイル生成 =====
        created = 0
        skipped = []

        for idx, dept_base in enumerate(DEPT_ORDER, start=1):
            norm = normalize_dept(dept_base)
            if norm not in dept_map_norm_to_raw:
                skipped.append(dept_base)
                continue

            dept_raw = dept_map_norm_to_raw[norm]
            out_name = f"{idx:02d}_{safe_filename(dept_raw)}_KPIフォローシート.xlsx"
            out_path = os.path.join(OUT_DIR, out_name)

            if os.path.exists(out_path):
                try:
                    os.remove(out_path)
                except PermissionError:
                    print(f"スキップ（ファイルがロック中）：{out_path}")
                    skipped.append(dept_base)
                    continue

            shutil.copy2(INPUT_XLSX, out_path)
            wb = excel.Workbooks.Open(out_path, ReadOnly=False, UpdateLinks=0)

            # (1) keep以外を削除
            delete_other_sheets_keep(wb, KEEP_SHEETS)

            # (2) 先に「見せたい2枚」を確実に Visible にして、フォローシートをアクティブ化
            set_sheet_visible(wb, SHEET_FOLLOW)
            set_sheet_visible(wb, SHEET_HQ)
            wb.Worksheets(SHEET_FOLLOW).Activate()

            # (3) フォローシート：対象課以外の行を削除
            ws_follow = wb.Worksheets(SHEET_FOLLOW)
            dept_col2 = find_dept_col_by_header(ws_follow, header_row=1, header_text="課", fallback=dept_col)
            delete_rows_not_matching_dept(ws_follow, dept_raw, dept_col2, DATA_START_ROW)

            # (4) 補助シートは VeryHidden（ダメならHiddenへフォールバック）
            for s in existing_aux:
                if s in (SHEET_FOLLOW, SHEET_HQ):
                    continue
                set_sheet_very_hidden_safe(wb, s)

            # (5) 念のため：再計算
            wb.Application.CalculateFullRebuild()

            wb.Save()
            wb.Close(SaveChanges=False)
            created += 1

        print("完了：2枚だけ表示（現場入力＋推進部_月次管理）／補助シートはVeryHiddenで課別テンプレを作成しました。")
        print(f"入力：{INPUT_XLSX}")
        print(f"出力フォルダ：{OUT_DIR}")
        print(f"作成数：{created} / 指定数：{len(DEPT_ORDER)}")
        if skipped:
            print("※データに存在せずスキップした課：", " / ".join(skipped))

    finally:
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    main()