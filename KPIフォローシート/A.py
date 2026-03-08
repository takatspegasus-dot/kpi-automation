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

# ===== 補助シート候補（参照維持のため残す：非表示にする）=====
AUX_SHEETS_CANDIDATES = [
    "_tpl_lists",
    "_lists",
    "_dv_lists",
    "Sheet1",
    "月次KPIフォロー",   # もし推進部_月次管理が参照している場合があるので保険
    "WS_サンプル回答",  # 参照している可能性があれば残す（不要なら消してOK）
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
    xlUp = -4162  # xlUp
    return ws.Cells(ws.Rows.Count, col).End(xlUp).Row


def get_cell_value_safe(ws, r: int, c: int):
    try:
        return ws.Cells(r, c).Value
    except Exception:
        return None


def find_dept_col_by_header(ws, header_row: int = 1, header_text: str = "課", fallback: int = 2) -> int:
    xlToLeft = -4159
    try:
        last_col = ws.Cells(header_row, ws.Columns.Count).End(xlToLeft).Column
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


def set_sheet_hidden(wb, sheet_name: str) -> None:
    """
    0:表示 / 1:非表示 / 2:再表示不可（very hidden）
    ここでは「非表示(1)」にする。運用で触らせたくなければ 2 も可。
    """
    ws = wb.Worksheets(sheet_name)
    ws.Visible = 1


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

        # 必須2シートの存在確認
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
            norm = normalize_dept(raw)
            dept_map_norm_to_raw.setdefault(norm, raw)

        wb_src.Close(SaveChanges=False)

        if not dept_map_norm_to_raw:
            raise ValueError("課のデータが見つかりません（課列の値が空の可能性）。")

        # ===== ② このブックに存在する補助シートを「実在するものだけ」keep対象に追加 =====
        #   → 参照が壊れないように、補助シートは残して非表示にする
        #   → 見えるのは2シート、内部的には複数残る
        wb_check = excel.Workbooks.Open(INPUT_XLSX, ReadOnly=True, UpdateLinks=0)
        existing_aux = [s for s in AUX_SHEETS_CANDIDATES if sheet_exists(wb_check, s)]
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

            shutil.copy2(INPUT_XLSX, out_path)
            wb = excel.Workbooks.Open(out_path, ReadOnly=False, UpdateLinks=0)

            # (1) keep以外を削除
            delete_other_sheets_keep(wb, KEEP_SHEETS)

            # (2) フォローシート：対象課以外の行を削除
            ws_follow = wb.Worksheets(SHEET_FOLLOW)
            dept_col2 = find_dept_col_by_header(ws_follow, header_row=1, header_text="課", fallback=dept_col)
            delete_rows_not_matching_dept(ws_follow, dept_raw, dept_col2, DATA_START_ROW)

            # (3) 補助シートは非表示（ユーザーには2枚だけ見せる）
            for s in existing_aux:
                # 2つのメインは表示のまま
                if s in (SHEET_FOLLOW, SHEET_HQ):
                    continue
                set_sheet_hidden(wb, s)

            # 念のため：再計算
            wb.Application.CalculateFullRebuild()

            wb.Save()
            wb.Close(SaveChanges=False)
            created += 1

        print("完了：2枚表示（現場入力＋推進部_月次管理）＋補助シート非表示で課別テンプレを作成しました。")
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