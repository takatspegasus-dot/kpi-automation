# -*- coding: utf-8 -*-
from __future__ import annotations

import os
import re
import shutil

import pythoncom
import win32com.client as win32


# ===== 入力（元ブック：サンプル回答入り）=====
INPUT_XLSX = r"C:\Users\SujiT\OneDrive\ドキュメント\HiproBiz わかもと製薬\KPI振り返りシート\20研修_KPI振り返りシート_最終案\WSテンプレート\KPI振り返り研修_WSシートVer1.0.xlsx"
# ===== 出力（課別配布サンプル）=====
OUT_DIR = r"C:\Users\SujiT\OneDrive\ドキュメント\HiproBiz わかもと製薬\KPI振り返りシート\20研修_KPI振り返りシート_最終案\WSテンプレート\課別配布サンプルシート_重点化フェーズ"

# ===== 対象シート名 =====
TARGET_SHEET_NAME = "重点化フェーズ"

# ===== 確定情報 =====
DEPT_COL = 2        # 課はB列
DATA_START_ROW = 7  # データ開始は7行目

# ===== 課の表示順 =====
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
    return name[:120] if len(name) > 120 else name


def get_last_row_by_col(ws, col: int) -> int:
    xlUp = -4162
    return ws.Cells(ws.Rows.Count, col).End(xlUp).Row


def get_cell_value_safe(ws, r: int, c: int):
    try:
        return ws.Cells(r, c).Value
    except Exception:
        return None


def delete_rows_not_matching_dept(ws, dept_name: str, dept_col: int, start_row: int) -> None:
    """
    指定dept以外の行を、下から順に削除
    """
    last_row = get_last_row_by_col(ws, dept_col)
    for r in range(last_row, start_row - 1, -1):
        v = get_cell_value_safe(ws, r, dept_col)
        if v is None:
            continue
        s = str(v).strip()
        if not s:
            continue
        if s != dept_name:
            ws.Rows(r).Delete()


def rename_sheet(ws, new_name: str) -> None:
    """
    シート名を課名に（31文字/禁止文字対応）
    """
    name = str(new_name).strip()
    name = re.sub(r"[\[\]\:\*\?\/\\]", "_", name)
    name = name[:31]
    ws.Name = name


def delete_other_sheets(wb, keep_name: str) -> None:
    """
    指定シート以外を削除（逆順で安全に削除）
    """
    # ExcelのSheetsは1-based
    for i in range(wb.Worksheets.Count, 0, -1):
        ws = wb.Worksheets(i)
        if ws.Name != keep_name:
            ws.Delete()


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

        # ① 元データから「課一覧（B列）」を抽出
        wb_src = excel.Workbooks.Open(INPUT_XLSX)
        try:
            ws_src = wb_src.Worksheets(TARGET_SHEET_NAME)
        except Exception:
            wb_src.Close(SaveChanges=False)
            raise ValueError(f"対象シートが見つかりません: {TARGET_SHEET_NAME}")

        last_row = get_last_row_by_col(ws_src, DEPT_COL)

        dept_in_data = set()
        for r in range(DATA_START_ROW, last_row + 1):
            v = get_cell_value_safe(ws_src, r, DEPT_COL)
            if v is None:
                continue
            s = str(v).strip()
            if s:
                dept_in_data.add(s)

        wb_src.Close(SaveChanges=False)

        # ② 課別ファイル生成（指定順）
        created = 0
        skipped = []

        for idx, dept in enumerate(DEPT_ORDER, start=1):
            if dept not in dept_in_data:
                skipped.append(dept)
                continue

            out_name = f"{idx:02d}_{safe_filename(dept)}_重点化フェーズ.xlsx"
            out_path = os.path.join(OUT_DIR, out_name)

            # 元ブックをコピーして加工（数式/書式/図形なども保持）
            shutil.copy2(INPUT_XLSX, out_path)

            wb = excel.Workbooks.Open(out_path)

            # ③ 対象シート以外削除
            delete_other_sheets(wb, TARGET_SHEET_NAME)

            ws = wb.Worksheets(TARGET_SHEET_NAME)

            # ④ 対象課以外の行を削除（B列で判定）
            delete_rows_not_matching_dept(ws, dept, DEPT_COL, DATA_START_ROW)

            # ⑤ シート名を課名へ
            rename_sheet(ws, dept)

            wb.Save()
            wb.Close(SaveChanges=False)

            created += 1

        print("完了：00_WS_サンプル回答 を課別に分割し、指定フォルダーへ保存しました。")
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
