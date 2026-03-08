# -*- coding: utf-8 -*-
from __future__ import annotations

import os
import re
import shutil

import pythoncom
import win32com.client as win32


# ===== 入力（集計ファイル＝元データ/テンプレ元）=====
# ★実在するファイルパスに修正済み
INPUT_XLSX = r"C:\Users\SujiT\OneDrive\ドキュメント\HiproBiz わかもと製薬\KPI振り返りシート\09配布_課別ファイル\課別配布_研修テンプレ\提示用【テンプレ】KPI振り返りv3.xlsx"

# ===== 確定情報 =====
DEPT_COL = 2        # 課はB列
DATA_START_ROW = 7  # データ開始は7行目

# 課(B列)の右に並ぶ前提：KPI区分=C列、評価対象期間=D列（違えば直す）
KPI_TYPE_COL = 3
PERIOD_COL = 4

# ===== 出力 =====
BASE_DIR = os.path.dirname(INPUT_XLSX)
OUT_DIR = os.path.join(BASE_DIR, "課別配布_研修テンプレ")  # INPUTと同階層に出力
TEMPLATE_MASTER_NAME = "【研修テンプレ原本】KPI振り返り.xlsx"
TEMPLATE_MASTER_PATH = os.path.join(OUT_DIR, TEMPLATE_MASTER_NAME)

# ===== 課の表示順 =====
DEPT_ORDER = [
    "札幌", "仙台",
    "東京1", "東京2", "東京3", "東京4", "東京5",
    "名古屋1", "名古屋2",
    "大阪1", "大阪2",
    "広島",
    "福岡1", "福岡2",
]

# ===== プルダウン候補 =====
KPI_TYPE_LIST = ["発売前", "発売後"]
PERIOD_LIST = [
    "発売前（～11月末）",
    "発売後 第1Q（12-2月）",
    "発売後 第2Q（3-5月）",
    "発売後 第3Q（6-8月）",
    "発売後 第4Q（9-11月）",
]


def safe_filename(name: str) -> str:
    name = str(name).strip()
    name = re.sub(r'[\\/:*?"<>|]', "_", name)
    return name[:120] if len(name) > 120 else name


def set_list_validation(ws, col: int, row_start: int, row_end: int, options: list[str]) -> None:
    """
    指定列にリスト型データ検証（プルダウン）を設定
    """
    rng = ws.Range(ws.Cells(row_start, col), ws.Cells(row_end, col))
    try:
        rng.Validation.Delete()
    except Exception:
        pass

    formula = ",".join(options)  # 例: "発売前,発売後"

    # xlValidateList=3, xlValidAlertStop=1
    rng.Validation.Add(Type=3, AlertStyle=1, Operator=1, Formula1=formula)
    rng.Validation.IgnoreBlank = True
    rng.Validation.InCellDropdown = True


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


def rename_first_sheet_to_dept(wb, dept_name: str) -> None:
    """
    1枚目シート名を課名に（31文字/禁止文字対応＋重複回避）
    """
    ws = wb.Worksheets(1)

    name = str(dept_name).strip()
    name = re.sub(r"[\[\]\:\*\?\/\\]", "_", name)
    name = name[:31]

    existing = set()
    for i in range(1, wb.Worksheets.Count + 1):
        existing.add(wb.Worksheets(i).Name)

    if name in existing:
        base = name[:28]
        k = 2
        while True:
            cand = f"{base}_{k}"[:31]
            if cand not in existing:
                name = cand
                break
            k += 1

    ws.Name = name


def main() -> None:
    if not os.path.exists(INPUT_XLSX):
        raise FileNotFoundError(f"入力ファイルが見つかりません: {INPUT_XLSX}")

    # ★ INPUTが既に「課別配布_研修テンプレ」配下にあるため、
    # OUT_DIRが同一フォルダになり「フォルダの中に同名フォルダを作る」事故を防ぐ
    # OUT_DIR = BASE_DIR（INPUTと同じフォルダ）に固定する
    global OUT_DIR, TEMPLATE_MASTER_PATH
    OUT_DIR = BASE_DIR
    TEMPLATE_MASTER_PATH = os.path.join(OUT_DIR, TEMPLATE_MASTER_NAME)

    os.makedirs(OUT_DIR, exist_ok=True)

    pythoncom.CoInitialize()
    excel = None

    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = 0
        excel.DisplayAlerts = 0

        # ① テンプレ原本作成（入力ファイルをコピー）
        shutil.copy2(INPUT_XLSX, TEMPLATE_MASTER_PATH)

        # ② 原本にプルダウン設定
        wb = excel.Workbooks.Open(TEMPLATE_MASTER_PATH)
        ws = wb.Worksheets(1)

        last_row = get_last_row_by_col(ws, DEPT_COL)

        set_list_validation(ws, KPI_TYPE_COL, DATA_START_ROW, last_row, KPI_TYPE_LIST)
        set_list_validation(ws, PERIOD_COL, DATA_START_ROW, last_row, PERIOD_LIST)

        wb.Save()
        wb.Close(SaveChanges=False)

        # ③ 原本から課一覧取得
        wb_master = excel.Workbooks.Open(TEMPLATE_MASTER_PATH)
        ws_master = wb_master.Worksheets(1)
        last_row_master = get_last_row_by_col(ws_master, DEPT_COL)

        dept_in_data = set()
        for r in range(DATA_START_ROW, last_row_master + 1):
            v = get_cell_value_safe(ws_master, r, DEPT_COL)
            if v is None:
                continue
            s = str(v).strip()
            if s:
                dept_in_data.add(s)

        wb_master.Close(SaveChanges=False)

        # ④ 指定順で課別ファイル生成（同じフォルダに出す）
        created = 0
        skipped = []
        for idx, dept in enumerate(DEPT_ORDER, start=1):
            if dept not in dept_in_data:
                skipped.append(dept)
                continue

            out_name = f"{idx:02d}_{safe_filename(dept)}_KPI振り返り.xlsx"
            out_path = os.path.join(OUT_DIR, out_name)

            shutil.copy2(TEMPLATE_MASTER_PATH, out_path)

            wb_dept = excel.Workbooks.Open(out_path)
            ws_dept = wb_dept.Worksheets(1)

            delete_rows_not_matching_dept(ws_dept, dept, DEPT_COL, DATA_START_ROW)
            rename_first_sheet_to_dept(wb_dept, dept)

            wb_dept.Save()
            wb_dept.Close(SaveChanges=False)

            created += 1

        print("完了：課別配布ファイルを指定順で作成し、シート名も課名に変更しました。")
        print(f"テンプレ原本：{TEMPLATE_MASTER_PATH}")
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
