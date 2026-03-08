# -*- coding: utf-8 -*-
from __future__ import annotations

import os
import shutil
import sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from utils_excel import safe_filename, safe_sheetname

import pythoncom
import win32com.client as win32


# ===== 入力（母艦テンプレート）=====
INPUT_XLSX = r"C:\Users\SujiT\OneDrive\ドキュメント\HiproBiz わかもと製薬\2026.03.27 わかもと製薬　KPI振り返り研修フォローアップ\03_27事前整理シート_template.xlsx"

# ===== 出力（課別テンプレート保存先）=====
OUT_DIR = r"C:\Users\SujiT\OneDrive\ドキュメント\HiproBiz わかもと製薬\2026.03.27 わかもと製薬　KPI振り返り研修フォローアップ\10_課別_事前整理シート_テンプレート"


# ===== 課順 =====
DEPT_ORDER = [
    "札幌", "仙台",
    "東京1", "東京2", "東京3", "東京4", "東京5",
    "名古屋1", "名古屋2",
    "大阪1", "大阪2",
    "広島",
    "福岡1", "福岡2",
]

# ===== シート名（母艦の先頭シート名）=====
BASE_SHEET_NAME = "03_27事前整理"

# ===== 課名差し込みセル（B4想定）=====
DEPT_CELL = (4, 2)  # row=4, col=2 → B4

# ===== 任意：空欄にするセル =====
CLEAR_CELLS = [
    (4, 5),  # E4 記入者
    (5, 2),  # B5 記入日
]



def main() -> None:
    if not os.path.exists(INPUT_XLSX):
        raise FileNotFoundError(f"入力ファイルが見つかりません: {INPUT_XLSX}")

    os.makedirs(OUT_DIR, exist_ok=True)

    pythoncom.CoInitialize()
    excel = None

    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False
        excel.AskToUpdateLinks = False

        created = 0

        for idx, dept in enumerate(DEPT_ORDER, start=1):
            dept_full = f"{dept}課"
            out_name = f"{idx:02d}_{safe_filename(dept_full)}_KPI研修フォローアップシート.xlsx"
            out_path = os.path.join(OUT_DIR, out_name)

            # 既存ファイルがあれば上書きのため削除
            if os.path.exists(out_path):
                os.remove(out_path)

            # 母艦をコピー
            shutil.copy2(INPUT_XLSX, out_path)

            wb = excel.Workbooks.Open(out_path, UpdateLinks=0, ReadOnly=False)

            try:
                # シート取得（名前優先、なければ1枚目）
                try:
                    ws = wb.Worksheets(BASE_SHEET_NAME)
                except Exception:
                    ws = wb.Worksheets(1)

                # ① 課名差し込み
                ws.Cells(DEPT_CELL[0], DEPT_CELL[1]).Value = dept_full

                # ② 指定セルを空欄化
                for r, c in CLEAR_CELLS:
                    ws.Cells(r, c).Value = ""

                # ③ シート名変更
                new_sheet_name = safe_sheetname(dept_full)
                try:
                    ws.Name = new_sheet_name
                except Exception:
                    # 同名等の万一に備えて連番付与
                    ws.Name = safe_sheetname(f"{dept_full}_{idx}")

                # ④ 保存
                wb.Save()

            finally:
                wb.Close(SaveChanges=False)

            created += 1
            print(f"作成完了: {out_path}")

        print("========================================")
        print("完了：課別テンプレートを作成しました。")
        print(f"出力フォルダ：{OUT_DIR}")
        print(f"作成数：{created}")
        print("========================================")

    finally:
        if excel is not None:
            try:
                excel.ScreenUpdating = True
                excel.EnableEvents = True
                excel.DisplayAlerts = True
                excel.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    main()