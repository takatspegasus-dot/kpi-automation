# -*- coding: utf-8 -*-
from __future__ import annotations

import re
from pathlib import Path
from datetime import datetime
import win32com.client as win32

# ====== 入出力フォルダ（ご指定） ======
INPUT_DIR = Path(
    r"C:\Users\SujiT\OneDrive\ドキュメント\HiproBiz わかもと製薬\KPI振り返りシート\02回収_課別配布ファイル"
)
OUTPUT_DIR = Path(
    r"C:\Users\SujiT\OneDrive\ドキュメント\HiproBiz わかもと製薬\KPI振り返りシート\03_統合_課別回収ファイル"
)
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

# 01〜14の回収ファイルを対象（命名規則）
FILE_RE = re.compile(r"^(0[1-9]|1[0-4])_.*_KPI振り返りシート_回収\.xlsx$", re.IGNORECASE)

# 出力ファイル名（YYYYMMDD）
OUT_NAME = f"KPI振り返り_課別統合_{datetime.now():%Y%m%d}.xlsx"


def pick_sheet_to_copy(wb):
    """
    コピー対象シートを選ぶ。
    - 案内シートは除外
    - '振り返り' を含むシートがあれば優先
    - なければ最初の（案内以外の）シート
    """
    excluded = {"保存方法_自動回収名"}

    for ws in wb.Worksheets:
        if ws.Name not in excluded and "振り返り" in ws.Name:
            return ws

    for ws in wb.Worksheets:
        if ws.Name not in excluded:
            return ws

    return wb.Worksheets(1)


def safe_sheet_name(name: str) -> str:
    """Excelのシート名制約に合わせる（禁止文字除去、31文字まで）"""
    name = re.sub(r'[:\\/?*\[\]]', "_", name)
    return name[:31]


def sheet_exists(wb, name: str) -> bool:
    for s in wb.Worksheets:
        if s.Name == name:
            return True
    return False


def unique_sheet_name(wb, base: str) -> str:
    """同名があれば _2, _3... を付けてユニーク化"""
    base = safe_sheet_name(base)
    if not sheet_exists(wb, base):
        return base

    k = 2
    while True:
        suffix = f"_{k}"
        cand = base[: max(0, 31 - len(suffix))] + suffix
        if not sheet_exists(wb, cand):
            return cand
        k += 1


def main():
    if not INPUT_DIR.exists():
        raise FileNotFoundError(f"入力フォルダが見つかりません: {INPUT_DIR}")

    files = sorted([p for p in INPUT_DIR.glob("*.xlsx") if FILE_RE.match(p.name)])
    if not files:
        raise FileNotFoundError(
            "対象ファイルが見つかりませんでした。\n"
            "例: 01_札幌_KPI振り返りシート_回収.xlsx のような命名になっているか確認してください。"
        )

    out_path = OUTPUT_DIR / OUT_NAME

    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    try:
        # 統合先ブック作成
        out_wb = excel.Workbooks.Add()
        out_wb_name = out_wb.Name

        # デフォルトシート（後で削除）
        default_sheets = [out_wb.Worksheets(i) for i in range(1, out_wb.Worksheets.Count + 1)]

        for f in files:
            print(f"取り込み: {f.name}")
            wb = excel.Workbooks.Open(str(f), ReadOnly=True)

            try:
                ws = pick_sheet_to_copy(wb)

                # ★ここが肝：Copy が「新規ブック」に飛ぶ環境対策込み
                out_wb.Activate()
                ws.Copy(None, out_wb.Worksheets(out_wb.Worksheets.Count))  # Before=None, After=最後

                active_wb = excel.ActiveWorkbook
                if active_wb.Name != out_wb_name:
                    # Copyが別ブック（新規ブック）に行ってしまった → そこから統合へ再コピー
                    temp_wb = active_wb
                    temp_ws = temp_wb.Worksheets(1)
                    out_wb.Activate()
                    temp_ws.Copy(None, out_wb.Worksheets(out_wb.Worksheets.Count))
                    temp_wb.Close(SaveChanges=False)

                # シート名：ファイル名から "01_札幌" を抽出
                m = re.match(r"^(\d{2})_(.+?)_KPI振り返りシート_回収\.xlsx$", f.name)
                base = f.stem
                if m:
                    base = f"{m.group(1)}_{m.group(2)}"

                copied = out_wb.Worksheets(out_wb.Worksheets.Count)
                copied.Name = unique_sheet_name(out_wb, base)

            finally:
                wb.Close(SaveChanges=False)

        # デフォルトシート削除
        for sh in default_sheets:
            try:
                sh.Delete()
            except Exception:
                pass

        # 保存（同名が既に開いている等で失敗することがあるので注意）
        out_wb.SaveAs(str(out_path))
        out_wb.Close(SaveChanges=False)

        print(f"統合完了: {out_path}")
        print(f"取り込みファイル数: {len(files)}")

    finally:
        excel.Quit()


if __name__ == "__main__":
    main()
