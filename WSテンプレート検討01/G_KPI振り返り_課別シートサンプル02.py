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
SRC_XLSX = r"C:\Users\SujiT\OneDrive\ドキュメント\HiproBiz わかもと製薬\KPI振り返りシート\20研修_KPI振り返りシート_最終案\WSテンプレート\KPI振り返り_WSシート02.xlsx"

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

# 追加する「次フェーズKPI」行数（=3行固定）
ADD_NEXT_KPI_ROWS = 3

# 追加行に入れる値
NEXT_PHASE_KPI_KUBUN = "発売後"
NEXT_PHASE_PERIOD = "～6月末"

# S列（次フェーズ反映）に入れる値
NEXT_PHASE_FLAG_VALUE = "○"

# 既存行まで一括流し込みするか（事故防止でデフォルトOFF）
FILL_EXISTING_ROWS_TOO = False


# =========================
# KPIサンプル（課別：3本セット）※期間表記・期限表記は入れない
# =========================
KPI_TEXTS_BY_KA: Dict[str, List[str]] = {
    "札幌": [
        """KPI：重点薬局でLX別規格教育を実施し、誤指導をゼロにする
指標：①教育実施件数 ②誤指導・問い合わせ件数
目標：①月12件 ②月0件""",
        """KPI：重点施設で新規採用（処方開始）を創出する
指標：①新規採用施設数 ②意思決定者面談数
目標：①月1施設 ②月8回""",
        """KPI：重点施設で欠品ゼロ運用を維持する（在庫確認・補充提案）
指標：①欠品報告件数 ②在庫確認・補充提案件数
目標：①月0件 ②月12件""",
    ],
    "仙台": [
        """KPI：重点施設で新規採用（処方開始）を創出する
指標：①新規採用施設数 ②意思決定者面談数
目標：①月1施設 ②月8回""",
        """KPI：重点薬局でLX別規格教育を徹底し、誤指導をゼロにする
指標：①教育実施件数 ②誤指導・問い合わせ件数
目標：①月10件 ②月0件""",
        """KPI：既採用施設の処方継続を維持する
指標：①処方継続率 ②フォロー面談数
目標：①80％以上 ②月6回""",
    ],
    "東京1課": [
        """KPI：重点医療機関で新規採用を継続的に創出する
指標：①新規採用施設数 ②医師＋門前薬局セット面談数
目標：①月3施設 ②月10セット""",
        """KPI：重点薬局でLX別規格教育を徹底し、誤指導をゼロにする
指標：①教育実施件数 ②誤指導・問い合わせ件数
目標：①月20件 ②月0件""",
        """KPI：既採用施設で処方継続を維持し、中止理由を潰す
指標：①処方継続率 ②中止理由回収・対策完了率
目標：①80％以上 ②100％""",
    ],
    "東京2課": [
        """KPI：既採用施設で処方継続を維持し、中止理由を潰す
指標：①処方継続率 ②中止理由回収・対策完了率
目標：①80％以上 ②100％""",
        """KPI：重点エリアで医師→門前薬局の連動活動を実行する
指標：①連動セット実施数 ②セット後の採用・継続率
目標：①月8セット ②70％以上""",
        """KPI：重点施設で欠品ゼロ運用を維持する（在庫確認・補充提案）
指標：①欠品報告件数 ②在庫確認・補充提案件数
目標：①月0件 ②月15件""",
    ],
    "東京3課": [
        """KPI：重点薬局でLX別規格教育を徹底し、誤指導をゼロにする
指標：①教育実施件数 ②誤指導・問い合わせ件数
目標：①月20件 ②月0件""",
        """KPI：比較局面での切り返しトークを現場で再現する
指標：①切り返し実践記録数 ②処方継続率
目標：①月15回 ②80％以上""",
        """KPI：重点施設で新規採用（処方開始）を創出する
指標：①新規採用施設数 ②意思決定者面談数
目標：①月2施設 ②月10回""",
    ],
    "東京4課": [
        """KPI：比較局面での切り返しトークを現場で再現する
指標：①切り返し実践記録数 ②処方継続率
目標：①月15回 ②80％以上""",
        """KPI：重点薬局でLX別規格教育を実施し、誤指導をゼロにする
指標：①教育実施件数 ②誤指導・問い合わせ件数
目標：①月18件 ②月0件""",
        """KPI：重点医療機関で新規採用（処方開始）を創出する
指標：①新規採用施設数 ②意思決定者面談数
目標：①月2施設 ②月10回""",
    ],
    "東京5課": [
        """KPI：重点エリアで医師→門前薬局の連動活動を実行する
指標：①連動セット実施数 ②セット後の採用・継続率
目標：①月8セット ②70％以上""",
        """KPI：重点医療機関で新規採用（処方開始）を創出する
指標：①新規採用施設数 ②医師面談数
目標：①月3施設 ②月12回""",
        """KPI：既採用施設で処方継続を維持し、中止理由を潰す
指標：①処方継続率 ②中止理由回収・対策完了率
目標：①80％以上 ②100％""",
    ],
    "名古屋1課": [
        """KPI：重点施設で新規採用と処方継続を両立させる
指標：①新規採用施設数 ②処方継続率
目標：①月2施設 ②80％以上""",
        """KPI：重点薬局でLX別規格教育を実施し、誤指導をゼロにする
指標：①教育実施件数 ②誤指導・問い合わせ件数
目標：①月15件 ②月0件""",
        """KPI：既採用施設の処方継続を維持する
指標：①処方継続率 ②フォロー面談数
目標：①80％以上 ②月8回""",
    ],
    "名古屋2課": [
        """KPI：薬局教育を通じて誤用リスクを低減する
指標：①教育実施件数 ②誤指導・問い合わせ件数
目標：①月15件 ②月0件""",
        """KPI：重点施設で新規採用（処方開始）を創出する
指標：①新規採用施設数 ②意思決定者面談数
目標：①月2施設 ②月8回""",
        """KPI：重点施設で欠品ゼロ運用を維持する（在庫確認・補充提案）
指標：①欠品報告件数 ②在庫確認・補充提案件数
目標：①月0件 ②月15件""",
    ],
    "大阪1課": [
        """KPI：比較される局面でLX優位の判断を獲得する
指標：①切り返し実践件数 ②新規採用施設数
目標：①月12回 ②月2施設""",
        """KPI：重点薬局でLX別規格教育を徹底し、誤指導をゼロにする
指標：①教育実施件数 ②誤指導・問い合わせ件数
目標：①月18件 ②月0件""",
        """KPI：既採用施設で処方継続を維持する
指標：①処方継続率 ②フォロー面談数
目標：①80％以上 ②月10回""",
    ],
    "大阪2課": [
        """KPI：既採用施設の処方継続を安定させる
指標：①処方継続率 ②フォロー面談数
目標：①80％以上 ②月10回""",
        """KPI：重点施設で新規採用（処方開始）を創出する
指標：①新規採用施設数 ②意思決定者面談数
目標：①月2施設 ②月8回""",
        """KPI：重点薬局でLX別規格教育を実施し、誤指導をゼロにする
指標：①教育実施件数 ②誤指導・問い合わせ件数
目標：①月15件 ②月0件""",
    ],
    "広島課": [
        """KPI：重点施設を絞り込み、確実な採用と継続を作る
指標：①重点施設での採用・継続件数 ②重点面談数
目標：①月1施設 ②月6回""",
        """KPI：重点薬局でLX別規格教育を実施し、誤指導をゼロにする
指標：①教育実施件数 ②誤指導・問い合わせ件数
目標：①月10件 ②月0件""",
        """KPI：既採用施設で処方継続を維持し、中止理由を潰す
指標：①処方継続率 ②中止理由回収・対策完了率
目標：①80％以上 ②100％""",
    ],
    "福岡1課": [
        """KPI：新規採用の拡大と薬局教育を同時に進める
指標：①新規採用施設数 ②薬局教育実施件数
目標：①月2施設 ②月15件""",
        """KPI：既採用施設の処方継続を維持する
指標：①処方継続率 ②フォロー面談数
目標：①80％以上 ②月8回""",
        """KPI：重点施設で欠品ゼロ運用を維持する（在庫確認・補充提案）
指標：①欠品報告件数 ②在庫確認・補充提案件数
目標：①月0件 ②月15件""",
    ],
    "福岡2課": [
        """KPI：重点薬局での教育を通じ、誤用ゼロ運用を定着させる
指標：①教育実施件数 ②誤指導・問い合わせ件数
目標：①月15件 ②月0件""",
        """KPI：重点施設で新規採用（処方開始）を創出する
指標：①新規採用施設数 ②意思決定者面談数
目標：①月2施設 ②月8回""",
        """KPI：既採用施設で処方継続を維持する
指標：①処方継続率 ②フォロー面談数
目標：①80％以上 ②月8回""",
    ],
}


# =========================
# Excel定数（win32com）
# =========================
xlPasteFormats = -4122
xlPasteValidation = 6
xlLeft = -4131   # 左揃え
xlTop = -4160    # 上揃え


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


def find_header_row(ws, max_scan: int = 80) -> int:
    for r in range(1, max_scan + 1):
        row_texts = []
        for c in range(1, 61):
            s = norm(ws.Cells(r, c).Value)
            if s:
                row_texts.append(s)
        joined = " ".join(row_texts)
        if ("課" in joined) and ("KPI" in joined and "区分" in joined) and ("評価" in joined and "対象" in joined):
            return r
    return 4


def detect_columns(ws, header_row: int) -> Dict[str, int]:
    found: Dict[str, int] = {}
    for c in range(1, 61):
        s = norm(ws.Cells(header_row, c).Value)
        if not s:
            continue
        s2 = s.replace("\n", "")

        if s2 == "課":
            found["ka"] = c
        if ("KPI" in s2) and ("区分" in s2):
            found["kubun"] = c
        if ("評価" in s2) and ("対象" in s2) and ("期間" in s2):
            found["period"] = c
        if s2 == "KPI":
            found["kpi"] = c
        if ("次" in s2 and "フェーズ" in s2 and "反映" in s2) or ("次フェーズ" in s2 and "反映" in s2):
            found["next_phase"] = c

    found.setdefault("ka", 2)
    found.setdefault("kubun", 3)
    found.setdefault("period", 4)
    found.setdefault("kpi", 5)
    found.setdefault("next_phase", 19)  # S列
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
    ws.Rows(src_row).Copy()
    ws.Rows(dst_row).PasteSpecial(Paste=xlPasteFormats)
    ws.Rows(dst_row).PasteSpecial(Paste=xlPasteValidation)


def safe_save_close(wb) -> None:
    last_err = None
    for _ in range(5):
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


def is_reflect_flag(v) -> bool:
    if v is True:
        return True
    s = norm(v)
    return s in {"○", "〇", "1", "TRUE", "True"}


def set_kpi_cell(cell, text: str) -> None:
    """
    KPIセルに値を入れ、表示を
    - 左揃え
    - 上揃え
    - 折り返し
    に統一する
    """
    cell.Value = text
    cell.HorizontalAlignment = xlLeft
    cell.VerticalAlignment = xlTop
    cell.WrapText = True


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
            col_next = cols["next_phase"]

            target_key = ka_key(ka)

            # --- 課一致の行だけ残す／真テンプレ行は残す ---
            last_row = last_used_row(ws)
            for r in range(last_row, header_row, -1):
                if is_true_template_row(ws, r, col_ka, col_kpi, header_row):
                    continue
                ka_in_sheet = ka_key(ws.Cells(r, col_ka).Value)
                if ka_in_sheet != target_key:
                    ws.Rows(r).Delete()

            # --- No/課名の再付与 ---
            last_row2 = last_used_row(ws)
            for r in range(header_row + 1, last_row2 + 1):
                if is_true_template_row(ws, r, col_ka, col_kpi, header_row):
                    continue
                if row_has_any_value(ws, r, 2, 19):
                    ws.Cells(r, 1).Value = i
                    ws.Cells(r, col_ka).Value = ka

            # --- 次フェーズKPI入力行（3行）を追加 ---
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

            # ★追加3行に入れる「別KPI」3本
            kpi_list = KPI_TEXTS_BY_KA.get(ka, ["", "", ""])
            if len(kpi_list) < ADD_NEXT_KPI_ROWS:
                kpi_list = (kpi_list + [""] * ADD_NEXT_KPI_ROWS)[:ADD_NEXT_KPI_ROWS]

            for idx, rr in enumerate(range(insert_at, insert_at + ADD_NEXT_KPI_ROWS)):
                copy_formats_and_validation_from_row(ws, src_format_row, rr)
                clear_row_values(ws, rr, 1, 19)

                ws.Cells(rr, 1).Value = i
                ws.Cells(rr, col_ka).Value = ka
                ws.Cells(rr, col_kubun).Value = NEXT_PHASE_KPI_KUBUN
                ws.Cells(rr, col_period).Value = NEXT_PHASE_PERIOD
                ws.Cells(rr, col_next).Value = NEXT_PHASE_FLAG_VALUE

                # ★KPI記入＋（左揃え・上揃え・折り返し）
                set_kpi_cell(ws.Cells(rr, col_kpi), kpi_list[idx])

            # （任意）既存行にも反映したい場合
            if FILL_EXISTING_ROWS_TOO:
                last_row4 = last_used_row(ws)
                for r in range(header_row + 1, last_row4 + 1):
                    if insert_at <= r < insert_at + ADD_NEXT_KPI_ROWS:
                        continue  # 追加行はすでに埋めている
                    if is_true_template_row(ws, r, col_ka, col_kpi, header_row):
                        continue

                    period = norm(ws.Cells(r, col_period).Value)
                    flag = ws.Cells(r, col_next).Value
                    if period == norm(NEXT_PHASE_PERIOD) and is_reflect_flag(flag):
                        # 既存行は「1本目」を入れる（必要なら振り分けロジックに変更可能）
                        set_kpi_cell(ws.Cells(r, col_kpi), kpi_list[0])

            # --- シート名 ---
            new_sheet_name = safe_sheet_title(f"{ka}_①②③振り返り（研修用）")
            ws.Name = new_sheet_name

            safe_save_close(wb)
            print("作成完了:", out_path, " / sheet:", new_sheet_name)

    finally:
        excel.Quit()


if __name__ == "__main__":
    main()
