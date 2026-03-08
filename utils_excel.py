# -*- coding: utf-8 -*-
"""Excel操作スクリプト共通ユーティリティ"""
import re


def header_map(ws):
    """1行目のヘッダ名→列番号の辞書を返す"""
    m = {}
    for c in range(1, ws.max_column + 1):
        h = ws.cell(1, c).value
        if h not in (None, ""):
            m[str(h).strip()] = c
    return m


def norm(v):
    """値を文字列に正規化（None→空文字、/→-に統一）"""
    if v is None:
        return ""
    s = str(v).strip()
    s = s.replace("/", "-")
    return s


def safe_filename(name: str) -> str:
    """Windowsで使えない文字を除去し120文字以内に丸める"""
    name = str(name).strip()
    name = re.sub(r'[\\/:*?"<>|]', "_", name)
    return name[:120]


def safe_sheetname(name: str) -> str:
    """Excelのシート名制限対応（禁止文字除去・31文字以内）"""
    name = str(name).strip()
    name = re.sub(r"[\[\]\:\*\?\/\\]", "_", name)
    return name[:31]
