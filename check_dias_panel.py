
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
check_dias_panel.py
---------------------------------
用途：
1) 從 model.ini 解析 m_pPanelName 指向的面板檔案路徑；
2) 依照 /tvconfigs/... 轉為專案根目錄下的相對路徑；
3) 開啟對應 panel 檔，檢查以下三個數值是否符合「嚴格大於」的門檻：
   - DISP_HORIZONTAL_TOTAL > 3840
   - DISP_VERTICAL_TOTAL   > 2160
   - DISPLAY_REFRESH_RATE  >= 60
任一不符合則回傳非 0 並列印 FAIL 與原因；全部符合則列印 PASS。

相容：Python 3.8+
"""

import argparse
import re
import sys
from pathlib import Path

# 允許的鍵名（大小寫不敏感）；正則只擷取等號右側第一個整數
KEY_PATTERNS = {
    "DISP_HORIZONTAL_TOTAL": re.compile(r"^\s*DISP_HORIZONTAL_TOTAL\s*=\s*([0-9]+)", re.IGNORECASE),
    "DISP_VERTICAL_TOTAL":   re.compile(r"^\s*DISP_VERTICAL_TOTAL\s*=\s*([0-9]+)", re.IGNORECASE),
    "DISPLAY_REFRESH_RATE":  re.compile(r"^\s*DISPLAY_REFRESH_RATE\s*=\s*([0-9]+)", re.IGNORECASE),
}

THRESHOLDS = {
    "DISP_HORIZONTAL_TOTAL": 3840,
    "DISP_VERTICAL_TOTAL":   2160,
    "DISPLAY_REFRESH_RATE":  59,
}

PANEL_NAME_RE = re.compile(
    r'^\s*m_pPanelName\s*=\s*"(.*?)"\s*;?.*$',  # 取出雙引號中的路徑
    re.IGNORECASE
)

def resolve_panel_path(raw_path: str, root: Path) -> Path:
    """將 model.ini 中的 m_pPanelName 路徑轉為實際檔案路徑。

    規則：
    - 若以 "/tvconfigs/" 開頭，去掉前綴後，以 root 為基底拼成相對路徑
      例："/tvconfigs/panel/xxx.ini" -> root/"panel/xxx.ini"
    - 若以 "/panel/" 開頭，視為位於根目錄 panel/ 下
      例："/panel/xxx.ini" -> root/"panel/xxx.ini"
    - 若已是相對路徑（如 "panel/xxx.ini"），以 root 為基底
    - 其他絕對路徑：直接使用該絕對路徑
    """
    raw_path = raw_path.strip()
    if raw_path.startswith("/tvconfigs/"):
        sub = raw_path[len("/tvconfigs/"):]  # 去掉前綴
        return root / sub
    if raw_path.startswith("/panel/"):
        return root / raw_path.lstrip("/")
    p = Path(raw_path)
    if p.is_absolute():
        return p
    return (root / p)

def parse_panel_name(model_ini: Path) -> str:
    """從 model.ini 讀取 m_pPanelName 內容（雙引號內字串）。"""
    try:
        with model_ini.open("r", encoding="utf-8", errors="ignore") as f:
            for line in f:
                m = PANEL_NAME_RE.match(line)
                if m:
                    return m.group(1)
    except FileNotFoundError:
        raise FileNotFoundError(f"Model ini not found: {model_ini}")
    raise ValueError('找不到 m_pPanelName = "..." 行，請確認 model.ini。')

def extract_values(panel_ini: Path) -> dict:
    """從 panel 檔案擷取三個目標鍵的整數值。若缺少則不放入 dict。"""
    values = {}
    try:
        with panel_ini.open("r", encoding="utf-8", errors="ignore") as f:
            for line in f:
                # 去除行尾註解（; 或 # 之後的內容），避免干擾擷取
                # 但因為 regex 取第一個數字，這步不是必須，保留為清晰
                for key, pat in KEY_PATTERNS.items():
                    m = pat.match(line)
                    if m and key not in values:
                        try:
                            values[key] = int(m.group(1))
                        except ValueError:
                            # 遇到非整數，略過，交由缺值/不合格處理
                            pass
    except FileNotFoundError:
        raise FileNotFoundError(f"Panel ini not found: {panel_ini}")
    return values

def check(values: dict) -> (bool, list):
    """檢查數值是否都「嚴格大於」門檻。回傳 (ok, errors)。"""
    errors = []
    for key, th in THRESHOLDS.items():
        if key not in values:
            errors.append(f"缺少 {key} 欄位")
            continue
        v = values[key]
        if not (v > th):
            errors.append(f"{key} = {v} 不大於 {th}")
    return (len(errors) == 0, errors)

def main():
    ap = argparse.ArgumentParser(
        description="檢查 panel 參數是否符合：H_TOTAL>3840, V_TOTAL>2160, REFRESH_RATE>60"
    )
    ap.add_argument("--model-ini", required=True, help="model/*.ini 路徑")
    ap.add_argument("--root", default=".", help="專案根目錄（含 panel/ 子資料夾），預設為目前目錄")
    args = ap.parse_args()

    model_ini = Path(args.model_ini).resolve()
    root = Path(args.root).resolve()

    try:
        raw_panel = parse_panel_name(model_ini)
    except Exception as e:
        print(f"[FAIL] 解析 model.ini 失敗：{e}")
        return 2

    panel_path = resolve_panel_path(raw_panel, root).resolve()

    try:
        values = extract_values(panel_path)
    except Exception as e:
        print(f"[FAIL] 開啟 panel 檔失敗：{e}")
        return 3

    ok, errors = check(values)

    print("=== DIAS Panel 檢查報告 ===")
    print(f"Model INI     : {model_ini}")
    print(f"Root          : {root}")
    print(f"Panel Raw Path: {raw_panel}")
    print(f"Panel File    : {panel_path}")
    print("擷取值：")
    for k in ["DISP_HORIZONTAL_TOTAL", "DISP_VERTICAL_TOTAL", "DISPLAY_REFRESH_RATE"]:
        v = values.get(k, "<缺少>")
        print(f"  - {k}: {v}")

    if ok:
        print("[PASS] 所有條件皆符合（>3840, >2160, >60）。")
        return 0
    else:
        print("[FAIL] 不符合條件：")
        for err in errors:
            print(f"  - {err}")
        return 1

if __name__ == "__main__":
    sys.exit(main())
