#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
check_cltv_report.py

依照 check_cltv.py 的檢查邏輯：
  1) 讀取 model.ini，尋找 LaunchCLTVByCountry = "<path>"
  2) 將 /tvconfigs/ 前綴路徑映射到 --root 目錄下的實際檔案路徑
  3) 檢查該檔案是否存在

最後的輸出格式、報表產生方式參考 target_country_check.py：
  - 提供 --report 與 --report-xlsx 參數，輸出 Excel（每個 PID_* 分頁，首列粗體、欄位換行與垂直置頂、無值填 N/A）
  - 欄位：
      A: Result (PASS / FAIL / N/A)
      B: condition_1 → LaunchCLTVByCountry = "<原始設定值或 N/A>"
      C: condition_2 → Resolved Path = "<映射後的實際路徑或 N/A>"
      D: condition_3 → File Exists = <Yes/No/N/A>
      E: condition_4 → Model.ini = "<model.ini 路徑>"
      （其餘 condition_* 留空，預設共 8 欄 condition_*，可用 --conditions 指定數量）
"""
import argparse
import os
import re
from typing import Optional

# -----------------------------
# Utilities for report (aligned with target_country_check.py style)
# -----------------------------
def _sheet_name_for_model(model_ini_path: str) -> str:
    import os, re
    base = os.path.basename(model_ini_path or "")
    m = re.match(r"^(\d+)_", base)
    if m:
        return f"PID_{int(m.group(1))}"
    return "others"


def _ensure_openpyxl():
    try:
        import openpyxl  # noqa
    except ImportError:
        raise SystemExit(
            "[ERROR] 需要 openpyxl 以支援報表輸出與附加。\n"
            "  安裝： pip install --user openpyxl\n"
        )


def export_report(res: dict, xlsx_path: str = "kipling.xlsx", num_condition_cols: int = 8) -> None:
    """
    欄位：
      A: Result (PASS/FAIL/N/A)
      B..I: condition_1..condition_8 （可由 num_condition_cols 調整數量）
    無值時以 'N/A' 填入。依 model.ini 檔名前綴分頁（PID_1、PID_2…；非數字→others），既有資料則附加。
    """
    _ensure_openpyxl()
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Alignment, Font
    from openpyxl.utils import get_column_letter

    def _na(s: str) -> str:
        s = (s or "").strip()
        return s if s else "N/A"

    sheet_name = _sheet_name_for_model(res.get("model_ini", ""))

    # 開啟或新建 xlsx
    try:
        wb = load_workbook(xlsx_path)
    except Exception:
        wb = Workbook()

    # 取得或建立工作表與表頭
    if sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
    else:
        ws = wb.create_sheet(title=sheet_name)
        headers = ["Result"] + [f"condition_{i}" for i in range(1, num_condition_cols + 1)]
        ws.append(headers)

    # 取值
    result     = _na(res.get("result", ""))
    conditions = res.get("conditions", [])
    # 轉 N/A、補足長度
    conditions = [ _na(x) for x in conditions ]
    if len(conditions) < num_condition_cols:
        conditions += [""] * (num_condition_cols - len(conditions))

    # 寫入一列
    row = [result] + conditions[:num_condition_cols]
    ws.append(row)

    # 第一列粗體
    bold_font = Font(bold=True)
    for cell in ws['1']:
        cell.font = bold_font

    # 欄寬與對齊
    last_row_idx = ws.max_row

    # Result 垂直靠上
    ws.cell(row=last_row_idx, column=1).alignment = Alignment(vertical="top")

    # condition_* 欄位：寬、換行、垂直靠上
    for col_idx in range(2, 2 + num_condition_cols):
        col_letter = get_column_letter(col_idx)
        ws.column_dimensions[col_letter].width = 80
        ws.cell(row=last_row_idx, column=col_idx).alignment = Alignment(
            wrap_text=True, vertical="top"
        )

    # 移除預設空白 Sheet（若存在且非唯一）
    if "Sheet" in wb.sheetnames and len(wb.sheetnames) > 1:
        try:
            wb.remove(wb["Sheet"])
        except Exception:
            pass

    wb.save(xlsx_path)


# -----------------------------
# Core logic (based on check_cltv.py semantics)
# -----------------------------
def _read_text_lines(path: str):
    for enc in ("utf-8", "latin-1", "utf-16"):
        try:
            with open(path, "r", encoding=enc, errors="ignore") as f:
                return f.readlines()
        except UnicodeDecodeError:
            continue
        except FileNotFoundError:
            raise
    with open(path, "r") as f:
        return f.readlines()


def _strip_comment(line: str) -> str:
    # 僅去掉行首 # 註解（保留與原檔接近的寬鬆判斷）
    return line if not line.lstrip().startswith("#") else ""


def _resolve_tvconfigs_path(root: str, tvconfigs_like: str) -> str:
    """
    把 "/tvconfigs/xxx/yyy" 映射為 "<root>/xxx/yyy"
    其他相對路徑: 以 root 為基底
    絕對路徑（非 /tvconfigs 開頭）維持不動
    """
    tvconfigs_like = tvconfigs_like.strip()
    if tvconfigs_like.startswith("/tvconfigs/"):
        rel = tvconfigs_like[len("/tvconfigs/"):]
        return os.path.normpath(os.path.join(root, rel))
    if tvconfigs_like.startswith("./") or tvconfigs_like.startswith("../"):
        return os.path.normpath(os.path.join(root, tvconfigs_like))
    if tvconfigs_like.startswith("/"):
        # 非 /tvconfigs 絕對路徑：原樣返回
        return tvconfigs_like
    # 其他當作相對於 root
    return os.path.normpath(os.path.join(root, tvconfigs_like))


def parse_model_ini_for_launch_cltv(model_ini_path: str) -> Optional[str]:
    """
    從 model.ini 找：第一條未被 # 註解的 LaunchCLTVByCountry = "<value>"
    回傳原始 value（可含 /tvconfigs 前綴）；若未找到回傳 None。
    """
    lines = _read_text_lines(model_ini_path)
    for raw in lines:
        line = _strip_comment(raw)
        if not line:
            continue
        if "LaunchCLTVByCountry" in line and "=" in line:
            m = re.search(r'LaunchCLTVByCountry\s*=\s*"([^"]*)"', line)
            if m:
                return m.group(1).strip()
            # 容忍未加引號
            m2 = re.search(r'LaunchCLTVByCountry\s*=\s*(\S+)', line)
            if m2:
                return m2.group(1).strip()
            # 格式不正確時，回傳空字串以供後續 FAIL 判定
            return ""
    return None


def main():
    parser = argparse.ArgumentParser(description="Check LaunchCLTVByCountry in model.ini and export report (target_country_check style).")
    parser.add_argument("--root", required=True, help="專案根目錄路徑（將 /tvconfigs/* 映射到此）")
    parser.add_argument("--model-ini", required=True, help="model.ini 檔案路徑（例如 model/1_xxx.ini）")
    parser.add_argument("--report", action="store_true", help="輸出報表到 xlsx（預設檔名 kipling.xlsx）")
    parser.add_argument("--report-xlsx", metavar="FILE", help="指定報表 xlsx 檔案名稱")
    parser.add_argument("--conditions", type=int, default=8, help="condition_* 欄位數，預設 8")
    parser.add_argument("-v", "--verbose", action="store_true", help="顯示詳細過程")
    args = parser.parse_args()

    model_ini = args.model_ini
    if not os.path.exists(model_ini):
        raise SystemExit(f"[ERROR] model ini not found: {model_ini}")

    root = os.path.abspath(os.path.normpath(args.root))

    if args.verbose:
        print(f"[INFO] model_ini: {model_ini}")
        print(f"[INFO] root     : {root}")

    # 1) 擷取設定值
    raw_value = parse_model_ini_for_launch_cltv(model_ini)

    if raw_value is None:
        if args.verbose:
            print("[INFO] LaunchCLTVByCountry 未宣告（或只有註解）→ N/A")
        result = "N/A"
        resolved = ""
        exists_text = "N/A"
    else:
        if raw_value == "":
            if args.verbose:
                print("[WARN] LaunchCLTVByCountry 格式錯誤或值為空 → FAIL")
            result = "FAIL"
            resolved = ""
            exists_text = "N/A"
        else:
            # 2) 映射到實際路徑
            resolved = _resolve_tvconfigs_path(root, raw_value)
            # 3) 檢查檔案存在
            exists = os.path.exists(resolved)
            exists_text = "Yes" if exists else "No"
            result = "PASS" if exists else "FAIL"

    # 螢幕輸出（比對步驟）
    print("=== LaunchCLTVByCountry Check ===")
    print(f"Model.ini : {model_ini}")
    print(f"Setting   : {raw_value if raw_value is not None else 'N/A'}")
    print(f"Resolved  : {resolved if resolved else 'N/A'}")
    print(f"Exists?   : {exists_text}")
    print(f"Result    : {result}")

    # 準備報表資料
    conditions = [
        f'LaunchCLTVByCountry exist?',
        f'LaunchCLTVByCountry = {raw_value if raw_value is not None and raw_value != "" else "N/A"}',
        f"File Exists = {exists_text}",                                                               
        f"Model.ini = {model_ini}",                                               
    ]

    res = {
        "result": result,
        "model_ini": model_ini,
        "conditions": conditions,
    }

    # 報表輸出
    if args.report or args.report_xlsx:
        xlsx_path = args.report_xlsx if args.report_xlsx else "kipling.xlsx"
        export_report(res, xlsx_path=xlsx_path, num_condition_cols=args.conditions)
        sheet = _sheet_name_for_model(model_ini)
        print(f"[INFO] Report appended to: {xlsx_path} (sheet: {sheet})")


if __name__ == "__main__":
    main()
