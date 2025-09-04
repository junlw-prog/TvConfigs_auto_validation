# Create check_EWBS.py with the requested logic and formatting.
import os
import re
import argparse

COUNTRIES_TO_PRINT = [
    "Botswana", "Ecuador", "Maldives", "Philippines",
    "Japan", "Peru", "Venezuela", "Costa Rica",
]

def _resolve_tvconfigs_path(raw: str, root_dir: str) -> str:
    """
    將 /tvconfigs/... 轉為實際檔案路徑：${abs_root}/...
    規則：移除 'tvconfigs/' 前綴，接在 abs(root_dir) 後面。
    """
    root_dir = os.path.abspath(root_dir)
    rel_path = raw.lstrip("/").replace("tvconfigs/", "", 1)
    return os.path.join(root_dir, rel_path)

def _extract_quoted_value(line: str, key: str) -> str:
    """從 'key = "..."' 抽出雙引號字串，失敗回傳空字串。"""
    m = re.search(rf'{re.escape(key)}\s*=\s*"([^"]+)"', line)
    return m.group(1).strip() if m else ""

def _parse_bool_from_line(line: str, key: str):
    """
    從 'key = true/false' 解析布林值（忽略大小寫與空白），
    也接受 key = "true"/"false"。找不到回傳 None。
    """
    m = re.search(rf'{re.escape(key)}\s*=\s*("?)(true|false)\1', line, flags=re.IGNORECASE)
    if not m:
        return None
    return m.group(2).lower() == "true"

def _iter_non_comment_lines(path: str):
    with open(path, "r", encoding="utf-8", errors="ignore") as f:
        for raw in f:
            s = raw.strip()
            if not s or s.startswith("#"):
                continue
            yield s

def check_ewbs(model_ini_path: str, root_dir: str = ".") -> str:
    """
    規則：
      1) 在 model.ini 找到（且非註解行）
         - isSupportEWBS = true
         - isSupportNeverEnterSTR
         - isEwbsSettingOn
         符合條件 PASS，否則 FAIL。
      2) 再找 COUNTRY_PATH 指向的檔案，讀取內容；若包含 COUNTRIES_TO_PRINT 任一國家（忽略大小寫）則列印出來。
    """
    root_dir = os.path.abspath(root_dir)

    # --- 第一階段：讀取三個旗標 ---
    val_ewbs = val_never_str = val_setting_on = None
    ewbs_line = never_str_line = setting_on_line = None
    country_line = None

    for s in _iter_non_comment_lines(model_ini_path):
        if ewbs_line is None and "isSupportEWBS" in s and "=" in s:
            ewbs_line = s
            val_ewbs = _parse_bool_from_line(s, "isSupportEWBS")

        if never_str_line is None and "isSupportNeverEnterSTR" in s and "=" in s:
            never_str_line = s
            val_never_str = _parse_bool_from_line(s, "isSupportNeverEnterSTR")

        if setting_on_line is None and "isEwbsSettingOn" in s and "=" in s:
            setting_on_line = s
            val_setting_on = _parse_bool_from_line(s, "isEwbsSettingOn")

        if country_line is None and "COUNTRY_PATH" in s and "=" in s:
            country_line = s

        # 小優化：若四者皆已找到可提前結束
        if (ewbs_line is not None and never_str_line is not None and
            setting_on_line is not None and country_line is not None):
            break

    # 組合輸出
    print(f"\n{model_ini_path}:")
    if ewbs_line:        print(f"→ {ewbs_line}")
    if never_str_line:   print(f"→ {never_str_line}")
    if setting_on_line:  print(f"→ {setting_on_line}")

    flags_ok = (val_ewbs is True) and (never_str_line) and (setting_on_line)
    if flags_ok:
        print("→ PASS（isSupportEWBS = true/isSupportNeverEnterSTR is set/isEwbsSettingOn is set）")
        result = "PASS"
    else:
        reasons = []
        if val_ewbs is not True:        reasons.append("isSupportEWBS != true")
        if never_str_line is None:   reasons.append("isSupportNeverEnterSTR not set")
        if val_setting_on is None:  reasons.append("isEwbsSettingOn not set")
        if not reasons:
            reasons.append("缺少必要欄位（可能未宣告或格式錯誤）")
        print(f"→ FAIL（{', '.join(reasons)}）")
        result = "FAIL"

    # --- 第二階段：開 COUNTRY_PATH，列印指定國家（若存在） ---
    if country_line:
        raw_country_path = _extract_quoted_value(country_line, "COUNTRY_PATH")
        if raw_country_path:
            abs_country_path = _resolve_tvconfigs_path(raw_country_path, root_dir)
            if os.path.exists(abs_country_path):
                try:
                    with open(abs_country_path, "r", encoding="utf-8", errors="ignore") as cf:
                        content = cf.read()
                    found = []
                    for name in COUNTRIES_TO_PRINT:
                        if re.search(re.escape(name), content, flags=re.IGNORECASE):
                            found.append(name)
                    if found:
                        print(f"→ COUNTRY_PATH 內包含國家：{', '.join(found)}")
                    else:
                        print("→ COUNTRY_PATH 內未找到指定國家")
                except Exception as e:
                    print(f"→ 讀取 COUNTRY_PATH 檔案失敗：{e}")
            else:
                print(f"→ COUNTRY_PATH 檔案不存在：{abs_country_path}")
        else:
            print("→ COUNTRY_PATH 格式錯誤或缺少引號")
    else:
        print("→ COUNTRY_PATH = N/A")

    return result

def main():
    parser = argparse.ArgumentParser(description="Check EWBS flags and list specific countries from COUNTRY_PATH.")
    parser.add_argument("--root", required=True, help="專案根目錄路徑")
    parser.add_argument("--model-ini", required=True, help="model.ini 檔案路徑")
    args = parser.parse_args()

    check_ewbs(args.model_ini, args.root)

if __name__ == "__main__":
    main()

