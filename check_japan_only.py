import os
import re
import argparse

def _resolve_tvconfigs_path(raw: str, root_dir: str) -> str:
    """
    將 model.ini 內的 /tvconfigs/... 轉為實際檔案路徑：${abs_root}/...
    規則：移除前綴 '/tvconfigs/'，其餘部分接在 abs(root_dir) 後面。
    """
    root_dir = os.path.abspath(root_dir)
    rel_path = raw.lstrip("/").replace("tvconfigs/", "", 1)
    return os.path.join(root_dir, rel_path)

def _extract_quoted_value(line: str, key: str) -> str:
    """
    從行文字中抽取 key = "..." 的雙引號內容；忽略後面的 ; 和 # 註解。
    失敗回傳空字串。
    """
    m = re.search(rf'{re.escape(key)}\s*=\s*"([^"]+)"', line)
    return m.group(1).strip() if m else ""

def check_japan_only(model_ini_path: str, root_dir: str = ".") -> str:
    """
    在 model.ini 內尋找 COUNTRY_PATH（忽略以 # 開頭的註解行）。
    取得雙引號中的路徑，轉為 ${abs_root}/...，讀取內容並計數 'japan'（忽略大小寫）。
    規則：恰好出現一次 -> PASS；其他（0 次或 >=2 次）-> FAIL。
    若未宣告 COUNTRY_PATH -> N/A。
    """
    root_dir = os.path.abspath(root_dir)
    target_line = None

    with open(model_ini_path, "r", encoding="utf-8", errors="ignore") as f:
        for line in f:
            s = line.strip()
            if not s or s.startswith("#"):
                continue
            if "COUNTRY_PATH" in s and "=" in s:
                target_line = s
                break

    if not target_line:
        print(f"{model_ini_path}: COUNTRY_PATH = N/A")
        return "N/A"

    raw_path = _extract_quoted_value(target_line, "COUNTRY_PATH")
    if not raw_path:
        print(f"\n{model_ini_path}:\n→ {target_line}\n→ FAIL（COUNTRY_PATH 格式錯誤或缺少引號）")
        return "FAIL"

    abs_path = _resolve_tvconfigs_path(raw_path, root_dir)
    if not os.path.exists(abs_path):
        print(f"\n{model_ini_path}:\n→ {target_line}\n→ FAIL（檔案不存在: {abs_path}）")
        return "FAIL"

    try:
        with open(abs_path, "r", encoding="utf-8", errors="ignore") as cf:
            content = cf.read()
    except Exception as e:
        print(f"\n{model_ini_path}:\n→ {target_line}\n→ FAIL（讀檔錯誤: {e}）")
        return "FAIL"

    count = len(re.findall(r"japan", content, flags=re.IGNORECASE))
    if count == 1:
        print(f"\n{model_ini_path}:\n→ {target_line}\n→ PASS（'japan' 僅出現一次）")
        return "PASS"
    else:
        print(f"\n{model_ini_path}:\n→ {target_line}\n→ FAIL（'japan' 出現 {count} 次）")
        return "FAIL"


def main():
    parser = argparse.ArgumentParser(description="Check if COUNTRY_PATH file contains exactly one 'japan' (ignore case).")
    parser.add_argument("--root", required=True, help="專案根目錄路徑")
    parser.add_argument("--model-ini", required=True, help="model.ini 檔案路徑")
    args = parser.parse_args()

    check_japan_only(args.model_ini, args.root)


if __name__ == "__main__":
    main()
