import os
import re
import argparse

def check_launch_cltv_by_country(model_ini_path: str, root_dir: str = ".") -> str:
    """
    檢查 model.ini 內是否有 LaunchCLTVByCountry，並驗證檔案存在性。
    
    Args:
        model_ini_path (str): model.ini 的檔案路徑
        root_dir (str): 專案根目錄 (用來拼接相對路徑)
    
    Returns:
        str: 結果字串 ("PASS", "FAIL", "N/A")
    """
    root_dir = os.path.abspath(root_dir)  # 確保是絕對路徑
    launch_line = None
    with open(model_ini_path, "r", encoding="utf-8", errors="ignore") as f:
        for line in f:
            # 跳過註解行
            if line.strip().startswith("#"):
                continue
            if "LaunchCLTVByCountry" in line and "=" in line:
                launch_line = line.strip()
                break

    if not launch_line:
        print(f"{model_ini_path}:\n→ LaunchCLTVByCountry = N/A")
        return "N/A"

    # 用 regex 抓取雙引號內的字串
    match = re.search(r'LaunchCLTVByCountry\s*=\s*"([^"]+)"', launch_line)
    if not match:
        print(f"{model_ini_path}: LaunchCLTVByCountry 格式錯誤 → FAIL")
        return "FAIL"

    value = match.group(1).strip()

    if not value:
        print(f"{model_ini_path}: LaunchCLTVByCountry 路徑未設定 → FAIL")
        return "FAIL"

    # 移除 "/tvconfigs/" 前綴
    rel_path = value.lstrip("/").replace("tvconfigs/", "", 1)
    abs_path = os.path.join(root_dir, rel_path)

    if os.path.exists(abs_path):
        print(f"\n{model_ini_path}:\n→ {launch_line} \n→ PASS")
        return "PASS"
    else:
        print(f"\n{model_ini_path}:\n→ {launch_line} \n→ FAIL (檔案不存在)")
        return "FAIL"


def main():
    parser = argparse.ArgumentParser(description="Check LaunchCLTVByCountry in model.ini")
    parser.add_argument("--root", required=True, help="專案根目錄路徑")
    parser.add_argument("--model-ini", required=True, help="model.ini 檔案路徑")
    args = parser.parse_args()

    check_launch_cltv_by_country(args.model_ini, args.root)


if __name__ == "__main__":
    main()

