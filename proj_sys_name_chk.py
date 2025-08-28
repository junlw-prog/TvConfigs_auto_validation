# -*- coding: utf-8 -*-
import os
import re
import sys
from glob import glob
from typing import Optional

# ========== 功能 1 ==========
# 修改 model/ 1_~20_ 檔內「PROJECT_NAME =」行為「PROJECT_NAME = <檔名(不含.ini)>;」
# 注意：這個 function 不做排除清單（需求指定）
def update_project_name_in_model(root_folder, dry_run=False):
    model_dir = os.path.join(root_folder, "model")
    prefixes = [f"{i}_" for i in range(1, 21)]

    for root, dirs, files in os.walk(model_dir):
        for filename in files:
            if any(filename.startswith(prefix) for prefix in prefixes) and filename.endswith(".ini"):
                filepath = os.path.join(root, filename)
                base_name = os.path.splitext(filename)[0]

                try:
                    lines_out = []
                    modified = False
                    with open(filepath, "r", encoding="utf-8", errors="ignore") as f:
                        for lineno, line in enumerate(f, start=1):
                            if line.strip().startswith("PROJECT_NAME ="):
                                new_line = f"PROJECT_NAME = {base_name};\n"
                                if line != new_line:
                                    print(f"{filepath}:{lineno}: 修改 {line.strip()} → {new_line.strip()}")
                                    line = new_line
                                    modified = True
                            lines_out.append(line)

                    if modified:
                        if dry_run:
                            print(f"[Dry-run] 將修改 {filepath}")
                        else:
                            with open(filepath, "w", encoding="utf-8") as f:
                                f.writelines(lines_out)
                            print(f"已修改 {filepath}")

                except Exception as e:
                    print(f"處理 {filepath} 失敗: {e}")


# ========== 功能 2 ==========
# 更新 sys/ 下 .ini 裡的 Model_1 ~ Model_20 路徑
# 等號右側覆蓋為 "/tvconfigs/model/<對應檔名>"，但「找不到對應檔案就不修改」（不新增、不覆蓋）
def _pick_model_filename(model_dir: str, idx: int) -> Optional[str]:
    pattern = os.path.join(model_dir, f"{idx}_*.ini")
    matches = sorted(glob(pattern))
    if matches:
        return os.path.basename(matches[0])
    return None

def update_sys_models(root_folder: str, sys_dir: str = "sys", model_dir: str = "model", dry_run: bool = False):
    sys_path = os.path.join(root_folder, sys_dir)
    model_path = os.path.join(root_folder, model_dir)

    index_to_filename = {i: _pick_model_filename(model_path, i) for i in range(1, 21)}

    for root, _, files in os.walk(sys_path):
        for fn in files:
            if not fn.lower().endswith(".ini"):
                continue
            fullpath = os.path.join(root, fn)

            try:
                with open(fullpath, "r", encoding="utf-8", errors="ignore") as f:
                    original_lines = f.readlines()

                out_lines = []
                model_line_re = re.compile(r'^\s*Model_(\d+)\s*=')
                changed = False

                for line in original_lines:
                    m = model_line_re.match(line)
                    if m:
                        idx = int(m.group(1))
                        # 僅在 1..20 且找得到對應檔案時才覆蓋
                        if 1 <= idx <= 20 and index_to_filename[idx]:
                            filename = index_to_filename[idx]
                            new_line = f'Model_{idx} = "/tvconfigs/model/{filename}"\n'
                            if line != new_line:
                                print(f'{fullpath}: 覆寫 Model_{idx} → {new_line.strip()}')
                                line = new_line
                                changed = True
                    out_lines.append(line)

                if changed:
                    if dry_run:
                        print(f"[Dry-run] 將更新：{fullpath}")
                    else:
                        with open(fullpath, "w", encoding="utf-8") as f:
                            f.writelines(out_lines)
                        print(f"已更新：{fullpath}")

            except Exception as e:
                print(f"處理 {fullpath} 失敗：{e}")

if __name__ == "__main__":
    # 根目錄（底下應有 model/、tvserv_ini/、sys/）
    folder_path = "./"
    dry_run = "--dry-run" in sys.argv

    # 依需求順序執行
    update_project_name_in_model(folder_path, dry_run=dry_run)
    update_sys_models(folder_path, "sys", "model", dry_run=dry_run)


