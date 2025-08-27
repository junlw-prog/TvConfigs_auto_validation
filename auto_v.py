#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
tvconfigs_path_check.py
掃描 configs/ 下的 .ini，檢查 ini 內指向的 /tvconfigs/... 路徑是否存在於專案根
"""

import argparse
import sys
from pathlib import Path
import re
from typing import Dict, Iterable, List, NamedTuple, Optional, Set

# 允許的副檔名（不含點）
DEFAULT_EXTS = {"ini", "bin", "dat", "img", "xml", "json", "png", "jpg", "webp"}

# 抓出 ini 行內所有 /tvconfigs/... 片段（直到空白/引號/; 結束）
TV_PATH_RE = re.compile(r'(/tvconfigs/[^\s";]+)')

# 粗略抓 key=val，僅為了報告顯示，避免誤解析不影響檢查
KEYVAL_RE = re.compile(r'^\s*([^;#=\s][^=]{0,80}?)\s*=\s*(.*)$')

class Ref(NamedTuple):
    ini_file: Path
    line_no: int
    line_preview: str
    raw_tv_path: str
    resolved_path: Path

def iter_ini_files(root: Path) -> Iterable[Path]:
    """遞迴尋找所有 .ini（含 sys/model/tvserv_ini/...）"""
    for p in root.rglob("*.ini"):
        # 跳過隱藏或非一般檔案
        if p.is_file():
            yield p

def sanitize_tv_path(p: str) -> str:
    """去掉可能的收尾符號（逗號、右括號、尾巴的引號/分號等）"""
    return p.rstrip('",; \t\r\n)')

def looks_like_file_of_interest(tv_path: str, allowed_exts: Set[str]) -> bool:
    """只檢查我們關心的副檔名（或沒有副檔名但仍想檢）"""
    # 把查詢字串（例如 ?v=）等簡單剔除
    base = tv_path.split("?")[0]
    # 取副檔名（不含點），大小寫忽略
    ext = Path(base).suffix.lower().lstrip(".")
    if not ext:
        # 沒有副檔名：通常是目錄或特殊檔案，預設不檢；如要檢，可在此返回 True
        return False
    return ext in allowed_exts

def resolve_to_project(root: Path, tv_path: str, prefix_map: Optional[Dict[str, Path]] = None) -> Optional[Path]:
    """
    把 /tvconfigs/... 這類邏輯路徑，映射成專案中的實體檔案路徑。
    預設：/tvconfigs/xxx  ->  root/xxx
    你也可在 prefix_map 補更多映射（例如不同專案自訂前綴）。
    """
    tv_path = sanitize_tv_path(tv_path)
    prefix_map = prefix_map or {"/tvconfigs/": root}

    for prefix, base in prefix_map.items():
        if tv_path.startswith(prefix):
            rel = tv_path[len(prefix):]  # 去掉前綴
            return (base / rel).resolve()
    return None  # 非我們能解析的前綴（目前僅處理 /tvconfigs/）

def scan_ini_for_tv_paths(ini_file: Path, root: Path, allowed_exts: Set[str], prefix_map: Optional[Dict[str, Path]]) -> List[Ref]:
    """逐行掃描 ini，找出 /tvconfigs/... 參照並回傳 Ref 清單"""
    refs: List[Ref] = []
    with ini_file.open("r", encoding="utf-8", errors="ignore") as f:
        for i, raw_line in enumerate(f, start=1):
            line = raw_line.strip()
            if not line or line.startswith(("#", ";")):
                continue  # 跳過註解/空行
            matches = TV_PATH_RE.findall(line)
            if not matches:
                continue
            # 擷取 key=value 片段做報告 preview（避免整行過長）
            m = KEYVAL_RE.match(line)
            preview = (m.group(0) if m else line)[:200]

            for tv_path in matches:
                if not looks_like_file_of_interest(tv_path, allowed_exts):
                    continue
                resolved = resolve_to_project(root, tv_path, prefix_map)
                if resolved is None:
                    continue
                refs.append(Ref(ini_file=ini_file, line_no=i, line_preview=preview, raw_tv_path=tv_path, resolved_path=resolved))
    return refs

def main():
    parser = argparse.ArgumentParser(description="檢查 ini 內指向的 /tvconfigs/... 檔案是否存在")
    parser.add_argument("--root", required=True, type=Path,
                        help="tvconfigs 專案根（例如 ~/tvconfigs_home/tv109/kipling/configs）")
    parser.add_argument("--exts", type=str, default=",".join(sorted(DEFAULT_EXTS)),
                        help=f"要檢查的副檔名（逗號分隔），預設：{','.join(sorted(DEFAULT_EXTS))}")
    parser.add_argument("--fail-warning", action="store_true",
                        help="把所有缺檔當 Blocking（退出碼非 0）")
    args = parser.parse_args()

    root: Path = args.root.resolve()
    if not root.exists():
        print(f"[ERROR] root 不存在：{root}", file=sys.stderr)
        sys.exit(2)

    allowed_exts = {e.strip().lower() for e in args.exts.split(",") if e.strip()}
    prefix_map = {"/tvconfigs/": root}  # 如需擴充其它前綴，可在此加入

    all_refs: List[Ref] = []
    for ini in iter_ini_files(root):
        all_refs.extend(scan_ini_for_tv_paths(ini, root, allowed_exts, prefix_map))

    missing: List[Ref] = [r for r in all_refs if not r.resolved_path.exists()]

    # 報告
    print(f"掃描根目錄：{root}")
    print(f"發現參照總數：{len(all_refs)}，缺檔：{len(missing)}")
    if missing:
        print("\n=== 缺檔清單（Blocking 建議修正） ===")
        for r in missing:
            print(f"- {r.ini_file.relative_to(root)}:{r.line_no}")
            print(f"  key/line : {r.line_preview}")
            print(f"  tv_path  : {r.raw_tv_path}")
            print(f"  resolved : {r.resolved_path}")
    else:
        print("✔ 未發現缺檔。")

    # 退出碼：有缺檔→非 0，方便 CI 擋下
    if missing:
        sys.exit(1 if args.fail_warning else 1)
    sys.exit(0)

if __name__ == "__main__":
    main()

