#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from pathlib import Path
import json
import sys
from typing import Dict, List, Any

# === 你提供的預設規則（可被 sys/required_rules.json 覆蓋） ===
DEFAULT_RULES: Dict[str, List[str]] = {
    "must_dirs": [
        "AQ","audio","edid","freqtables","icon","IR","key","language",
        "lcn_config","oem","panel","pcb","PQ","PQ_OSD","prechlist","refChipCfg","sys",
        "tv_config","tvdata","tv_misc","tvserv_ini","TvSysMap","menu","model",
        "operator","overscan","logo"
    ],
    "should_dirs": ["board","device","country","led"],
    "optional_dirs": ["camera","karaoke","sat_tab","sdx","ssc"],
    "must_files": ["customer.ini"],
    "ignore_files": ["Readme.txt"],
    "ignore_dirs": [".git", ".repo"]
}

def load_required_rules(root: Path) -> Dict[str, Any]:
    """
    若存在 sys/required_rules.json 則用該設定覆蓋預設。
    （已移除 B 方案：不再將 should_dirs 併入 must_dirs）
    """
    rules = {k: (v[:] if isinstance(v, list) else v) for k, v in DEFAULT_RULES.items()}
    cfg = root / "sys" / "required_rules.json"
    if cfg.exists():
        try:
            override = json.loads(cfg.read_text(encoding="utf-8"))
            for k, v in override.items():
                rules[k] = v
        except Exception as e:
            print(f"[required_rules] 無法解析 {cfg}: {e}，改用預設規則")
    return rules

def check_required_structure(root_dir: str) -> Dict[str, Any]:
    """
    在做「引用檔案存在性檢查」之前，先檢查專案骨架是否完整。
    回傳 dict（不包含任何 ignore 項目的列舉），由呼叫端決定是否中止流程或繼續。
    """
    root = Path(root_dir).resolve()
    if not root.exists():
        raise FileNotFoundError(f"根目錄不存在：{root}")

    rules = load_required_rules(root)

    # --- 1) 檢查必須目錄 ---
    missing_must_dirs = [d for d in rules["must_dirs"] if not (root / d).is_dir()]

    # --- 2) 檢查建議目錄（提醒用，不致命） ---
    missing_should_dirs = [d for d in rules.get("should_dirs", []) if not (root / d).is_dir()]

    # --- 3) 檢查必須檔案 ---
    missing_must_files = [f for f in rules.get("must_files", []) if not (root / f).exists()]

    # --- 4) 列出未知目錄/檔案（排除 ignore，僅供參考） ---
    ignore_dirs = set(rules.get("ignore_dirs", []))
    known_dirs = set(rules.get("must_dirs", [])) \
                 | set(rules.get("should_dirs", [])) \
                 | set(rules.get("optional_dirs", [])) \
                 | ignore_dirs

    present_dirs = [p.name for p in root.iterdir() if p.is_dir()]
    unknown_dirs = sorted([d for d in present_dirs
                           if d not in known_dirs and not d.startswith(".")])

    present_files = [p.name for p in root.iterdir() if p.is_file()]
    ignore_files = set(rules.get("ignore_files", []))
    must_files = set(rules.get("must_files", []))
    # 既不是忽略、也不是必須的檔案 -> 僅列出供參考
    extraneous_files = sorted([f for f in present_files if f not in (ignore_files | must_files)])

    # --- 5) 統一輸出結果（不含任何 ignore 項目的列舉/提示） ---
    summary: Dict[str, Any] = {
        "root": str(root),
        "missing_must_dirs": missing_must_dirs,
        "missing_should_dirs": missing_should_dirs,
        "missing_must_files": missing_must_files,
        "unknown_dirs": unknown_dirs,
        "extraneous_files": extraneous_files,
        # 回傳規則給上游參考，但不含 ignore_* 設定
        "rules_used": {
            "must_dirs": rules.get("must_dirs", []),
            "should_dirs": rules.get("should_dirs", []),
            "optional_dirs": rules.get("optional_dirs", []),
            "must_files": rules.get("must_files", [])
        }
    }
    pretty_print_summary(summary)
    return summary

def pretty_print_summary(summary: Dict[str, Any]) -> None:
    root = summary.get("root", "")
    print(f"== 檢查必要目錄/檔案：{root} ==")

    def _p(title: str, items: List[str]) -> None:
        if items:
            print(f"[{title}] ({len(items)})")
            for it in items:
                print(f"  - {it}")

    _p("缺少【必須目錄】", summary.get("missing_must_dirs", []))
    _p("缺少【建議目錄】(提醒)", summary.get("missing_should_dirs", []))
    _p("缺少【必須檔案】", summary.get("missing_must_files", []))
    _p("未知目錄（非規則內，僅列出供參考）", summary.get("unknown_dirs", []))

    extraneous = summary.get("extraneous_files", [])
    if extraneous:
        print("[非忽略且非必要的檔案]（僅列出供參考）")
        for f in extraneous:
            print(f"  - {f}")

    if not summary.get("missing_must_dirs") and not summary.get("missing_must_files"):
        print("\n✅ 必須目錄/檔案齊備，可進入後續檢查。")
    else:
        print("\n❌ 必須項缺失，建議先補齊後再執行後續『引用檔案存在性』檢查。")

def main(argv: List[str]) -> int:
    if len(argv) < 2:
        print("用法：python3 required_structure_check.py <configs_root_path>")
        print("範例：python3 required_structure_check.py /home/jun/tvconfigs_home/tv001/common/configs")
        return 2
    root_dir = argv[1]
    try:
        check_required_structure(root_dir)
    except Exception as e:
        print(f"執行失敗：{e}")
        return 1
    return 0

if __name__ == "__main__":
    sys.exit(main(sys.argv))

