import argparse
import os
import re
import sys
from pathlib import Path
from typing import Dict, List, Any, Tuple, Optional, Iterable, Union
from importlib import import_module

from reporting import ReportBook  # 與本檔同資料夾

# === 你要串的檢查模組（檔名 = module 名） ===
MODULES = {
    "cltv":               {"desc": "CLTV 包含全部制式的國家",             "module": "check_cltv_pid5"},
    "pic_mode":           {"desc": "標準測試 image (QC/CCC)",              "module": "pic_mode_test"},
    "setupwizard":        {"desc": "model.ini : isShowSetupwizard = true", "module": "check_setupwizard_flag"},
    "multi_std":          {"desc": "要 enable 多制式切換",                 "module": "tv_multi_standard_validation_pid5"},
    "aipq":               {"desc": "support AIPQ",                          "module": "ai_aipq_check"},
    "dolby_darkdetail":   {"desc": "support Dolby vision / darkdetail",     "module": "check_darkdetail_flag_pid5"},
    "ewbs":               {"desc": "EWBS 驗證",                             "module": "check_EWBS"},
}

# =========================
# 工具：PID 分頁、device_name
# =========================

_PID_PREFIX = re.compile(r"^(\d+)_")
_TV_DIR_RE  = re.compile(r"^tv\d{3}$", re.IGNORECASE)

def guess_pid_sheet_name(model_ini: str) -> str:
    base = os.path.basename(model_ini)
    m = _PID_PREFIX.match(base)
    return f"PID_{m.group(1)}" if m else "PID"

def get_device_name_from_path(path: Union[str, os.PathLike, Path]) -> Optional[str]:
    """
    接受 str / Path / PathLike。
    若傳入的是檔案（例如 *.ini），會先取 parent 再往上找 tvxxx/<device_name>/...
    找不到則回傳 None。
    """
    p = path if isinstance(path, Path) else Path(path)
    p = p.resolve()

    # 若傳進來的是檔案路徑（含 .ini 字串），往上一層目錄再判斷
    if p.is_file() or p.suffix.lower() == ".ini":
        p = p.parent

    _TV_RE = re.compile(r"^tv\d{3}$", re.IGNORECASE)
    parts = p.parts
    for i in range(len(parts) - 1):
        if _TV_RE.match(parts[i]):
            return parts[i + 1] if i + 1 < len(parts) else None
    return None

# =========================
# 動態呼叫：模組 → 列資料
# =========================

def _normalize_rows(rows: Any) -> List[List[Any]]:
    """容錯：支援 run() 回傳 list[list] / list[dict] / dict"""
    if rows is None:
        return []
    if isinstance(rows, list):
        if not rows:
            return []
        if isinstance(rows[0], dict):
            out: List[List[Any]] = []
            for d in rows:
                rule = d.get("rule") or d.get("Rules") or ""
                result = d.get("result") or d.get("Result") or ""
                conds = d.get("conditions") or []
                if isinstance(conds, dict):
                    conds = [f"{k}={v}" for k, v in conds.items()]
                out.append([rule, result] + list(conds))
            return out
        return rows
    if isinstance(rows, dict):
        if "rows" in rows:
            return _normalize_rows(rows["rows"])
        rule = rows.get("rule") or rows.get("Rules") or ""
        result = rows.get("result") or rows.get("Result") or ""
        conds = rows.get("conditions") or []
        if isinstance(conds, dict):
            conds = [f"{k}={v}" for k, v in conds.items()]
        return [[rule, result] + list(conds)]
    return []

def _call_check_module(mod_name: str,
                       model_ini: str,
                       root: str,
                       standard: str,
                       verbose: bool) -> List[List[Any]]:

    mod = import_module(mod_name)
    # 1) 首選：模組自帶 run()
    if hasattr(mod, "run"):
        runfn = getattr(mod, "run")
        try:
            # 盡量傳齊一點參數；不同模組可忽略不用的
            rows = runfn(
                model_ini=model_ini,
                root=root,
                standard=standard,
                verbose=verbose,
                conditions="",
                report_xlsx=get_device_name_from_path(model_ini),  # report_xlsx
                ctx=None,
            )
            return []
            #return _normalize_rows(rows)
        except TypeError:
            # 參數不合，就用最小集
            rows = runfn(model_ini=model_ini, root=root)
            #return _normalize_rows(rows)
            return []

    raise RuntimeError(f"模組 {mod_name} 未提供 run()，也沒有已知的 fallback 呼叫法")

def run_checks_into_book(model_ini: str,
                         root: str,
                         standard: str,
                         verbose: bool) -> None:
                         #book: ReportBook) -> None:
    #sheet = guess_pid_sheet_name(model_ini)
    for key in MODULES:
        modname = MODULES[key]["module"]
        try:
            rows = _call_check_module(modname, model_ini, root, standard, verbose)
            print(f"[OK] {key:>18}  rows={len(rows)}  ({os.path.basename(model_ini)})")
        except Exception as e:
            print(f"[ERR] {key:>18}  {e}  ({os.path.basename(model_ini)})", file=sys.stderr)

# =========================
# 掃描：找 5_* model ini
# =========================

def iter_device_model_inis(root: Path, prefix: str = "5") -> Dict[Path, List[Path]]:
    """
    回傳 { device_dir(tvxxx/<device_name>): [5_*.ini 路徑, ...], ... }
    只掃到 tvxxx/<device_name>/configs/model/*.ini
    """
    result: Dict[Path, List[Path]] = {}
    for tvdir in root.glob("tv[0-9][0-9][0-9]"):
        if not tvdir.is_dir():
            continue
        for dev in tvdir.iterdir():
            model_dir = dev / "configs" / "model"
            if not model_dir.is_dir():
                continue
            files = sorted(p for p in model_dir.glob(f"{prefix}_*.ini") if p.is_file())
            if files:
                result[dev] = files
    #print("Result:", result)
    return result

# =========================
# CLI
# =========================

def parse_args(argv: Optional[Iterable[str]] = None) -> argparse.Namespace:
    p = argparse.ArgumentParser(description=__doc__)
    # 共通
    p.add_argument("--root", default=".", help="專案根目錄（對應 /tvconfigs 映射）")
    p.add_argument("--standard", default="DVB", help="制式（部分檢查需要）")
    p.add_argument("-v", "--verbose", action="store_true", help="顯示更詳細輸出")
    # 掃描模式
    p.add_argument("--scan", action="store_true", help="啟用掃描模式：tvxxx/<device_name>/configs/model/ 下的 5_*.ini")
    p.add_argument("--prefix", default="5", help="掃描時的檔名前綴（預設 5）")
    return p.parse_args(argv)

def main() -> None:
    args = parse_args()
    root = str(Path(args.root).resolve())

    # 掃描模式
    if args.scan:
        mapping = iter_device_model_inis(Path(root), prefix=args.prefix)
        if not mapping:
            print("[INFO] 掃描不到符合的 model.ini", file=sys.stderr)
            sys.exit(1)

        for dev_dir, ini_list in mapping.items():
            device_name = dev_dir.name
            out_xlsx = dev_dir / f"{device_name}.xlsx"
            #print(f"\n[DEVICE] {dev_dir}  →  {out_xlsx.name}")
            #book = ReportBook(str(out_xlsx))
            for ini in ini_list:
                run_checks_into_book(str(ini), root, args.standard, args.verbose)
            #book.save()
        print("\n[SUMMARY] 完成所有 device 報告輸出。")
        sys.exit(0)

if __name__ == "__main__":
    main()

