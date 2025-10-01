import argparse
import os
import re
import sys
from pathlib import Path
from typing import Dict, List, Any, Tuple, Optional, Iterable, Union
from importlib import import_module

MODULE1 = {
    "dvb_country":        {"desc": "EU 只有包含 DVB 的國家",             "module": "target_country_check"},
    "cltv":               {"desc": "可enable CLTV,也可不整合",           "module": "check_cltv"},
    "tv_multi_standard":  {"desc": "多制式切換可選擇開或是關",            "module": "tv_multi_standard_validation"},
}
MODULE2 = {
    "atsc_country":       {"desc": "NA 只有包含 ASTC 的國家",           "module": "target_country_check"},
    "cltv":               {"desc": "可enable CLTV,也可不整合",           "module": "check_cltv"},
    "tv_multi_standard":  {"desc": "多制式切換可選擇開或是關",            "module": "tv_multi_standard_validation"},
}
MODULE3 = {
    "isdb_country":       {"desc": "BRA 只有包含 ISDB(不含 JP) 的國家",  "module": "target_country_check"},
    "cltv":               {"desc": "可enable CLTV,也可不整合",           "module": "check_cltv"},
    "tv_multi_standard":  {"desc": "多制式切換可選擇開或是關",            "module": "tv_multi_standard_validation"},
}
MODULE4 = {
    "Japan":               {"desc": "ISDB-Japan only",                   "module": "check_japan_only"},
}
MODULE5 = {
    "cltv":               {"desc": "CLTV 包含全部制式的國家",                "module": "check_cltv_pid5"},
    "pic_mode":           {"desc": "標準測試 image (QC/CCC)",               "module": "pic_mode_test"},
    "setupwizard":        {"desc": "model.ini : isShowSetupwizard = true", "module": "check_setupwizard_flag"},
    "multi_std":          {"desc": "要 enable 多制式切換",                  "module": "tv_multi_standard_validation_pid5"},
    "aipq":               {"desc": "support AIPQ",                          "module": "ai_aipq_check"},
    "dolby_darkdetail":   {"desc": "support Dolby vision / darkdetail",     "module": "check_darkdetail_flag_pid5"},
    "ewbs":               {"desc": "EWBS 驗證",                             "module": "check_EWBS"},
}
MODULE6 = {
    "cltv":               {"desc": "CLTV 包含全部制式的國家",                "module": "check_cltv_pid5"},
    "CCC":                {"desc": "標準測試 image (QC/CCC)",               "module": "pic_mode_test"},
    "isShowSetupwizard":  {"desc": "model.ini : isShowSetupwizard = true", "module": "check_setupwizard_flag"},
    "tv_multi_standard":  {"desc": "要 enable 多制式切換",                  "module": "tv_multi_standard_validation_pid5"},
    "aipq":               {"desc": "support AIPQ",                         "module": "ai_aipq_check"},
    "dolby_darkdetail":   {"desc": "support Dolby vision / darkdetail",    "module": "check_darkdetail_flag_pid5"},
    "EWBS":               {"desc": "EWBS 驗證",                            "module": "check_EWBS"},
    "Dolby audio cert ":  {"desc": "Dolby Audio 認證用",                   "module": "dolby_cert_check"}
}
MODULE7 = {
    "dvbt & ntsc":        {"desc": "dvbt and ntsc是for columbia and tawian", "module": "check_tvconfig_and_mheg5"},
    #"沒有mheg5":          {"desc": "沒有mheg5,",                              "module": "check_tvconfig_and_mheg5"},
    #"model_ini":          {"desc": "tvconfig=tv.config.dvb_ntsc",            "module": "check_tvconfig_and_mheg5"},
}
MODULE8 = {
}
MODULE9 = {
    "dias":               {"desc": "dias project",             "module": "check_dias_project"},
    "low latency":        {"desc": "low latency = true",       "module": "low_latency_ctrl_check"},
}
MODULE10 = {
    "dias":               {"desc": "dias project",             "module": "check_dias_project"},
    "Dolby audio cert ":  {"desc": "Dolby Audio 認證用",        "module": "dolby_cert_check"}
}
MODULE11 = {
    "netflix":            {"desc": "set picture mode",          "module": "check_netflix_cert"},
    "memc":               {"desc": "Memc off",                  "module": "check_ostable_memc"},
}
MODULE12 = {
    "Dolby vison":        {"desc": "Dolby vison 認証",         "module": "dolby_cert_check_pid12"},
    "AIPQ/AI":            {"desc": "AIPQ/AI default off",      "module": "ai_aipq_check_pid12"},
    "Dolyb Dark":         {"desc": "Dolyb Dark UI 要開",       "module": "check_darkdetail_flag_pid12"},
    "ALLM":               {"desc": "ALLM 要打開",              "module": "check_allm_enable_pid12"},
    "new PQ Menu":        {"desc": "new PQ Menu 架構",         "module": "check_pq_assets"},
    "low_latency":        {"desc": "low latency 預設 off",     "module": "low_latency_ctrl_check"},
    "GDBS":               {"desc": "GDBS 設定",                "module": "check_gdbs_mode"},
    "color space":        {"desc": "ColorSpace off",           "module": "check_osdtable_colorspace"}
}
MODULE13 = {
    "dias":               {"desc": "dias project",             "module": "check_dias_project"},
    "low latency":        {"desc": "low latency = true",       "module": "low_latency_ctrl_check"},
}
MODULE14 = {
    "dias":               {"desc": "dias project",             "module": "check_dias_5k"},
    "low latency":        {"desc": "low latency = true",       "module": "low_latency_ctrl_check"},
}
MODULE15 = {
    "netflix":            {"desc": "set picture mode",          "module": "check_netflix_cert"},
    "memc":               {"desc": "Memc off",                  "module": "check_ostable_memc"},
}
MODULE16 = {
    "Dolby vison":        {"desc": "Dolby vison 認証",         "module": "dolby_cert_check_pid12"},
    "AIPQ/AI":            {"desc": "AIPQ/AI default off",      "module": "ai_aipq_check_pid12"},
    "Dolyb Dark":         {"desc": "Dolyb Dark UI 要開",       "module": "check_darkdetail_flag_pid12"},
    "ALLM":               {"desc": "ALLM 要打開",              "module": "check_allm_enable_pid12"},
    "new PQ Menu":        {"desc": "new PQ Menu 架構",         "module": "check_pq_assets"},
    "low_latency":        {"desc": "low latency 預設 off",     "module": "low_latency_ctrl_check"},
    "GDBS":               {"desc": "GDBS 設定",                "module": "check_gdbs_mode"},
    "color space":        {"desc": "ColorSpace off",           "module": "check_osdtable_colorspace"},
    "dias":               {"desc": "dias project",             "module": "check_dias_5k"},
}
MODULE17 = {
}
MODULE18 = {
}
MODULE19 = {
}
MODULE20 = {
    "PVR":         {"desc": "disable PVR",              "module": "check_show_pvr_flag"},
    "dvb-s":       {"desc": "Disable DVB C/S/S2",       "module": "check_dvbs_satellite_flag"},
    "cltv":        {"desc": "disable CLTV",             "module": "check_cltv_pid20"},
}
MODULES = {
     "1": MODULE1,
     "2": MODULE2,
     "3": MODULE3,
     "4": MODULE4,
     "5": MODULE5,
     "6": MODULE6,
     "7": MODULE7,
     "8": MODULE8,
     "9": MODULE9,
    "10": MODULE10,
    "11": MODULE11,
    "12": MODULE12,
    "13": MODULE13,
    "14": MODULE14,
    "15": MODULE15,
    "16": MODULE16,
    "17": MODULE17,
    "18": MODULE18,
    "19": MODULE19,
    "20": MODULE20
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

def _call_check_module(mod_name: str,
                       model_ini: str,
                       root: str,
                       standard: str,
                       verbose: bool) -> List[List[Any]]:

    mod = import_module(mod_name)
    # 模組自帶 run()
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
        except TypeError:
            # 參數不合，就用最小集
            rows = runfn(model_ini=model_ini, root=root)
            return []

    raise RuntimeError(f"模組 {mod_name} 未提供 run()")

def run_checks_into_book(model_ini: str,
                         root: str,
                         standard: str,
                         verbose: bool,
                         prefix: int ) -> None:
    if prefix not in MODULES:
        print(f"[ERR] 無效的 prefix: {prefix} ({os.path.basename(model_ini)})", file=sys.stderr)
        return

    modules_dict = MODULES[prefix]
    for key in modules_dict:
        modname = modules_dict[key]["module"]
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

    if args.prefix == 1:
        args.standard = "DVB"
    elif args.prefix == 2:
        args.standard = "ATSC"
    elif args.prefix == 3:
        args.standard = "ISDB"

    # 掃描模式
    if args.scan:
        mapping = iter_device_model_inis(Path(root), prefix=args.prefix)
        if not mapping:
            print("[INFO] 掃描不到符合的 model.ini", file=sys.stderr)
            sys.exit(1)

        for dev_dir, ini_list in mapping.items():
            device_name = dev_dir.name
            out_xlsx = dev_dir / f"{device_name}.xlsx"
            for ini in ini_list:
                run_checks_into_book(str(ini), root, args.standard, args.verbose, args.prefix)
        print("\n[SUMMARY] 完成所有 device 報告輸出。")
        sys.exit(0)

if __name__ == "__main__":
    main()
