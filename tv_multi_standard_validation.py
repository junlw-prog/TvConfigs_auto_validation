#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PID_TV_standard_validation.py

Generic TV standard consistency check by COUNTRY_NAME (not ISO-2).
- Supports DVB / ATSC / ISDB via --standard
- If model.ini lacks tvSysMap, auto-fallback to find TvSysMap/tvSysMapCfgs.xml under --root
- For --standard, COUNTRY_PATH is FORCED to: country/{STANDARD}.ini (relative to --root or /tvconfigs prefix)

Exports:
    run_tv_standard_check(model_ini, root, standard='DVB', verbose=False) -> dict
"""

from __future__ import annotations

import argparse
import re
import sys
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import List, Tuple, Dict, Set
import xlsxwriter

def _read_text(p: Path) -> str:
    try:
        return p.read_text(encoding="utf-8", errors="ignore")
    except Exception:
        try:
            return p.read_text(encoding="latin-1", errors="ignore")
        except Exception:
            return ""

def _norm_country_name(s: str) -> str:
    return (s or "").strip().upper().replace(" ", "_")

def _resolve_from_root(root: Path, given: str) -> Path:
    given = (given or "").strip().strip('"\';) ')
    if not given:
        return root
    if given.startswith("/tvconfigs/"):
        return (root / given[len("/tvconfigs/"):]).resolve()
    if given.startswith("/"):
        return (root / given.lstrip("/")).resolve()
    return (root / given).resolve()

_PATH_VALUE_RE = re.compile(r'=\s*("?)([^";\n]+)\1')

def _extract_ini_path(line: str) -> str | None:
    m = _PATH_VALUE_RE.search(line)
    return m.group(2).strip() if m else None

def _find_tv_sys_map_and_country_path(model_ini: Path) -> Tuple[str|None, str|None]:
    text = _read_text(model_ini)
    tv_sys_map = None
    country_path = None
    sect = None
    for raw in text.splitlines():
        line = raw.strip()
        if not line or line.startswith("#") or line.startswith(";"):
            continue
        if line.startswith("[") and line.endswith("]"):
            sect = line[1:-1].strip().upper()
            continue
        low = line.lower()
        if sect == "COUNTRY":
            if low.startswith("tvsysmap"):
                p = _extract_ini_path(line)
                if p: tv_sys_map = p
            elif low.startswith("country_path"):
                p = _extract_ini_path(line)
                if p: country_path = p
    if tv_sys_map is None:
        for raw in text.splitlines():
            if "tvsysmap" in raw.lower():
                p = _extract_ini_path(raw)
                if p:
                    tv_sys_map = p
                    break
    if country_path is None:
        for raw in text.splitlines():
            if "country_path" in raw.lower():
                p = _extract_ini_path(raw)
                if p:
                    country_path = p
                    break
    return tv_sys_map, country_path

def _auto_locate_tvsysmap(root: Path) -> Path | None:
    """Try common fallbacks to find TvSysMap/tvSysMapCfgs.xml under root."""
    candidates = [
        root / "TvSysMap" / "tvSysMapCfgs.xml",
        root / "tvSysMap" / "tvSysMapCfgs.xml",
        root / "TvSysMapCfgs.xml",
    ]
    for c in candidates:
        if c.exists():
            return c.resolve()
    # last resort: shallow glob
    for p in root.glob("**/tvSysMapCfgs.xml"):
        try:
            # limit depth to avoid scanning giant trees
            rel = p.relative_to(root)
            if len(rel.parts) <= 4:
                return p.resolve()
        except Exception:
            continue
    return None

def _resolve_country_tv_map_xmls(tvsysmap_path: Path, root: Path) -> List[Path]:
    xmls: List[Path] = []
    txt = _read_text(tvsysmap_path)
    if not txt.strip():
        return xmls
    try:
        rt = ET.fromstring(txt)
        for el in rt.iter():
            tag = el.tag.split('}')[-1].lower()
            if tag == "countrytvsysmapxml":
                v = (el.text or "").strip()
                if v:
                    p = _resolve_from_root(root, v)
                    xmls.append(p)
    except Exception:
        for m in re.finditer(r'<\s*CountryTvSysMapXML\s*>(.*?)</\s*CountryTvSysMapXML\s*>', txt, re.IGNORECASE|re.DOTALL):
            v = (m.group(1) or "").strip()
            if v:
                p = _resolve_from_root(root, v)
                xmls.append(p)
    seen = set()
    uniq = []
    for p in xmls:
        sp = str(p)
        if sp not in seen:
            seen.add(sp)
            uniq.append(p)
    return uniq

def _parse_country_tv_pairs(country_tv_xml: Path) -> List[Tuple[str, str, Path]]:
    out: List[Tuple[str, str, Path]] = []
    txt = _read_text(country_tv_xml)
    if not txt.strip():
        return out
    try:
        rt = ET.fromstring(txt)
        for blk in rt.iter():
            if blk.tag.split('}')[-1].upper() == "COUNTRY_TVCONFIG_MAP":
                country = ""
                system = ""
                for ch in blk:
                    tag = ch.tag.split('}')[-1].upper()
                    if tag == "COUNTRY_NAME":
                        country = (ch.text or "").strip()
                    elif tag == "TV_SYSTEM":
                        system = (ch.text or "").strip()
                if country:
                    out.append((_norm_country_name(country), (system or "").strip().upper(), country_tv_xml))
        return out
    except Exception:
        for m in re.finditer(r'<\s*COUNTRY_TVCONFIG_MAP\s*>(.*?)</\s*COUNTRY_TVCONFIG_MAP\s*>', txt, re.IGNORECASE|re.DOTALL):
            blk = m.group(1) or ""
            def _tag(t):
                mm = re.search(rf'<\s*{t}\s*>(.*?)</\s*{t}\s*>', blk, re.IGNORECASE|re.DOTALL)
                return (mm.group(1) or "").strip() if mm else ""
            country = _tag("COUNTRY_NAME")
            system  = _tag("TV_SYSTEM")
            if country:
                out.append((_norm_country_name(country), (system or "").strip().upper(), country_tv_xml))
        return out

def _parse_country_list(country_file: Path) -> Set[str]:
    txt = _read_text(country_file)
    if not txt:
        return set()
    no_comments = []
    for line in txt.splitlines():
        line = re.split(r'[;#]', line, maxsplit=1)[0]
        no_comments.append(line)
    txt2 = "\n".join(no_comments)
    tokens = re.split(r'[,\n;]+', txt2)
    out: Set[str] = set()
    for tok in tokens:
        t = tok.strip()
        if not t:
            continue
        if "=" in t:
            _, val = t.split("=", 1)
            t = val.strip()
        out.add(_norm_country_name(t))
    return out

def run_tv_standard_check(model_ini: str | Path, root: str | Path, standard: str = "DVB", verbose: bool=False) -> Dict:
    """
    standard: 'DVB', 'ATSC', or 'ISDB' (case-insensitive, prefix match on TV_SYSTEM).
    COUNTRY_PATH is forced to country/{STANDARD}.ini under --root (or '/tvconfigs/country/{STANDARD}.ini').
    """
    standard = (standard or "DVB").strip().upper()
    model_ini = Path(model_ini).resolve()
    root = Path(root).resolve()

    tv_sys_map_rel, _country_path_rel_ignored = _find_tv_sys_map_and_country_path(model_ini)
    details = []

    # Force COUNTRY_PATH mapping per requirement
    forced_country_path_rel = f"/tvconfigs/country/{standard}.ini"
    country_path = _resolve_from_root(root, forced_country_path_rel)

    # Resolve tvSysMap; if missing in model.ini, try auto-locate
    if not tv_sys_map_rel:
        tv_sys_map = _auto_locate_tvsysmap(root)
        tv_sys_map_info = f"(auto) {tv_sys_map}" if tv_sys_map else "(auto) NOT FOUND"
    else:
        tv_sys_map = _resolve_from_root(root, tv_sys_map_rel)
        tv_sys_map_info = f"(from model.ini) {tv_sys_map}"

    if not tv_sys_map or not Path(tv_sys_map).exists():
        details.append(f"[ERROR] 無法取得 TvSysMap 檔案，請確認根目錄存在 TvSysMap/tvSysMapCfgs.xml。嘗試來源：{tv_sys_map_info}")
        return {"passed": False, "model_ini": str(model_ini), "tv_sys_map": str(tv_sys_map) if tv_sys_map else "",
                "country_xmls": [], "country_path": str(country_path), "customer_countries_all": [],
                "customer_target_countries": [], "allowed_countries": [], "missing": [], "details": details,
                "standard": standard}

    if not country_path.exists():
        details.append(f"[ERROR] 找不到 {standard} 名單檔: {country_path}")
        return {"passed": False, "model_ini": str(model_ini), "tv_sys_map": str(tv_sys_map),
                "country_xmls": [], "country_path": str(country_path), "customer_countries_all": [],
                "customer_target_countries": [], "allowed_countries": [], "missing": [], "details": details,
                "standard": standard}

    country_xmls = _resolve_country_tv_map_xmls(tv_sys_map, root)
    if not country_xmls:
        details.append(f"[ERROR] {Path(tv_sys_map).name}: 沒有 <CountryTvSysMapXML> 指向的檔案")
        return {"passed": False, "model_ini": str(model_ini), "tv_sys_map": str(tv_sys_map),
                "country_xmls": [], "country_path": str(country_path), "customer_countries_all": [],
                "customer_target_countries": [], "allowed_countries": [], "missing": [], "details": details,
                "standard": standard}

    pairs = []
    for p in country_xmls:
        if not p.exists():
            details.append(f"[ERROR] 缺少 countryTvSysMap.xml 檔: {p}")
            return {"passed": False, "model_ini": str(model_ini), "tv_sys_map": str(tv_sys_map),
                    "country_xmls": [str(x) for x in country_xmls], "country_path": str(country_path),
                    "customer_countries_all": [], "customer_target_countries": [], "allowed_countries": [],
                    "missing": [], "details": details, "standard": standard}
        pairs.extend(_parse_country_tv_pairs(p))

    customer_all = sorted({c for (c,_s,_p) in pairs})
    customer_target = sorted({c for (c,s,_p) in pairs if s.startswith(standard)})
    allowed_list = sorted(_parse_country_list(country_path))

    missing = sorted(set(customer_target) - set(allowed_list))
    extra_note = sorted(set(allowed_list) - set(customer_all))

    details.append(f"=== {standard} 國家比對（以國家名稱）===")
    details.append(f"Model.ini       : {model_ini}")
    details.append(f"TvSysMap        : {tv_sys_map_info}")
    for i, p in enumerate(country_xmls, 1):
        details.append(f"  Country XML[{i}]: {p}")
    details.append(f"COUNTRY_PATH    : {country_path} (forced by --standard)")
    details.append(f"- 客戶設定國家（全部）: {customer_all}")
    details.append(f"- 客戶設定國家（{standard}）: {customer_target}")
    details.append(f"- 允許的國家名單       : {allowed_list}")

    if not missing:
        details.append("✔ 檢查通過")
        if verbose and extra_note:
            details.append(f"(提示) 名單中但未在客戶映射列出的國家：{extra_note}")
            if "(auto)" in tv_sys_map_info:
                details.append(f"(提示) 沒有設定 TvSysMap")
        return {"passed": True, "model_ini": str(model_ini), "tv_sys_map": str(tv_sys_map),
                "country_xmls": [str(x) for x in country_xmls], "country_path": str(country_path),
                "customer_countries_all": customer_all, "customer_target_countries": customer_target,
                "allowed_countries": allowed_list, "missing": [], "details": details, "standard": standard}

    details.append("✖ 檢查失敗：以下國家缺漏: " + ", ".join(missing))
    return {"passed": False, "model_ini": str(model_ini), "tv_sys_map": str(tv_sys_map),
            "country_xmls": [str(x) for x in country_xmls], "country_path": str(country_path),
            "customer_countries_all": customer_all, "customer_target_countries": customer_target,
            "allowed_countries": allowed_list, "missing": missing, "details": details, "standard": standard}

def export_report(res: Dict, out_xlsx: Path):
    wb = xlsxwriter.Workbook(out_xlsx)
    ws = wb.add_worksheet("TV Std Check")
    bold = wb.add_format({"bold": True})
    ws.write(0,0,"Standard",bold); ws.write(0,1,res["standard"])
    ws.write(1,0,"Model.ini",bold); ws.write(1,1,res["model_ini"])
    ws.write(2,0,"TvSysMap",bold); ws.write(2,1,res["tv_sys_map"])
    ws.write(3,0,"Country Path",bold); ws.write(3,1,res["country_path"])
    ws.write(4,0,"Result",bold); ws.write(4,1,"PASS" if res["passed"] else "FAIL")
    ws.write(6,0,"Customer Target Countries",bold)
    ws.write(6,1,"In Allowed List?",bold)
    for idx,c in enumerate(res["customer_target_countries"],start=7):
        ws.write(idx,0,c)
        ws.write(idx,1,"YES" if c in res["allowed_countries"] else "NO")
    if res["missing"]:
        ws.write(6,3,"Missing",bold)
        for j, m in enumerate(res["missing"], start=7):
            ws.write(j,3,m)
    ws.autofilter(6,0, max(7,len(res["customer_target_countries"])+7), 3)
    wb.close()

def main(argv: List[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description="PID TV Standard Validation (DVB/ATSC/ISDB) by COUNTRY_NAME")
    ap.add_argument("--model-ini", required=True, help="Path to model.ini")
    ap.add_argument("--root", required=True, help="Project root")
    ap.add_argument("--standard", default="DVB", help="Target TV standard: DVB, ATSC, or ISDB (default: DVB)")
    ap.add_argument("-v", "--verbose", action="store_true", help="Verbose print")
    ap.add_argument("--report", action="store_true", help="Generate Excel report")
    args = ap.parse_args(argv)

    res = run_tv_standard_check(args.model_ini, args.root, standard=args.standard, verbose=args.verbose)
    for line in res["details"]:
        print(line)
    if args.report:
        std = res["standard"]
        out_xlsx = Path(f"tv_standard_check_{std}.xlsx").resolve()
        export_report(res,out_xlsx)
        print(f"[INFO] Report written: {out_xlsx}")
    return 0 if res["passed"] else 1

if __name__ == "__main__":
    sys.exit(main())
