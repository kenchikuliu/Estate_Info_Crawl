# -*- coding: utf-8 -*-
"""
Second-pass Ali attachment enrichment from the existing local cache only.

This script does not download anything. It reparses cached attachment files and
streams a new combined JSONL with extra real values filled where the current row
is blank or only has a semantic placeholder.
"""
from __future__ import annotations

import argparse
import json
import re
import sys
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from scripts.ali_sf_attachment_text_backfill import parse_attachment_fields  # noqa: E402
from utils.jd_detail_parser import clean_text, extract_attachment_text  # noqa: E402


DEFAULT_INPUT_JSONL = PROJECT_ROOT / r"output\ali_sf_gd_full_h5_20260611\阿里法拍房_广东_详情回填_API_含附件正文.jsonl"
DEFAULT_OUTPUT_JSONL = PROJECT_ROOT / r"output\ali_sf_gd_full_h5_20260611\阿里法拍房_广东_详情回填_API_含附件正文_二次解析.jsonl"
DEFAULT_OVERLAY_JSONL = PROJECT_ROOT / r"output\ali_sf_gd_full_h5_20260611\阿里法拍房_广东_附件正文二次解析_API.jsonl"
DEFAULT_REPORT = PROJECT_ROOT / r"output\ali_sf_gd_full_h5_20260611\阿里法拍房_广东_附件正文二次解析_API.report.json"
DEFAULT_CACHE = PROJECT_ROOT / r"output\ali_sf_attachment_text_cache"

ALLOWED_SUFFIXES = {".doc", ".docx", ".xls", ".xlsx", ".pdf", ".txt", ".csv"}
PRICE_LABELS = {
    "评估价_元": [
        "评估价",
        "评估价值",
        "评估总价",
        "估价结果",
        "评估结果",
    ],
    "市场价_元": [
        "市场价值",
        "市场价",
        "市场总价",
        "询价结果",
        "议价结果",
        "参考价",
        "处置参考价",
    ],
}
FINANCE_KEYWORDS = ["金融服务", "一键贷款", "最高可贷", "参考利率", "贷款期限", "月供", "可贷比例", "银行贷款"]
PLACEHOLDERS = {"", "未知", "无", "不适用", "nan", "None", "null"}
FIELD_MAP = {
    "建筑面积": "建筑面积",
    "权证情况": "权证情况",
    "房屋用途": "房屋用途",
    "房屋类型": "房屋类型",
    "所在层": "所在层",
    "总层数": "总层数",
    "竣工时间": "竣工时间",
    "购买时间": "购买时间",
    "土地性质": "土地性质",
    "土地用途": "土地用途",
    "使用期限": "使用期限",
    "权利来源": "权利来源",
    "所有权来源": "所有权来源",
    "钥匙": "钥匙/占用情况",
    "腾空情况": "腾空情况",
    "户籍注册": "户籍/工商注册",
    "欠费情况": "欠费情况",
    "提供文件": "提供文件",
    "权利限制状况及抵押状况": "权利限制状况及抵押状况",
    "房屋权属状况": "房屋权属状况",
    "土地权属状况": "土地权属状况",
}


def normalize_id(value: Any) -> str:
    text = str(value or "").strip()
    if text.endswith(".0") and text[:-2].isdigit():
        text = text[:-2]
    return text if text.isdigit() else text


def is_blankish(value: Any) -> bool:
    return clean_text(str(value or "")) in PLACEHOLDERS


def write_jsonl_line(handle: Any, row: Dict[str, Any]) -> None:
    handle.write(json.dumps(row, ensure_ascii=False, separators=(",", ":")) + "\n")


def money_to_yuan(value: str) -> Optional[float]:
    text = clean_text(value).replace(",", "").replace("，", "")
    match = re.search(r"([0-9]+(?:\.[0-9]+)?)", text)
    if not match:
        return None
    number = float(match.group(1))
    if "万" in text:
        number *= 10000
    if number < 1000 or number > 100000000000:
        return None
    return round(number, 2)


def extract_price(text: str, labels: Sequence[str]) -> str:
    if not text:
        return ""
    label_pattern = "|".join(re.escape(label) for label in labels)
    pattern = re.compile(
        rf"(?:{label_pattern})(?:合计|总计|总额|为|是|：|:|\s)*"
        rf"(?:人民币|约)?\s*([0-9][0-9,，]*(?:\.[0-9]+)?\s*(?:万元|万|元)?)"
    )
    for match in pattern.finditer(text):
        value = money_to_yuan(match.group(1))
        if value is None:
            continue
        return str(int(value) if float(value).is_integer() else value)
    return ""


def finance_snippets(text: str) -> Tuple[str, Dict[str, str]]:
    snippets: List[str] = []
    compact = re.sub(r"\s+", " ", text or "")
    for keyword in FINANCE_KEYWORDS:
        index = compact.find(keyword)
        if index < 0:
            continue
        snippet = compact[max(0, index - 80) : min(len(compact), index + 240)].strip()
        if snippet and snippet not in snippets:
            snippets.append(snippet)
        if len(snippets) >= 3:
            break
    joined = "；".join(snippets)
    fields: Dict[str, str] = {}
    if joined:
        match = re.search(r"([\u4e00-\u9fff]{2,24}(?:银行|金融|担保|贷款)[\u4e00-\u9fff]{0,16})", joined)
        if match:
            fields["金融机构"] = match.group(1)
        match = re.search(r"(?:最高可贷|可贷比例|贷款成数)[^0-9一二三四五六七八九十]{0,20}([0-9一二三四五六七八九十]{1,3}\s*(?:成|%|％))", joined)
        if match:
            fields["最高可贷比例"] = match.group(1).replace(" ", "")
        match = re.search(r"(?:参考利率|利率)[^0-9]{0,20}([0-9]+(?:\.[0-9]+)?\s*[%％])", joined)
        if match:
            fields["参考利率"] = match.group(1).replace(" ", "")
    return joined, fields


def candidate_files(item_dir: Path, max_attachment_mb: float, file_limit: int) -> List[Path]:
    if not item_dir.exists():
        return []
    files = [
        path
        for path in item_dir.iterdir()
        if path.is_file()
        and path.suffix.lower() in ALLOWED_SUFFIXES
        and (max_attachment_mb <= 0 or path.stat().st_size <= max_attachment_mb * 1024 * 1024)
    ]
    useful = ["评估", "估价", "询价", "调查", "现调", "房产", "不动产", "财产", "明细", "信息", "报告"]
    suffix_rank = {".doc": 0, ".docx": 1, ".xls": 2, ".xlsx": 3, ".txt": 4, ".csv": 5, ".pdf": 6}

    def rank(path: Path) -> Tuple[int, int, str]:
        name = path.name
        return (0 if any(key in name for key in useful) else 1, suffix_rank.get(path.suffix.lower(), 9), name)

    files.sort(key=rank)
    if file_limit > 0:
        return files[:file_limit]
    return files


def parse_cached_item(item_id: str, title: str, cache_root: Path, max_attachment_mb: float, file_limit: int, raw_text_limit: int) -> Dict[str, Any]:
    files = candidate_files(cache_root / item_id, max_attachment_mb=max_attachment_mb, file_limit=file_limit)
    texts: List[str] = []
    parsed_files: List[Dict[str, Any]] = []
    errors: List[str] = []
    for path in files:
        try:
            text = extract_attachment_text(str(path))
            text = clean_text(text)
            parsed_files.append({"path": str(path), "size": path.stat().st_size, "text_length": len(text)})
            if text:
                texts.append(f"【{path.name}】\n{text}")
        except Exception as exc:  # noqa: BLE001
            errors.append(f"{path.name}:{exc}")
    combined_text = clean_text("\n".join(texts))
    fields = parse_attachment_fields(combined_text, title) if combined_text else {}
    for target, labels in PRICE_LABELS.items():
        value = extract_price(combined_text, labels)
        if value:
            fields[target] = value
            if target == "评估价_元":
                fields.setdefault("评估价", value)
                fields.setdefault("评估价_详情_元", value)
            if target == "市场价_元":
                fields.setdefault("市场价", value)
                fields.setdefault("市场价_详情_元", value)
    finance_text, finance_fields = finance_snippets(combined_text)
    if finance_text:
        fields["金融服务原文"] = finance_text
        fields["是否有金融服务_详情"] = "是"
        fields["是否有金融服务"] = "是"
        fields.update(finance_fields)
    return {
        "fields": fields,
        "text": combined_text[:raw_text_limit] if raw_text_limit > 0 else combined_text,
        "parsed_files": parsed_files,
        "errors": errors,
    }


def merge_fields(row: Dict[str, Any], parsed: Dict[str, Any]) -> Tuple[Dict[str, Any], List[str]]:
    result = dict(row)
    changed: List[str] = []
    fields = parsed.get("fields", {})
    for source, target in FIELD_MAP.items():
        value = clean_text(str(fields.get(source, "")))
        if value and is_blankish(result.get(target, "")):
            result[target] = value
            changed.append(target)
    for target in [
        "评估价",
        "评估价_元",
        "评估价_详情_元",
        "市场价",
        "市场价_元",
        "市场价_详情_元",
        "金融服务原文",
        "是否有金融服务",
        "是否有金融服务_详情",
        "金融机构",
        "最高可贷比例",
        "参考利率",
        "金融其他费用",
    ]:
        value = clean_text(str(fields.get(target, "")))
        if value and is_blankish(result.get(target, "")):
            result[target] = value
            changed.append(target)
    if parsed.get("text") and is_blankish(result.get("附件正文文本", "")):
        result["附件正文文本"] = parsed["text"]
        changed.append("附件正文文本")
    if fields:
        result["附件正文二次解析字段"] = json.dumps(fields, ensure_ascii=False, separators=(",", ":"))
    result["附件正文二次解析时间"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    result["附件正文二次解析文件数"] = len(parsed.get("parsed_files", []))
    result["附件正文二次解析错误"] = "；".join(parsed.get("errors", []))
    if changed:
        result["附件正文二次解析更新列"] = "；".join(changed)
    return result, changed


def load_candidates(input_jsonl: Path, cache_root: Path, limit: int, only_missing: bool) -> List[Dict[str, Any]]:
    candidates: List[Dict[str, Any]] = []
    with input_jsonl.open("r", encoding="utf-8") as handle:
        for line in handle:
            line = line.strip()
            if not line:
                continue
            try:
                row = json.loads(line)
            except Exception:
                continue
            if not isinstance(row, dict):
                continue
            item_id = normalize_id(row.get("标的物ID"))
            if not item_id or not (cache_root / item_id).exists():
                continue
            if only_missing and not any(
                is_blankish(row.get(field, ""))
                for field in ["评估价_元", "市场价_元", "建筑面积", "权证情况", "房屋用途", "所在层", "金融服务原文"]
            ):
                continue
            row["_二次解析ID"] = item_id
            candidates.append(row)
            if limit > 0 and len(candidates) >= limit:
                break
    return candidates


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Reparse Ali local attachment cache and write enhanced combined JSONL.")
    parser.add_argument("--input-jsonl", default=str(DEFAULT_INPUT_JSONL))
    parser.add_argument("--output-jsonl", default=str(DEFAULT_OUTPUT_JSONL))
    parser.add_argument("--overlay-jsonl", default=str(DEFAULT_OVERLAY_JSONL))
    parser.add_argument("--report", default=str(DEFAULT_REPORT))
    parser.add_argument("--cache-root", default=str(DEFAULT_CACHE))
    parser.add_argument("--workers", type=int, default=4)
    parser.add_argument("--limit", type=int, default=0)
    parser.add_argument("--file-limit", type=int, default=0)
    parser.add_argument("--raw-text-limit", type=int, default=8000)
    parser.add_argument("--max-attachment-mb", type=float, default=20)
    parser.add_argument("--only-missing", action="store_true")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    input_jsonl = Path(args.input_jsonl)
    output_jsonl = Path(args.output_jsonl)
    overlay_jsonl = Path(args.overlay_jsonl)
    report = Path(args.report)
    cache_root = Path(args.cache_root)
    started = time.time()

    candidates = load_candidates(input_jsonl, cache_root, args.limit, args.only_missing)
    print(f"local_second_pass_plan candidates={len(candidates)} workers={args.workers}", flush=True)

    overlays: Dict[str, Dict[str, Any]] = {}
    changed_by_column: Dict[str, int] = {}
    parsed = 0
    with ThreadPoolExecutor(max_workers=max(1, args.workers)) as executor:
        future_map = {}
        for row in candidates:
            item_id = row["_二次解析ID"]
            title = clean_text(str(row.get("标的物名称") or row.get("详情标题") or row.get("标题") or ""))
            future = executor.submit(parse_cached_item, item_id, title, cache_root, args.max_attachment_mb, args.file_limit, args.raw_text_limit)
            future_map[future] = row
        for future in as_completed(future_map):
            row = future_map[future]
            item_id = row["_二次解析ID"]
            parsed += 1
            try:
                parsed_payload = future.result()
            except Exception as exc:  # noqa: BLE001
                parsed_payload = {"fields": {}, "text": "", "parsed_files": [], "errors": [str(exc)]}
            merged, changed = merge_fields(row, parsed_payload)
            merged.pop("_二次解析ID", None)
            if changed:
                overlays[item_id] = merged
                for column in changed:
                    changed_by_column[column] = changed_by_column.get(column, 0) + 1
            if parsed % 500 == 0:
                elapsed = time.time() - started
                print(f"local_second_pass_progress parsed={parsed}/{len(candidates)} changed_rows={len(overlays)} rate={parsed/elapsed:.2f}/s", flush=True)

    output_jsonl.parent.mkdir(parents=True, exist_ok=True)
    overlay_jsonl.parent.mkdir(parents=True, exist_ok=True)
    rows_written = 0
    with overlay_jsonl.open("w", encoding="utf-8") as overlay_handle:
        for row in overlays.values():
            write_jsonl_line(overlay_handle, row)
    with input_jsonl.open("r", encoding="utf-8") as input_handle, output_jsonl.open("w", encoding="utf-8") as output_handle:
        for line in input_handle:
            stripped = line.strip()
            if not stripped:
                continue
            try:
                row = json.loads(stripped)
            except Exception:
                continue
            item_id = normalize_id(row.get("标的物ID"))
            write_jsonl_line(output_handle, overlays.get(item_id, row))
            rows_written += 1

    payload = {
        "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "input_jsonl": str(input_jsonl),
        "output_jsonl": str(output_jsonl),
        "overlay_jsonl": str(overlay_jsonl),
        "cache_root": str(cache_root),
        "candidates": len(candidates),
        "changed_rows": len(overlays),
        "rows_written": rows_written,
        "changed_by_column": dict(sorted(changed_by_column.items())),
        "elapsed_seconds": round(time.time() - started, 2),
    }
    report.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(payload, ensure_ascii=False, indent=2), flush=True)


if __name__ == "__main__":
    main()
