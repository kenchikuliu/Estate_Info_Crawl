# -*- coding: utf-8 -*-
"""
Backfill Alibaba auction fields from downloadable attachment text.

The public detail-ext endpoint often returns attachment IDs only. The actual
files can be downloaded from sf.taobao.com/download_attach.htm?attach_id=...
This script downloads a small number of attachments per item, extracts text,
parses structured fields, and writes merged JSONL rows that can be exported by
ali_sf_api_detail_backfill.py.
"""
from __future__ import annotations

import argparse
import json
import re
import sys
import threading
import time
from concurrent.futures import FIRST_COMPLETED, ThreadPoolExecutor, as_completed, wait
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Set, Tuple

import requests

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from utils.jd_detail_parser import (  # noqa: E402
    clean_text,
    extract_attachment_text,
    extract_labeled_fields,
    extract_rights_status_text,
    parse_intro_sections,
    postprocess_structured_fields,
)


DEFAULT_DETAIL_JSONL = r"output\ali_sf_gd_full_h5_20260611\阿里法拍房_广东_详情回填_API.jsonl"
DEFAULT_OUTPUT_JSONL = r"output\ali_sf_gd_full_h5_20260611\阿里法拍房_广东_附件正文回填_API.jsonl"
DEFAULT_COMBINED_JSONL = r"output\ali_sf_gd_full_h5_20260611\阿里法拍房_广东_详情回填_API_含附件正文.jsonl"
DEFAULT_CACHE = r"output\ali_sf_attachment_text_cache"

thread_local = threading.local()


def make_session() -> requests.Session:
    session = requests.Session()
    session.trust_env = False
    session.headers.update(
        {
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/148.0.0.0 Safari/537.36"
            ),
            "Accept": "*/*",
        }
    )
    return session


def get_session() -> requests.Session:
    session = getattr(thread_local, "session", None)
    if session is None:
        session = make_session()
        thread_local.session = session
    return session


def normalize_id(value: Any) -> str:
    text = str(value or "").strip()
    if text.endswith(".0") and text[:-2].isdigit():
        text = text[:-2]
    return text if text.isdigit() else ""


def safe_name(value: str, default: str = "file") -> str:
    text = clean_text(value)
    text = re.sub(r'[\\/:*?"<>|]+', "_", text)
    text = text.strip(" .")
    return text or default


def split_joined(value: Any) -> List[str]:
    text = clean_text(str(value or ""))
    if not text:
        return []
    return [part.strip() for part in re.split(r"[；;|,\n]+", text) if part.strip()]


def read_jsonl(path: Path) -> List[Dict[str, Any]]:
    if not path.exists():
        return []
    rows: List[Dict[str, Any]] = []
    with path.open("r", encoding="utf-8") as handle:
        for line in handle:
            line = line.strip()
            if not line:
                continue
            try:
                item = json.loads(line)
            except Exception:
                continue
            if isinstance(item, dict):
                rows.append(item)
    return rows


def latest_rows_by_id(path: Path) -> Dict[str, Dict[str, Any]]:
    result: Dict[str, Dict[str, Any]] = {}
    for row in read_jsonl(path):
        item_id = normalize_id(row.get("标的物ID"))
        if item_id:
            copied = dict(row)
            copied["标的物ID"] = item_id
            result[item_id] = copied
    return result


def append_jsonl(path: Path, row: Dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("a", encoding="utf-8") as handle:
        handle.write(json.dumps(row, ensure_ascii=False, separators=(",", ":")) + "\n")


def attachment_items(row: Dict[str, Any], limit: int, allowed_suffixes: Optional[Set[str]] = None) -> List[Dict[str, str]]:
    ids = split_joined(row.get("附件ID"))
    names = split_joined(row.get("附件名称"))
    candidates: List[Dict[str, str]] = []
    for index, attach_id in enumerate(ids):
        name = names[index] if index < len(names) else attach_id
        suffix = Path(name).suffix.lower()
        if suffix and suffix not in {".pdf", ".doc", ".docx", ".xls", ".xlsx", ".txt", ".csv"}:
            continue
        if allowed_suffixes is not None and suffix not in allowed_suffixes:
            continue
        candidates.append({"id": attach_id, "title": name, "suffix": suffix})
    suffix_priority = {".doc": 0, ".docx": 1, ".xls": 2, ".xlsx": 3, ".txt": 4, ".csv": 5, ".pdf": 6, "": 7}
    useful_keywords = ["评估", "估价", "询价", "调查", "现调", "房产", "不动产", "财产", "明细", "信息", "报告"]
    low_value_keywords = ["裁定", "告知", "公告", "须知", "保证金", "竞买", "悔拍", "通知书"]

    def title_priority(item: Dict[str, str]) -> Tuple[int, int, str]:
        title = item.get("title", "")
        if any(keyword in title for keyword in useful_keywords):
            title_score = 0
        elif any(keyword in title for keyword in low_value_keywords):
            title_score = 2
        else:
            title_score = 1
        return title_score, suffix_priority.get(item.get("suffix", ""), 9), title

    candidates.sort(key=title_priority)
    result = [{"id": item["id"], "title": item["title"]} for item in candidates]
    if limit > 0:
        result = result[:limit]
    return result


def download_attachment(item_id: str, attach: Dict[str, str], cache_root: Path, timeout: int, max_attachment_mb: float = 0) -> Tuple[Dict[str, Any], str]:
    attach_id = attach["id"]
    title = attach.get("title") or attach_id
    suffix = Path(title).suffix.lower() or ".bin"
    item_dir = cache_root / item_id
    item_dir.mkdir(parents=True, exist_ok=True)
    local_path = item_dir / f"{safe_name(attach_id)}_{safe_name(Path(title).stem)}{suffix}"
    url = f"https://sf.taobao.com/download_attach.htm?attach_id={attach_id}"
    meta: Dict[str, Any] = {
        "id": attach_id,
        "title": title,
        "url": url,
        "local_path": str(local_path),
        "download_status": "",
        "parsed_text_length": 0,
    }
    if not local_path.exists() or local_path.stat().st_size == 0:
        session = get_session()
        response = session.get(
            url,
            headers={"Referer": f"https://sf-item.taobao.com/sf_item/{item_id}.htm"},
            timeout=timeout,
        )
        response.raise_for_status()
        content = response.content
        if b"<html" in content[:512].lower():
            raise RuntimeError("attachment returned html")
        if max_attachment_mb > 0 and len(content) > max_attachment_mb * 1024 * 1024:
            raise RuntimeError(f"attachment too large: {len(content) / 1024 / 1024:.2f}MB")
        local_path.write_bytes(content)
    elif max_attachment_mb > 0 and local_path.stat().st_size > max_attachment_mb * 1024 * 1024:
        raise RuntimeError(f"attachment too large: {local_path.stat().st_size / 1024 / 1024:.2f}MB")
    text = extract_attachment_text(str(local_path))
    meta["download_status"] = "成功"
    meta["parsed_text_length"] = len(clean_text(text))
    return meta, text


def title_cert_candidates(title: str) -> List[str]:
    text = clean_text(title)
    candidates: List[str] = []
    patterns = [
        r"(?:房产证号|产权证号|不动产权证号|权证号)[：:\s]*([A-Za-z0-9（）()粤第\-号]+)",
        r"(?:第)([0-9A-Za-z]{4,})(?:号)",
        r"([0-9]{5,})",
    ]
    for pattern in patterns:
        for match in re.finditer(pattern, text):
            value = clean_text(match.group(1)).strip(" ：:，,。；;【】[]")
            if value and value not in candidates:
                candidates.append(value)
    return candidates


def title_floor_candidates(title: str) -> List[str]:
    text = clean_text(title)
    values: List[str] = []
    patterns = [
        r"(第[一二三四五六七八九十0-9]+层)",
        r"([0-9]+层)",
        r"([0-9]+[楼幢栋][0-9]+(?:房|室|号)?)",
    ]
    for pattern in patterns:
        for match in re.finditer(pattern, text):
            value = clean_text(match.group(1))
            if value and value not in values:
                values.append(value)
    return values


def parse_pipe_cells(line: str) -> List[str]:
    return [clean_text(part).strip(" ：:") for part in line.split("|") if clean_text(part).strip(" ：:")]


def fields_from_matching_table_row(text: str, title: str) -> Dict[str, str]:
    certs = title_cert_candidates(title)
    floors = title_floor_candidates(title)
    if not certs and not floors:
        return {}
    lines = [clean_text(line) for line in text.split("\n") if clean_text(line)]
    result: Dict[str, str] = {}
    for index, line in enumerate(lines):
        hay = line.replace(" ", "")
        matched = False
        for cert in certs:
            if cert.replace(" ", "") and cert.replace(" ", "") in hay:
                matched = True
                break
        if not matched:
            for floor in floors:
                if floor.replace(" ", "") and floor.replace(" ", "") in hay:
                    matched = True
                    break
        if not matched:
            continue
        context = " ".join(lines[max(0, index - 1): min(len(lines), index + 3)])
        cells = parse_pipe_cells(context)
        numbers = re.findall(r"(?<![0-9])([0-9]{1,4}(?:\.[0-9]{1,2})?)(?![0-9])", context)
        area_candidates = [num for num in numbers if 5 <= float(num) <= 2000]
        if area_candidates and not result.get("建筑面积"):
            result["建筑面积"] = area_candidates[0] + "平方米"
        for cell in cells:
            if cell in {"住宅", "商业", "车库", "办公", "工业", "商铺"} and not result.get("房屋用途"):
                result["房屋用途"] = cell
            if any(cert in cell for cert in certs) and not result.get("权证情况"):
                result["权证情况"] = cell
        for floor in floors:
            if floor.replace(" ", "") in context.replace(" ", "") and not result.get("所在层"):
                result["所在层"] = floor
        return result
    return result


def parse_attachment_fields(text: str, title: str) -> Dict[str, str]:
    fields: Dict[str, str] = {}
    sections = parse_intro_sections(text)
    if sections.get("权证情况"):
        fields["权证情况"] = sections["权证情况"]
    if sections.get("拍品名称"):
        fields["标的物名称"] = sections["拍品名称"]
    fields.update(extract_labeled_fields(text))
    fields.update(extract_rights_status_text(text))
    matched = fields_from_matching_table_row(text, title)
    for key, value in matched.items():
        fields[key] = value
    return postprocess_structured_fields(fields)


def merge_prefer_existing(base: Dict[str, Any], fields: Dict[str, str], force_empty_only: bool) -> Dict[str, Any]:
    result = dict(base)
    mapping = {
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
    for source, target in mapping.items():
        value = clean_text(str(fields.get(source, "")))
        if not value:
            continue
        if force_empty_only and clean_text(str(result.get(target, ""))):
            continue
        result[target] = value
    return result


def build_enrichment_row(
    row: Dict[str, Any],
    cache_root: Path,
    attachment_limit: int,
    timeout: int,
    raw_text_limit: int,
    allowed_suffixes: Optional[Set[str]] = None,
    max_attachment_mb: float = 0,
) -> Dict[str, Any]:
    item_id = normalize_id(row.get("标的物ID"))
    title = clean_text(str(row.get("标的物名称") or row.get("详情标题") or row.get("标的名称") or ""))
    attachments = attachment_items(row, attachment_limit, allowed_suffixes)
    downloaded: List[Dict[str, Any]] = []
    texts: List[str] = []
    errors: List[str] = []
    for attach in attachments:
        try:
            meta, text = download_attachment(item_id, attach, cache_root, timeout, max_attachment_mb)
            downloaded.append(meta)
            if text:
                texts.append(text)
        except Exception as exc:  # noqa: BLE001
            copied = dict(attach)
            copied["download_status"] = "失败"
            copied["error"] = str(exc)
            downloaded.append(copied)
            errors.append(f"{attach.get('id')}:{exc}")
    combined_text = clean_text("\n".join(texts))
    fields = parse_attachment_fields(combined_text, title)
    merged = merge_prefer_existing(row, fields, force_empty_only=True)
    merged.update(
        {
            "附件正文抓取时间": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "附件正文抓取状态": "成功" if texts and not errors else ("部分成功" if texts else "失败"),
            "附件正文抓取错误": "；".join(errors),
            "附件抓取成功数": sum(1 for item in downloaded if item.get("download_status") == "成功"),
            "附件解析成功数": sum(1 for item in downloaded if int(item.get("parsed_text_length") or 0) > 0),
            "附件本地路径": "；".join(clean_text(str(item.get("local_path", ""))) for item in downloaded if item.get("local_path")),
            "附件正文文本": combined_text[:raw_text_limit] if raw_text_limit > 0 else combined_text,
            "附件正文解析字段": json.dumps(fields, ensure_ascii=False, separators=(",", ":")) if fields else "",
            "附件下载原文": json.dumps(downloaded, ensure_ascii=False, separators=(",", ":")),
        }
    )
    return merged


def needs_attachment(row: Dict[str, Any], force: bool, empty_desc_only: bool) -> bool:
    if force:
        return bool(clean_text(str(row.get("附件ID", ""))))
    if not clean_text(str(row.get("附件ID", ""))):
        return False
    if empty_desc_only:
        return not clean_text(str(row.get("标的物介绍文本", "")))
    important_empty = any(not clean_text(str(row.get(field, ""))) for field in ["建筑面积", "权证情况", "标的物介绍文本"])
    return important_empty


def write_combined_jsonl(detail_jsonl: Path, attachment_jsonl: Path, combined_jsonl: Path) -> int:
    rows = latest_rows_by_id(detail_jsonl)
    for item_id, row in latest_rows_by_id(attachment_jsonl).items():
        rows[item_id] = row
    combined_jsonl.parent.mkdir(parents=True, exist_ok=True)
    with combined_jsonl.open("w", encoding="utf-8") as handle:
        for row in rows.values():
            handle.write(json.dumps(row, ensure_ascii=False, separators=(",", ":")) + "\n")
    return len(rows)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Backfill Ali attachment text and merge fields.")
    parser.add_argument("--detail-jsonl", default=DEFAULT_DETAIL_JSONL)
    parser.add_argument("--output-jsonl", default=DEFAULT_OUTPUT_JSONL)
    parser.add_argument("--combined-jsonl", default=DEFAULT_COMBINED_JSONL)
    parser.add_argument("--cache-root", default=DEFAULT_CACHE)
    parser.add_argument("--ids", default="")
    parser.add_argument("--start", type=int, default=0)
    parser.add_argument("--limit", type=int, default=0)
    parser.add_argument(
        "--completed-jsonl",
        action="append",
        default=[],
        help="Additional attachment JSONL file(s) whose IDs should be skipped.",
    )
    parser.add_argument("--workers", type=int, default=3)
    parser.add_argument("--attachment-limit", type=int, default=2)
    parser.add_argument("--allowed-suffixes", default="", help="Comma-separated attachment suffixes to process, e.g. .doc,.docx,.xls,.xlsx")
    parser.add_argument("--timeout", type=int, default=45)
    parser.add_argument("--raw-text-limit", type=int, default=4000)
    parser.add_argument("--max-attachment-mb", type=float, default=0)
    parser.add_argument("--force", action="store_true")
    parser.add_argument("--empty-desc-only", action="store_true", help="Only process rows whose detail description text is empty.")
    parser.add_argument("--combine-only", action="store_true")
    parser.add_argument("--no-combine", action="store_true", help="Do not write the combined detail+attachment JSONL after fetching.")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    detail_jsonl = Path(args.detail_jsonl)
    output_jsonl = Path(args.output_jsonl)
    combined_jsonl = Path(args.combined_jsonl)
    if args.combine_only:
        count = write_combined_jsonl(detail_jsonl, output_jsonl, combined_jsonl)
        print(f"combined_done rows={count} output={combined_jsonl}", flush=True)
        return

    rows_by_id = latest_rows_by_id(detail_jsonl)
    done_ids = set(latest_rows_by_id(output_jsonl))
    for completed_jsonl in args.completed_jsonl:
        done_ids.update(latest_rows_by_id(Path(completed_jsonl)))
    wanted = {normalize_id(part.strip()) for part in args.ids.split(",") if normalize_id(part.strip())} if args.ids else set()
    allowed_suffixes = {
        suffix.strip().lower() if suffix.strip().startswith(".") else f".{suffix.strip().lower()}"
        for suffix in args.allowed_suffixes.split(",")
        if suffix.strip()
    }
    allowed_suffixes_or_none: Optional[Set[str]] = allowed_suffixes or None
    todo = [
        row
        for item_id, row in rows_by_id.items()
        if item_id not in done_ids and (not wanted or item_id in wanted) and needs_attachment(row, args.force, args.empty_desc_only)
        and attachment_items(row, args.attachment_limit, allowed_suffixes_or_none)
    ]
    if args.start > 0:
        todo = todo[args.start :]
    if args.limit and args.limit > 0:
        todo = todo[: args.limit]
    print(f"attachment_text_plan source={len(rows_by_id)} done_existing={len(done_ids)} todo={len(todo)} workers={args.workers}", flush=True)

    completed = 0
    failed = 0
    started = time.time()
    cache_root = Path(args.cache_root)
    max_pending = max(1, args.workers) * 4
    with ThreadPoolExecutor(max_workers=max(1, args.workers)) as executor:
        future_map: Dict[Any, Dict[str, Any]] = {}
        for row in todo:
            future_map[
                executor.submit(
                    build_enrichment_row,
                    row,
                    cache_root,
                    args.attachment_limit,
                    args.timeout,
                    args.raw_text_limit,
                    allowed_suffixes_or_none,
                    args.max_attachment_mb,
                )
            ] = row
            if len(future_map) >= max_pending:
                done, _ = wait(future_map.keys(), return_when=FIRST_COMPLETED)
                for future in done:
                    base = future_map.pop(future)
                    try:
                        enriched = future.result()
                    except Exception as exc:  # noqa: BLE001
                        failed += 1
                        enriched = dict(base)
                        enriched.update(
                            {
                                "附件正文抓取时间": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                "附件正文抓取状态": "失败",
                                "附件正文抓取错误": str(exc),
                            }
                        )
                    append_jsonl(output_jsonl, enriched)
                    completed += 1
                    if completed % 50 == 0:
                        elapsed = time.time() - started
                        print(f"attachment_text_progress done={completed}/{len(todo)} failed={failed} rate={completed/elapsed:.2f}/s", flush=True)

        for future in as_completed(future_map):
            base = future_map[future]
            try:
                enriched = future.result()
            except Exception as exc:  # noqa: BLE001
                failed += 1
                enriched = dict(base)
                enriched.update(
                    {
                        "附件正文抓取时间": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "附件正文抓取状态": "失败",
                        "附件正文抓取错误": str(exc),
                    }
                )
            append_jsonl(output_jsonl, enriched)
            completed += 1

    if args.no_combine:
        print(f"attachment_text_done done={completed} failed={failed} output={output_jsonl} combined_skipped=true", flush=True)
        return
    combined_count = write_combined_jsonl(detail_jsonl, output_jsonl, combined_jsonl)
    print(f"attachment_text_done done={completed} failed={failed} combined_rows={combined_count} output={output_jsonl}", flush=True)


if __name__ == "__main__":
    main()
