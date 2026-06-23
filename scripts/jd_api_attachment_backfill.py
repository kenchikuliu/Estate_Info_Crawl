# -*- coding: utf-8 -*-
"""
Backfill JD auction attachment index through the public attachment API.

This only records attachment names/links/metadata. It does not download or parse
PDF/DOC files, so it is much lighter than a full attachment pipeline.
"""
from __future__ import annotations

import argparse
import json
import re
import sys
import threading
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Set

import pandas as pd
import requests

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from utils.jd_detail_parser import clean_text  # noqa: E402


API_URL = "https://api.m.jd.com/api"
DEFAULT_INPUT = r"output\京东法拍房_广东_详情回填_API.xlsx"
DEFAULT_OUTPUT = r"output\京东法拍房_广东_详情回填_API_含附件索引.xlsx"
DEFAULT_JSONL = r"output\京东法拍房_广东_附件索引_API.jsonl"
DEFAULT_CHECKPOINT = r"output\京东法拍房_广东_附件索引_API.checkpoint.json"

thread_local = threading.local()


def make_session() -> requests.Session:
    session = requests.Session()
    session.headers.update(
        {
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/148.0.0.0 Safari/537.36"
            ),
            "Accept": "application/json,text/plain,*/*",
            "Origin": "https://paimai.jd.com",
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
    if value in ("", None):
        return ""
    text = str(value).strip()
    if not text or text.lower() == "nan":
        return ""
    if text.endswith(".0") and text[:-2].isdigit():
        text = text[:-2]
    match = re.search(r"paimai\.jd\.com/(\d+)", text)
    if match:
        return match.group(1)
    return text if text.isdigit() else ""


def normalize_url(value: Any) -> str:
    text = clean_text(str(value or ""))
    if not text:
        return ""
    if text.startswith("//"):
        return "https:" + text
    if text.startswith("/"):
        return "https://paimai.jd.com" + text
    return text


def dump_json(value: Any) -> str:
    if value in ("", None, [], {}):
        return ""
    return json.dumps(value, ensure_ascii=False, separators=(",", ":"))


def join_unique(values: Iterable[Any]) -> str:
    result: List[str] = []
    seen = set()
    for value in values:
        text = clean_text(str(value or ""))
        if not text or text in seen:
            continue
        seen.add(text)
        result.append(text)
    return "；".join(result)


def read_ids_from_excel(path: Path) -> List[str]:
    df = pd.read_excel(path, usecols=lambda col: col in {"标的物ID", "链接"}).fillna("").astype(object)
    if "标的物ID" in df.columns:
        raw_values = df["标的物ID"].tolist()
    elif "链接" in df.columns:
        raw_values = df["链接"].tolist()
    else:
        raise ValueError("Input Excel must contain 标的物ID or 链接.")
    return ordered_unique_ids(raw_values)


def ordered_unique_ids(values: Iterable[Any]) -> List[str]:
    result: List[str] = []
    seen: Set[str] = set()
    for value in values:
        paimai_id = normalize_id(value)
        if not paimai_id or paimai_id in seen:
            continue
        seen.add(paimai_id)
        result.append(paimai_id)
    return result


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


def load_completed_ids(path: Path) -> Set[str]:
    return {
        normalize_id(row.get("标的物ID"))
        for row in read_jsonl(path)
        if normalize_id(row.get("标的物ID")) and row.get("附件抓取状态") != "失败"
    }


def append_jsonl(path: Path, row: Dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("a", encoding="utf-8") as handle:
        handle.write(json.dumps(row, ensure_ascii=False, separators=(",", ":")) + "\n")


def write_checkpoint(path: Path, payload: Dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    payload = dict(payload)
    payload["updated_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def fetch_attachments(paimai_id: str, retries: int = 4, timeout: int = 25) -> Dict[str, Any]:
    body = {"paimaiId": paimai_id, "custom": 9}
    encoded_body = json.dumps(body, ensure_ascii=False, separators=(",", ":"))
    params = {"appid": "paimai", "functionId": "queryAttachFilesForIntro", "loginType": "3"}
    headers = {
        "Referer": f"https://paimai.jd.com/{paimai_id}",
        "Content-Type": "application/x-www-form-urlencoded",
    }
    last_error: Optional[Exception] = None
    for attempt in range(1, retries + 1):
        try:
            response = get_session().post(
                API_URL,
                params=params,
                data={"body": encoded_body},
                headers=headers,
                timeout=timeout,
            )
            response.raise_for_status()
            payload = response.json()
            code = payload.get("code")
            if code not in (None, 0, "0"):
                raise RuntimeError(f"code={code} message={payload.get('message') or payload.get('msg')}")
            data = payload.get("data")
            attachments = data if isinstance(data, list) else []
            attachments = [item for item in attachments if isinstance(item, dict)]
            return summarize_attachments(paimai_id, attachments, "成功", "")
        except Exception as exc:  # noqa: BLE001 - keep crawler resumable.
            last_error = exc
            time.sleep(0.4 * attempt)
    return summarize_attachments(paimai_id, [], "失败", str(last_error))


def summarize_attachments(
    paimai_id: str,
    attachments: Sequence[Dict[str, Any]],
    status: str,
    error: str,
) -> Dict[str, Any]:
    cleaned: List[Dict[str, Any]] = []
    for item in attachments:
        copied = dict(item)
        if copied.get("attachmentAddress"):
            copied["attachmentAddress"] = normalize_url(copied.get("attachmentAddress"))
        cleaned.append(copied)
    return {
        "标的物ID": paimai_id,
        "链接": f"https://paimai.jd.com/{paimai_id}",
        "附件抓取时间": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "附件抓取状态": status,
        "附件抓取错误": error,
        "附件数量": len(cleaned),
        "附件名称": join_unique(item.get("attachmentName") for item in cleaned),
        "附件链接": join_unique(item.get("attachmentAddress") for item in cleaned),
        "附件索引原文": dump_json(cleaned),
    }


def export_excel(input_path: Path, jsonl_path: Path, output_path: Path) -> int:
    base = pd.read_excel(input_path).fillna("").astype(object)
    base["标的物ID"] = base["标的物ID"].map(normalize_id)
    attach_rows = read_jsonl(jsonl_path)
    if not attach_rows:
        output_path.parent.mkdir(parents=True, exist_ok=True)
        with pd.ExcelWriter(
            output_path,
            engine="xlsxwriter",
            engine_kwargs={"options": {"strings_to_urls": False}},
        ) as writer:
            base.to_excel(writer, index=False)
        return len(base)

    attach_df = pd.DataFrame(attach_rows).fillna("").astype(object)
    attach_df["标的物ID"] = attach_df["标的物ID"].map(normalize_id)
    attach_df["_order"] = range(len(attach_df))
    attach_df = attach_df.sort_values("_order").drop_duplicates("标的物ID", keep="last").drop(columns=["_order"])

    update_cols = [
        "附件抓取时间", "附件抓取状态", "附件抓取错误",
        "附件数量", "附件名称", "附件链接", "附件索引原文",
    ]
    base = base.drop(columns=[col for col in update_cols if col in base.columns], errors="ignore")
    overlap = [col for col in attach_df.columns if col in base.columns and col != "标的物ID"]
    attach_df = attach_df.drop(columns=overlap, errors="ignore")
    merged = base.merge(attach_df, on="标的物ID", how="left")

    output_path.parent.mkdir(parents=True, exist_ok=True)
    tmp_path = output_path.with_suffix(output_path.suffix + ".tmp.xlsx")
    with pd.ExcelWriter(
        tmp_path,
        engine="xlsxwriter",
        engine_kwargs={"options": {"strings_to_urls": False}},
    ) as writer:
        merged.to_excel(writer, index=False)
    tmp_path.replace(output_path)
    return len(merged)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Backfill JD auction attachment index.")
    parser.add_argument("--input", default=DEFAULT_INPUT)
    parser.add_argument("--output", default=DEFAULT_OUTPUT)
    parser.add_argument("--jsonl-output", default=DEFAULT_JSONL)
    parser.add_argument("--checkpoint", default=DEFAULT_CHECKPOINT)
    parser.add_argument("--workers", type=int, default=24)
    parser.add_argument("--save-every", type=int, default=1000)
    parser.add_argument("--start", type=int, default=0)
    parser.add_argument("--limit", type=int, default=0)
    parser.add_argument("--retries", type=int, default=4)
    parser.add_argument("--timeout", type=int, default=25)
    parser.add_argument("--no-resume", action="store_true")
    parser.add_argument("--export-only", action="store_true")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    input_path = Path(args.input)
    output_path = Path(args.output)
    jsonl_path = Path(args.jsonl_output)
    checkpoint_path = Path(args.checkpoint)

    if args.export_only:
        rows = export_excel(input_path, jsonl_path, output_path)
        print(f"attachment_export_done output={output_path} rows={rows}", flush=True)
        return

    ids = read_ids_from_excel(input_path)
    selected = ids[args.start :]
    if args.limit and args.limit > 0:
        selected = selected[: args.limit]
    completed = load_completed_ids(jsonl_path) if not args.no_resume else set()
    todo = [paimai_id for paimai_id in selected if paimai_id not in completed]
    print(
        f"attachment_plan unique_ids={len(ids)} selected={len(selected)} "
        f"completed_existing={len(completed)} todo={len(todo)} workers={args.workers}",
        flush=True,
    )

    started_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    started = time.time()
    done = 0
    failed = 0
    with ThreadPoolExecutor(max_workers=max(1, args.workers)) as executor:
        future_map = {
            executor.submit(fetch_attachments, paimai_id, args.retries, args.timeout): paimai_id
            for paimai_id in todo
        }
        for future in as_completed(future_map):
            paimai_id = future_map[future]
            try:
                row = future.result()
            except Exception as exc:  # noqa: BLE001
                row = summarize_attachments(paimai_id, [], "失败", str(exc))
            append_jsonl(jsonl_path, row)
            done += 1
            failed += row.get("附件抓取状态") == "失败"
            if done % max(1, args.save_every) == 0 or done == len(todo):
                elapsed = time.time() - started
                rate = done / elapsed if elapsed > 0 else 0
                remaining = len(todo) - done
                eta = remaining / rate if rate > 0 else 0
                write_checkpoint(
                    checkpoint_path,
                    {
                        "started_at": started_at,
                        "input": str(input_path),
                        "jsonl_output": str(jsonl_path),
                        "excel_output": str(output_path),
                        "selected": len(selected),
                        "completed_existing": len(completed),
                        "completed_this_run": done,
                        "failed_this_run": failed,
                        "todo": len(todo),
                        "remaining": remaining,
                        "elapsed_seconds": round(elapsed, 1),
                        "items_per_second": round(rate, 3),
                        "eta_seconds": round(eta, 1),
                    },
                )
                print(
                    f"attachment_progress done={done}/{len(todo)} failed={failed} "
                    f"rate={rate:.2f}/s eta={eta/60:.1f}m",
                    flush=True,
                )

    rows = export_excel(input_path, jsonl_path, output_path)
    print(f"attachment_done output={output_path} rows={rows} jsonl={jsonl_path}", flush=True)


if __name__ == "__main__":
    main()
