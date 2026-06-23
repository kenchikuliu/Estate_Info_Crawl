# -*- coding: utf-8 -*-
"""
Backfill JD auction detail fields through the public detail APIs.

The crawler writes one JSON line per finished asset before exporting Excel, so
large Guangdong runs can be resumed without losing completed detail records.
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
from typing import Any, Dict, Iterable, List, Optional, Sequence, Set, Tuple

import pandas as pd
import requests
from bs4 import BeautifulSoup

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from utils.jd_detail_parser import (  # noqa: E402
    clean_text,
    extract_labeled_fields,
    extract_region_fields,
    extract_rights_status_text,
    parse_intro_sections,
    parse_survey_table_rows,
    postprocess_structured_fields,
)


API_URL = "https://api.m.jd.com/api"
DEFAULT_INPUT = r"output\京东法拍房_广东_全量索引_API合并.xlsx"
DEFAULT_OUTPUT = r"output\京东法拍房_广东_详情回填_API.xlsx"
DEFAULT_JSONL = r"output\京东法拍房_广东_详情回填_API.jsonl"
DEFAULT_CHECKPOINT = r"output\京东法拍房_广东_详情回填_API.checkpoint.json"

DETAIL_COLUMNS = [
    "详情抓取时间",
    "详情抓取状态",
    "详情接口错误",
    "详情标题",
    "标的名称",
    "标的物名称",
    "权证情况",
    "标的所有人",
    "被执行人/标的所有人",
    "详情地址",
    "详情省",
    "详情市",
    "详情区县",
    "经度",
    "纬度",
    "经纬度来源",
    "法院/处置机构",
    "拍卖状态_详情",
    "auctionStatus_详情",
    "displayStatus_详情",
    "是否成交",
    "是否流拍",
    "成交价/获拍价_元",
    "成交时间",
    "当前价_详情_元",
    "起拍价_详情_元",
    "保证金_详情_元",
    "评估价_详情_元",
    "出价次数_详情",
    "围观次数_详情",
    "成交确认书链接",
    "是否有金融服务_详情",
    "金融机构",
    "最高可贷比例",
    "参考利率",
    "金融其他费用",
    "金融服务原文",
    "建筑面积",
    "房屋用途",
    "房屋类型",
    "所在层",
    "总层数",
    "竣工时间",
    "购买时间",
    "土地性质",
    "土地用途",
    "使用期限",
    "权利来源",
    "所有权来源",
    "钥匙/占用情况",
    "腾空情况",
    "户籍/工商注册",
    "欠费情况",
    "提供文件",
    "权利限制状况及抵押状况",
    "房屋权属状况",
    "土地权属状况",
    "附件数量",
    "附件名称",
    "附件链接",
    "附件索引原文",
    "标的物介绍文本",
    "竞买公告文本",
]

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


def dump_json(value: Any) -> str:
    if value in ("", None, [], {}):
        return ""
    return json.dumps(value, ensure_ascii=False, separators=(",", ":"))


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


def first_non_empty(*values: Any) -> Any:
    for value in values:
        if value in (None, ""):
            continue
        if isinstance(value, float) and pd.isna(value):
            continue
        return value
    return ""


def to_int(value: Any, default: int = 0) -> int:
    try:
        if value in ("", None):
            return default
        return int(float(value))
    except Exception:
        return default


def ts_to_text(value: Any) -> str:
    if value in ("", None):
        return ""
    try:
        number = float(value)
    except (TypeError, ValueError):
        return ""
    if number <= 0:
        return ""
    if number > 10_000_000_000:
        number = number / 1000
    return datetime.fromtimestamp(number).strftime("%Y-%m-%d %H:%M:%S")


def status_text(auction_status: Any, display_status: Any) -> str:
    auction = to_int(auction_status, -1)
    display = to_int(display_status, -1)
    display_map = {3: "已暂停", 5: "已撤回", 6: "已暂缓", 7: "已中止"}
    if display in display_map:
        return display_map[display]
    return {0: "预告中", 1: "进行中", 2: "已结束"}.get(auction, "")


def api_call(
    function_id: str,
    paimai_id: str,
    body: Dict[str, Any],
    method: str = "POST",
    retries: int = 4,
    timeout: int = 25,
    sleep_seconds: float = 0.5,
) -> Any:
    last_error: Optional[Exception] = None
    params = {
        "appid": "paimai",
        "functionId": function_id,
        "loginType": "3",
    }
    encoded_body = json.dumps(body, ensure_ascii=False, separators=(",", ":"))
    headers = {
        "Referer": f"https://paimai.jd.com/{paimai_id}",
        "Content-Type": "application/x-www-form-urlencoded",
    }
    for attempt in range(1, retries + 1):
        try:
            session = get_session()
            if method.upper() == "GET":
                response = session.get(
                    API_URL,
                    params={**params, "body": encoded_body},
                    headers=headers,
                    timeout=timeout,
                )
            else:
                response = session.post(
                    API_URL,
                    params=params,
                    data={"body": encoded_body},
                    headers=headers,
                    timeout=timeout,
                )
            response.raise_for_status()
            payload = response.json()
            code = payload.get("code")
            result_code = payload.get("resultCode")
            if code not in (None, 0, "0") and result_code not in (None, "0000", 0, "0"):
                raise RuntimeError(f"{function_id} code={code} resultCode={result_code}")
            return payload.get("data")
        except Exception as exc:  # noqa: BLE001 - crawler should retry and keep moving.
            last_error = exc
            time.sleep(sleep_seconds * attempt)
    raise RuntimeError(f"{function_id} failed for {paimai_id}: {last_error}")


def try_api(
    errors: List[str],
    label: str,
    function_id: str,
    paimai_id: str,
    body: Dict[str, Any],
    method: str = "POST",
    retries: int = 4,
    timeout: int = 25,
) -> Any:
    try:
        return api_call(function_id, paimai_id, body, method=method, retries=retries, timeout=timeout)
    except Exception as exc:  # noqa: BLE001
        errors.append(f"{label}:{exc}")
        return None


def html_to_text_and_rows(html: str) -> Tuple[str, List[List[str]]]:
    if not html:
        return "", []
    soup = BeautifulSoup(html, "lxml")
    for tag in soup(["script", "style"]):
        tag.decompose()

    rows: List[List[str]] = []
    for tr in soup.find_all("tr"):
        cells = [clean_text(cell.get_text("\n")) for cell in tr.find_all(["td", "th"])]
        cells = [cell for cell in cells if cell]
        if cells:
            rows.append(cells)

    text = clean_text(soup.get_text("\n"))
    return text, rows


def parse_description_fields(html: str) -> Tuple[Dict[str, str], str]:
    text, rows = html_to_text_and_rows(html)
    fields: Dict[str, str] = {}

    sections = parse_intro_sections(text)
    if sections.get("拍品名称"):
        fields["标的物名称"] = sections["拍品名称"]
    if sections.get("权证情况"):
        fields["权证情况"] = sections["权证情况"]
    if sections.get("拍品所有人"):
        fields["被执行人"] = sections["拍品所有人"]
    if sections.get("成交后提供的文件"):
        fields["提供文件"] = sections["成交后提供的文件"]

    fields.update(extract_labeled_fields(text))
    fields.update(parse_survey_table_rows(rows))
    fields.update(extract_rights_status_text(text))
    return postprocess_structured_fields(fields), text


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


def summarize_finance(banks: Sequence[Dict[str, Any]]) -> Dict[str, str]:
    banks = [bank for bank in banks if isinstance(bank, dict)]
    return {
        "金融机构": join_unique(bank.get("bankName") for bank in banks),
        "最高可贷比例": join_unique(bank.get("maxLoanRatio") for bank in banks),
        "参考利率": join_unique(bank.get("loanRate") for bank in banks),
        "金融其他费用": join_unique(bank.get("otherExpenses") for bank in banks),
        "金融服务原文": dump_json(banks),
    }


def summarize_attachments(attachments: Sequence[Dict[str, Any]]) -> Dict[str, str]:
    attachments = [item for item in attachments if isinstance(item, dict)]
    names = join_unique(item.get("attachmentName") for item in attachments)
    urls = join_unique(normalize_url(item.get("attachmentAddress")) for item in attachments)
    return {
        "附件数量": len(attachments),
        "附件名称": names,
        "附件链接": urls,
        "附件索引原文": dump_json(attachments),
    }


def address_fields(basic: Dict[str, Any]) -> Dict[str, str]:
    address_result = basic.get("productAddressResult") or {}
    address = first_non_empty(
        basic.get("productAddress"),
        address_result.get("address"),
        basic.get("address"),
        basic.get("title"),
    )
    province = clean_text(str(address_result.get("province") or ""))
    city = clean_text(str(address_result.get("city") or ""))
    county = clean_text(str(address_result.get("county") or ""))
    region = extract_region_fields(str(address or ""))
    return {
        "详情地址": clean_text(str(address or "")),
        "详情省": province or region.get("省", ""),
        "详情市": city or region.get("市", ""),
        "详情区县": county or region.get("区县", ""),
    }


def best_bid(bid_list: Any) -> Dict[str, Any]:
    if not isinstance(bid_list, list) or not bid_list:
        return {}
    candidates = [bid for bid in bid_list if isinstance(bid, dict)]
    if not candidates:
        return {}

    def bid_key(bid: Dict[str, Any]) -> Tuple[float, int]:
        try:
            price = float(bid.get("price") or 0)
        except Exception:
            price = 0.0
        return price, to_int(bid.get("bidTime"), 0)

    return max(candidates, key=bid_key)


def truncate_text(text: str, limit: int) -> str:
    normalized = clean_text(text)
    if limit <= 0 or len(normalized) <= limit:
        return normalized
    return normalized[:limit] + "..."


def fetch_detail(
    paimai_id: str,
    include_finance: bool,
    include_notice: bool,
    include_attachments: bool,
    include_description: bool,
    raw_text_limit: int,
    retries: int,
    timeout: int,
) -> Dict[str, Any]:
    errors: List[str] = []
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    basic = try_api(errors, "basic", "getProductBasicInfo", paimai_id, {"paimaiId": paimai_id}, retries=retries, timeout=timeout)
    realtime = try_api(errors, "realtime", "getPaimaiRealTimeData", paimai_id, {"paimaiId": paimai_id}, retries=retries, timeout=timeout)
    basic = basic if isinstance(basic, dict) else {}
    realtime = realtime if isinstance(realtime, dict) else {}

    description_fields: Dict[str, str] = {}
    description_text = ""
    if include_description:
        description_html = try_api(
            errors,
            "description",
            "queryProductDescription",
            paimai_id,
            {"paimaiId": paimai_id, "source": 5},
            retries=retries,
            timeout=timeout,
        )
        if isinstance(description_html, str):
            description_fields, description_text = parse_description_fields(description_html)

    notice_text = ""
    if include_notice:
        notice_html = try_api(errors, "notice", "queryNotice", paimai_id, {"paimaiId": paimai_id}, retries=retries, timeout=timeout)
        if isinstance(notice_html, str):
            notice_text, _ = html_to_text_and_rows(notice_html)

    attachments: List[Dict[str, Any]] = []
    if include_attachments:
        attach_data = try_api(
            errors,
            "attachments",
            "queryAttachFilesForIntro",
            paimai_id,
            {"paimaiId": paimai_id, "custom": 9},
            retries=retries,
            timeout=timeout,
        )
        if isinstance(attach_data, list):
            attachments = [item for item in attach_data if isinstance(item, dict)]

    loan_flag: Optional[bool] = None
    banks: List[Dict[str, Any]] = []
    if include_finance:
        loan_data = try_api(
            errors,
            "loan_flag",
            "wareWhetherLoan",
            paimai_id,
            {"paimaiId": paimai_id},
            method="GET",
            retries=retries,
            timeout=timeout,
        )
        if isinstance(loan_data, bool):
            loan_flag = loan_data
        if loan_flag:
            bank_data = try_api(
                errors,
                "banks",
                "getMergedBankList",
                paimai_id,
                {"paimaiId": paimai_id},
                method="GET",
                retries=retries,
                timeout=timeout,
            )
            if isinstance(bank_data, list):
                banks = [item for item in bank_data if isinstance(item, dict)]

    auction_status = first_non_empty(realtime.get("auctionStatus"), basic.get("auctionStatus"))
    display_status = first_non_empty(realtime.get("displayStatus"), basic.get("displayStatus"))
    bid_count = to_int(first_non_empty(realtime.get("bidCount"), basic.get("bidCount")), 0)
    bid = best_bid(realtime.get("bidList"))
    current_price = first_non_empty(realtime.get("currentPrice"), basic.get("currentPrice"))
    start_price = first_non_empty(realtime.get("startPrice"), basic.get("startPrice"))
    end_time = first_non_empty(realtime.get("endTime"), basic.get("endTime"))
    ended = to_int(auction_status, -1) == 2
    special_display = to_int(display_status, -1) in {3, 5, 6, 7}
    sold = ended and bid_count > 0 and not special_display
    unsold = ended and bid_count == 0 and not special_display
    deal_price = first_non_empty(bid.get("price"), current_price) if sold else ""
    deal_time = first_non_empty(ts_to_text(bid.get("bidTime")), ts_to_text(end_time)) if sold else ""

    lat = first_non_empty(basic.get("lat"), basic.get("latitude"))
    lng = first_non_empty(basic.get("lng"), basic.get("longitude"))
    finance_summary = summarize_finance(banks)
    attachment_summary = summarize_attachments(attachments)
    title = first_non_empty(basic.get("title"), basic.get("productName"), realtime.get("title"))
    owner = description_fields.get("被执行人", "")

    row: Dict[str, Any] = {
        "标的物ID": paimai_id,
        "链接": f"https://paimai.jd.com/{paimai_id}",
        "详情抓取时间": now,
        "详情抓取状态": "成功" if not errors else "部分成功",
        "详情接口错误": " | ".join(errors),
        "详情标题": title,
        "标的名称": first_non_empty(description_fields.get("标的物名称"), title),
        "标的物名称": first_non_empty(description_fields.get("标的物名称"), title),
        "权证情况": description_fields.get("权证情况", ""),
        "标的所有人": owner,
        "被执行人/标的所有人": owner,
        **address_fields(basic),
        "经度": lng,
        "纬度": lat,
        "经纬度来源": "京东详情API" if lng and lat else "",
        "法院/处置机构": first_non_empty(
            basic.get("courtName"),
            basic.get("courtVendorName"),
            basic.get("vendorName"),
            basic.get("shopName"),
        ),
        "拍卖状态_详情": status_text(auction_status, display_status),
        "auctionStatus_详情": auction_status,
        "displayStatus_详情": display_status,
        "是否成交": "是" if sold else "否",
        "是否流拍": "是" if unsold else "否",
        "成交价/获拍价_元": deal_price,
        "成交时间": deal_time,
        "当前价_详情_元": current_price,
        "起拍价_详情_元": start_price,
        "保证金_详情_元": first_non_empty(basic.get("ensurePrice"), realtime.get("ensurePrice")),
        "评估价_详情_元": basic.get("assessmentPrice", ""),
        "出价次数_详情": bid_count,
        "围观次数_详情": first_non_empty(realtime.get("accessNum"), basic.get("accessNum")),
        "成交确认书链接": normalize_url(realtime.get("confirmationUrl")),
        "是否有金融服务_详情": "是" if (loan_flag or banks) else "否",
        **finance_summary,
        "建筑面积": description_fields.get("建筑面积", ""),
        "房屋用途": description_fields.get("房屋用途", ""),
        "房屋类型": description_fields.get("房屋类型", ""),
        "所在层": description_fields.get("所在层", ""),
        "总层数": description_fields.get("总层数", ""),
        "竣工时间": description_fields.get("竣工时间", ""),
        "购买时间": description_fields.get("购买时间", ""),
        "土地性质": description_fields.get("土地性质", ""),
        "土地用途": description_fields.get("土地用途", ""),
        "使用期限": description_fields.get("使用期限", ""),
        "权利来源": description_fields.get("权利来源", ""),
        "所有权来源": description_fields.get("所有权来源", ""),
        "钥匙/占用情况": description_fields.get("钥匙", ""),
        "腾空情况": description_fields.get("腾空情况", ""),
        "户籍/工商注册": description_fields.get("户籍注册", ""),
        "欠费情况": description_fields.get("欠费情况", ""),
        "提供文件": description_fields.get("提供文件", ""),
        "权利限制状况及抵押状况": description_fields.get("权利限制状况及抵押状况", ""),
        "房屋权属状况": description_fields.get("房屋权属状况", ""),
        "土地权属状况": description_fields.get("土地权属状况", ""),
        **attachment_summary,
        "标的物介绍文本": truncate_text(description_text, raw_text_limit),
        "竞买公告文本": truncate_text(notice_text, raw_text_limit),
    }
    return row


def find_column(columns: Sequence[str], candidates: Sequence[str]) -> str:
    for candidate in candidates:
        if candidate in columns:
            return candidate
    return ""


def read_index(path: Path) -> pd.DataFrame:
    df = pd.read_excel(path).fillna("").astype(object)
    id_col = find_column(list(df.columns), ["标的物ID", "拍品ID", "paimaiId", "id"])
    link_col = find_column(list(df.columns), ["链接", "详情链接", "url", "URL"])
    if id_col:
        df["标的物ID"] = df[id_col].map(normalize_id)
    elif link_col:
        df["标的物ID"] = df[link_col].map(normalize_id)
    else:
        raise ValueError("Input Excel must contain 标的物ID or 链接.")
    if "链接" not in df.columns:
        df["链接"] = df["标的物ID"].map(lambda value: f"https://paimai.jd.com/{value}" if value else "")
    return df[df["标的物ID"].astype(str).str.len() > 0].copy()


def ordered_unique_ids(values: Iterable[Any]) -> List[str]:
    ids: List[str] = []
    seen: Set[str] = set()
    for value in values:
        paimai_id = normalize_id(value)
        if not paimai_id or paimai_id in seen:
            continue
        seen.add(paimai_id)
        ids.append(paimai_id)
    return ids


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
        if normalize_id(row.get("标的物ID")) and row.get("详情抓取状态") != "失败"
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


def export_excel(index_df: pd.DataFrame, jsonl_path: Path, output_path: Path, completed_only: bool) -> int:
    detail_rows = read_jsonl(jsonl_path)
    if not detail_rows:
        output_path.parent.mkdir(parents=True, exist_ok=True)
        pd.DataFrame().to_excel(output_path, index=False)
        return 0

    detail_df = pd.DataFrame(detail_rows).fillna("").astype(object)
    detail_df["标的物ID"] = detail_df["标的物ID"].map(normalize_id)
    detail_df["_detail_order"] = range(len(detail_df))
    detail_df = detail_df.sort_values("_detail_order").drop_duplicates("标的物ID", keep="last")
    detail_df = detail_df.drop(columns=["_detail_order"])

    base = index_df.copy()
    base["标的物ID"] = base["标的物ID"].map(normalize_id)
    how = "inner" if completed_only else "left"
    duplicate_detail_cols = [col for col in detail_df.columns if col in base.columns and col not in {"标的物ID"}]
    detail_df = detail_df.drop(columns=duplicate_detail_cols)
    merged = base.merge(detail_df, on="标的物ID", how=how)

    ordered_cols = list(base.columns)
    ordered_cols.extend(col for col in DETAIL_COLUMNS if col in merged.columns and col not in ordered_cols)
    ordered_cols.extend(col for col in merged.columns if col not in ordered_cols)
    merged = merged[ordered_cols]

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
    parser = argparse.ArgumentParser(description="Backfill JD Guangdong auction details through APIs.")
    parser.add_argument("--input", default=DEFAULT_INPUT)
    parser.add_argument("--output", default=DEFAULT_OUTPUT)
    parser.add_argument("--jsonl-output", default=DEFAULT_JSONL)
    parser.add_argument("--checkpoint", default=DEFAULT_CHECKPOINT)
    parser.add_argument("--ids", default="", help="Comma-separated paimai IDs or JD detail URLs.")
    parser.add_argument("--start", type=int, default=0)
    parser.add_argument("--limit", type=int, default=0)
    parser.add_argument("--workers", type=int, default=8)
    parser.add_argument("--save-every", type=int, default=200)
    parser.add_argument("--export-every", type=int, default=0, help="0 means only export at the end.")
    parser.add_argument("--retries", type=int, default=4)
    parser.add_argument("--timeout", type=int, default=25)
    parser.add_argument("--raw-text-limit", type=int, default=2000)
    parser.add_argument("--no-resume", action="store_true", help="Do not skip IDs already present in JSONL.")
    parser.add_argument("--no-finance", action="store_true")
    parser.add_argument("--no-description", action="store_true")
    parser.add_argument("--no-attachments", action="store_true")
    parser.add_argument("--with-notice", action="store_true", help="Also fetch auction notice text. Slower.")
    parser.add_argument("--include-unfinished-index", action="store_true", help="Excel output keeps all index rows.")
    parser.add_argument("--export-only", action="store_true", help="Only rebuild Excel from existing JSONL.")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    input_path = Path(args.input)
    output_path = Path(args.output)
    jsonl_path = Path(args.jsonl_output)
    checkpoint_path = Path(args.checkpoint)

    index_df = read_index(input_path)
    all_ids = ordered_unique_ids(index_df["标的物ID"].tolist())
    if args.ids:
        wanted = ordered_unique_ids(part.strip() for part in args.ids.split(",") if part.strip())
        wanted_set = set(wanted)
        all_ids = [paimai_id for paimai_id in all_ids if paimai_id in wanted_set]
        missing = [paimai_id for paimai_id in wanted if paimai_id not in set(all_ids)]
        if missing:
            print(f"ids_missing_in_index count={len(missing)} ids={','.join(missing[:10])}", flush=True)

    if args.export_only:
        rows = export_excel(index_df, jsonl_path, output_path, completed_only=not args.include_unfinished_index)
        print(f"export_done output={output_path} rows={rows}", flush=True)
        return

    selected_ids = all_ids[args.start :]
    if args.limit and args.limit > 0:
        selected_ids = selected_ids[: args.limit]

    completed_before = load_completed_ids(jsonl_path) if not args.no_resume else set()
    todo_ids = [paimai_id for paimai_id in selected_ids if paimai_id not in completed_before]
    print(
        f"detail_plan input_rows={len(index_df)} unique_ids={len(all_ids)} selected={len(selected_ids)} "
        f"completed_existing={len(completed_before)} todo={len(todo_ids)} workers={args.workers}",
        flush=True,
    )

    started_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    write_checkpoint(
        checkpoint_path,
        {
            "started_at": started_at,
            "input": str(input_path),
            "jsonl_output": str(jsonl_path),
            "excel_output": str(output_path),
            "selected": len(selected_ids),
            "completed_existing": len(completed_before),
            "completed_this_run": 0,
            "failed_this_run": 0,
            "todo": len(todo_ids),
        },
    )

    completed_this_run = 0
    failed_this_run = 0
    started = time.time()
    with ThreadPoolExecutor(max_workers=max(1, args.workers)) as executor:
        future_map = {
            executor.submit(
                fetch_detail,
                paimai_id,
                not args.no_finance,
                args.with_notice,
                not args.no_attachments,
                not args.no_description,
                args.raw_text_limit,
                args.retries,
                args.timeout,
            ): paimai_id
            for paimai_id in todo_ids
        }
        for future in as_completed(future_map):
            paimai_id = future_map[future]
            try:
                row = future.result()
            except Exception as exc:  # noqa: BLE001
                failed_this_run += 1
                row = {
                    "标的物ID": paimai_id,
                    "链接": f"https://paimai.jd.com/{paimai_id}",
                    "详情抓取时间": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "详情抓取状态": "失败",
                    "详情接口错误": str(exc),
                }
            append_jsonl(jsonl_path, row)
            completed_this_run += 1

            if completed_this_run % max(1, args.save_every) == 0 or completed_this_run == len(todo_ids):
                elapsed = time.time() - started
                rate = completed_this_run / elapsed if elapsed > 0 else 0
                remaining = len(todo_ids) - completed_this_run
                eta = remaining / rate if rate > 0 else 0
                write_checkpoint(
                    checkpoint_path,
                    {
                        "started_at": started_at,
                        "input": str(input_path),
                        "jsonl_output": str(jsonl_path),
                        "excel_output": str(output_path),
                        "selected": len(selected_ids),
                        "completed_existing": len(completed_before),
                        "completed_this_run": completed_this_run,
                        "failed_this_run": failed_this_run,
                        "todo": len(todo_ids),
                        "remaining": remaining,
                        "elapsed_seconds": round(elapsed, 1),
                        "items_per_second": round(rate, 3),
                        "eta_seconds": round(eta, 1),
                    },
                )
                print(
                    f"detail_progress done={completed_this_run}/{len(todo_ids)} failed={failed_this_run} "
                    f"rate={rate:.2f}/s eta={eta/60:.1f}m",
                    flush=True,
                )

            if args.export_every and completed_this_run % args.export_every == 0:
                rows = export_excel(index_df, jsonl_path, output_path, completed_only=not args.include_unfinished_index)
                print(f"export_checkpoint output={output_path} rows={rows}", flush=True)

    rows = export_excel(index_df, jsonl_path, output_path, completed_only=not args.include_unfinished_index)
    print(f"detail_done output={output_path} rows={rows} jsonl={jsonl_path}", flush=True)


if __name__ == "__main__":
    main()
