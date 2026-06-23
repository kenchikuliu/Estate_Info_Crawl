# -*- coding: utf-8 -*-
"""
JD auction index crawler based on the public list API.

This is intentionally limited to list/index data. Detail-page fields and POI
enrichment should be backfilled from the deduplicated links after this stage.
"""
from __future__ import annotations

import argparse
import json
import math
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

import pandas as pd
import requests


API_URL = "https://api.m.jd.com/api"
REFERER = "https://pmsearch.jd.com/?publishSource=7&childrenCateId=12728"
REAL_ESTATE_CATE_ID = "12728"

PROVINCE_PRESETS: Dict[str, Dict[str, Any]] = {
    "gd": {
        "province_id": "19",
        "province_label": "广东",
        "cities": [
            ("1601", "广州市"),
            ("1607", "深圳市"),
            ("1609", "珠海市"),
            ("1611", "汕头市"),
            ("1617", "韶关市"),
            ("1627", "河源市"),
            ("1634", "梅州市"),
            ("1643", "惠州市"),
            ("1650", "汕尾市"),
            ("1655", "东莞市"),
            ("1657", "中山市"),
            ("1659", "江门市"),
            ("1666", "佛山市"),
            ("1672", "阳江市"),
            ("1677", "湛江市"),
            ("1684", "茂名市"),
            ("1690", "肇庆市"),
            ("1698", "云浮市"),
            ("1704", "清远市"),
            ("1705", "潮州市"),
            ("1709", "揭阳市"),
        ],
    },
    "bj": {
        "province_id": "1",
        "province_label": "北京",
        "cities": [("", "北京市")],
    },
}


@dataclass(frozen=True)
class DateRange:
    start: date
    end: date

    @property
    def key(self) -> str:
        return f"{self.start.isoformat()}_{self.end.isoformat()}"

    @property
    def label(self) -> str:
        return f"{self.start.isoformat()}~{self.end.isoformat()}"


@dataclass
class Bucket:
    city_id: str
    city_name: str
    range: DateRange
    total: int
    first_page: List[Dict[str, Any]]

    @property
    def key(self) -> str:
        return f"{self.city_id or self.city_name}|{self.range.key}"


def build_payload(province_id: str, city_id: str, page: int, page_size: int, drange: DateRange) -> Dict[str, Any]:
    return {
        "investmentType": "",
        "apiType": 12,
        "page": page,
        "pageSize": page_size,
        "keyword": "",
        "provinceId": str(province_id),
        "cityId": str(city_id),
        "countyId": "",
        "multiPaimaiStatus": "",
        "multiDisplayStatus": "",
        "multiPaimaiTimes": "",
        "childrenCateId": REAL_ESTATE_CATE_ID,
        "currentPriceRangeStart": "",
        "currentPriceRangeEnd": "",
        "timeRangeTime": "endTime",
        "timeRangeStart": drange.start.isoformat(),
        "timeRangeEnd": drange.end.isoformat(),
        "loan": "",
        "purchaseRestriction": "",
        "liupaiBuyAgain": "",
        "orgId": "",
        "orgType": "",
        "sortField": 8,
        "projectType": 1,
        "reqSource": 0,
        "labelSet": "",
        "publishSource": "7",
    }


def make_session() -> requests.Session:
    session = requests.Session()
    session.headers.update(
        {
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/148.0.0.0 Safari/537.36"
            ),
            "Referer": REFERER,
            "Origin": "https://pmsearch.jd.com",
            "Accept": "application/json,text/plain,*/*",
        }
    )
    return session


def fetch_search_page(
    province_id: str,
    city_id: str,
    page: int,
    page_size: int,
    drange: DateRange,
    retries: int = 4,
    sleep_seconds: float = 0.5,
) -> Dict[str, Any]:
    payload = build_payload(province_id, city_id, page, page_size, drange)
    params = {
        "appid": "paimai",
        "functionId": "paimai_unifiedSearch",
        "body": json.dumps(payload, ensure_ascii=False, separators=(",", ":")),
    }
    last_error: Optional[Exception] = None
    for attempt in range(1, retries + 1):
        try:
            session = make_session()
            response = session.post(
                API_URL,
                params=params,
                data="null",
                headers={"Content-Type": "application/json;charset=UTF-8"},
                timeout=30,
            )
            response.raise_for_status()
            data = response.json()
            if data.get("statusCode") not in (None, 200) and data.get("code") not in (0, "0"):
                raise RuntimeError(f"JD API returned status={data.get('statusCode')} code={data.get('code')}")
            return data
        except Exception as exc:  # noqa: BLE001 - keep crawler resilient.
            last_error = exc
            time.sleep(sleep_seconds * attempt)
    raise RuntimeError(
        f"Failed fetching province={province_id} city={city_id} page={page} range={drange.label}: {last_error}"
    )


def fetch_config(ids: Sequence[Any], retries: int = 3) -> Dict[str, Dict[str, Any]]:
    if not ids:
        return {}
    result: Dict[str, Dict[str, Any]] = {}
    for i in range(0, len(ids), 200):
        chunk = [str(x) for x in ids[i : i + 200] if str(x).strip()]
        if not chunk:
            continue
        body = {
            "paimaiIds": "-".join(chunk),
            "configEnums": "PAIMAI_INSURANCE_DATA,PAIMAI_LOAN_DATA,PAIMAI_LABEL_DATA,PAIMAI_ALLOCATION_DATA",
        }
        params = {
            "appid": "paimai",
            "functionId": "getPaimaiConfigInfo",
            "body": json.dumps(body, ensure_ascii=False, separators=(",", ":")),
        }
        last_error: Optional[Exception] = None
        for attempt in range(1, retries + 1):
            try:
                session = make_session()
                response = session.get(API_URL, params=params, timeout=30)
                response.raise_for_status()
                data = response.json().get("data") or {}
                result.update({str(k): v for k, v in data.items() if isinstance(v, dict)})
                break
            except Exception as exc:  # noqa: BLE001
                last_error = exc
                time.sleep(0.5 * attempt)
        else:
            print(f"config_failed ids={chunk[:2]}... error={last_error}", flush=True)
    return result


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
    try:
        auction = int(auction_status)
    except (TypeError, ValueError):
        auction = -1
    try:
        display = int(display_status)
    except (TypeError, ValueError):
        display = -1

    display_map = {3: "已暂停", 5: "已撤回", 6: "已暂缓", 7: "已中止"}
    if display in display_map:
        return display_map[display]
    return {0: "预告中", 1: "进行中", 2: "已结束"}.get(auction, "")


def label_text(config: Dict[str, Any], item: Dict[str, Any]) -> str:
    labels = []
    for label in config.get("paimaiLabelConfigList") or []:
        name = label.get("labelName")
        if name:
            labels.append(str(name))
    if labels:
        return "、".join(dict.fromkeys(labels))
    raw_labels = item.get("labelSet")
    if isinstance(raw_labels, list):
        return ",".join(str(x) for x in raw_labels)
    return ""


def map_row(
    item: Dict[str, Any],
    province_label: str,
    city_id: str,
    city_name: str,
    drange: DateRange,
    page_no: int,
    config: Optional[Dict[str, Any]] = None,
) -> Dict[str, Any]:
    config = config or {}
    asset_id = item.get("id", "")
    labels = item.get("labelSet") if isinstance(item.get("labelSet"), list) else []
    has_loan_label = 1057 in labels or "1057" in {str(x) for x in labels}
    is_support_loan = config.get("isSupport")
    has_finance = bool(is_support_loan or has_loan_label or item.get("loan") == 1)
    title = item.get("title", "")
    current_price = item.get("currentPriceCN") or item.get("currentPriceWithUnit") or item.get("currentPriceStr") or ""
    assessment_price = item.get("assessmentPriceCN") or item.get("assessmentPriceStr") or ""
    market_price = item.get("marketPriceCN") or item.get("marketPriceStr") or ""
    row = {
        "标的物ID": asset_id,
        "链接": f"https://paimai.jd.com/{asset_id}" if asset_id else "",
        "列表页码": page_no,
        "省份筛选": province_label,
        "城市筛选": city_name,
        "城市ID": city_id,
        "时间桶开始": drange.start.isoformat(),
        "时间桶结束": drange.end.isoformat(),
        "抓取时间": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "标题": title,
        "城市": item.get("city") or item.get("courtCityName") or city_name,
        "地址": item.get("productAddress", ""),
        "处置机构": item.get("shopName", ""),
        "当前价": current_price,
        "当前价_元": item.get("currentPrice", ""),
        "起拍价_元": item.get("minPrice", ""),
        "评估价": assessment_price,
        "评估价_元": item.get("assessmentPrice", ""),
        "市场价": market_price,
        "市场价_元": item.get("marketPrice", ""),
        "保证金_元": item.get("ensurePrice", ""),
        "出价次数": item.get("bidCount", ""),
        "围观次数": item.get("accessNumber", item.get("uvCount", "")),
        "竞价状态": status_text(item.get("auctionStatus"), item.get("displayStatus")),
        "auctionStatus": item.get("auctionStatus", ""),
        "displayStatus": item.get("displayStatus", ""),
        "开始时间": ts_to_text(item.get("startTimeMills", item.get("startTime"))),
        "结束时间": ts_to_text(item.get("endTimeMills", item.get("endTime"))),
        "startTime": item.get("startTimeMills", item.get("startTime", "")),
        "endTime": item.get("endTimeMills", item.get("endTime", "")),
        "是否有金融服务": "是" if has_finance else "否",
        "是否支持贷款": "" if is_support_loan is None else ("是" if is_support_loan else "否"),
        "是否支持保险": "" if config.get("isSupportInsurance") is None else ("是" if config.get("isSupportInsurance") else "否"),
        "是否配资服务": "" if config.get("whetherAllocation") is None else ("是" if config.get("whetherAllocation") else "否"),
        "标签": label_text(config, item),
        "labelSet": json.dumps(labels, ensure_ascii=False) if labels else "",
        "loan": item.get("loan", ""),
        "purchaseRestriction": item.get("purchaseRestriction", ""),
        "paimaiTimes": item.get("paimaiTimes", ""),
        "skuId": item.get("skuId", ""),
        "productId": item.get("productId", ""),
        "shopId": item.get("shopId", ""),
        "vendorId": item.get("vendorId", ""),
        "publishSource": item.get("publishSource", ""),
        "productImage": item.get("productImage", ""),
    }
    row["列表原文"] = "\n".join(
        str(x)
        for x in [
            title,
            row["城市"],
            f"当前价:{current_price}" if current_price else "",
            f"评估价:{assessment_price}" if assessment_price else "",
            f"市场价:{market_price}" if market_price else "",
            f"{row['出价次数']}次出价" if row["出价次数"] != "" else "",
            row["竞价状态"],
            row["结束时间"],
        ]
        if x
    )
    return row


def split_range(drange: DateRange) -> Tuple[DateRange, DateRange]:
    days = (drange.end - drange.start).days
    mid = drange.start + timedelta(days=days // 2)
    return DateRange(drange.start, mid), DateRange(mid + timedelta(days=1), drange.end)


def initial_ranges(start_year: int, end_year: int) -> List[DateRange]:
    ranges: List[DateRange] = []
    if start_year < 2020:
        ranges.append(DateRange(date(start_year, 1, 1), date(2019, 12, 31)))
        first_year = 2020
    else:
        first_year = start_year
    current_year = datetime.now().year
    annual_until = min(end_year, current_year)
    for year in range(first_year, annual_until + 1):
        ranges.append(DateRange(date(year, 1, 1), date(year, 12, 31)))
    if end_year > annual_until:
        ranges.append(DateRange(date(annual_until + 1, 1, 1), date(end_year, 12, 31)))
    return ranges


def resolve_buckets(
    province_id: str,
    city_id: str,
    city_name: str,
    ranges: Sequence[DateRange],
    page_size: int,
    max_bucket_size: int,
    workers: int,
) -> List[Bucket]:
    resolved: List[Bucket] = []
    pending = list(ranges)
    while pending:
        next_pending: List[DateRange] = []
        with ThreadPoolExecutor(max_workers=max(1, workers)) as executor:
            future_map = {
                executor.submit(fetch_search_page, province_id, city_id, 1, page_size, drange): drange
                for drange in pending
            }
            for future in as_completed(future_map):
                drange = future_map[future]
                data = future.result()
                total = int(data.get("totalItem") or 0)
                first_page = data.get("datas") or []
                if total > max_bucket_size and drange.start < drange.end:
                    left, right = split_range(drange)
                    print(f"split city={city_name} range={drange.label} total={total}", flush=True)
                    next_pending.extend([left, right])
                else:
                    resolved.append(Bucket(city_id, city_name, drange, total, first_page))
        pending = next_pending
    return sorted(resolved, key=lambda b: (b.range.start, b.range.end))


def load_checkpoint(path: Path) -> Dict[str, Any]:
    if not path.exists():
        return {"completed_buckets": []}
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return {"completed_buckets": []}


def write_checkpoint(path: Path, checkpoint: Dict[str, Any]) -> None:
    checkpoint["updated_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    path.write_text(json.dumps(checkpoint, ensure_ascii=False, indent=2), encoding="utf-8")


def load_existing(path: Path) -> List[Dict[str, Any]]:
    if not path.exists():
        return []
    return pd.read_excel(path).fillna("").astype(object).to_dict("records")


def save_rows(rows: List[Dict[str, Any]], path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    if not rows:
        with pd.ExcelWriter(
            path,
            engine="xlsxwriter",
            engine_kwargs={"options": {"strings_to_urls": False}},
        ) as writer:
            pd.DataFrame().to_excel(writer, index=False)
        return
    df = pd.DataFrame(rows)
    if "链接" in df.columns:
        df["_non_empty"] = df.apply(lambda row: sum(str(v).strip() != "" for v in row), axis=1)
        df = df.sort_values(["链接", "_non_empty"], ascending=[True, False])
        df = df.drop_duplicates("链接", keep="first").drop(columns=["_non_empty"])
    sort_cols = [col for col in ["城市筛选", "endTime", "标的物ID"] if col in df.columns]
    if sort_cols:
        df = df.sort_values(sort_cols, kind="stable")
    tmp_path = path.with_suffix(path.suffix + ".tmp.xlsx")
    with pd.ExcelWriter(
        tmp_path,
        engine="xlsxwriter",
        engine_kwargs={"options": {"strings_to_urls": False}},
    ) as writer:
        df.to_excel(writer, index=False)
    tmp_path.replace(path)


def crawl_bucket(
    bucket: Bucket,
    province_id: str,
    province_label: str,
    page_size: int,
    workers: int,
    with_config: bool,
) -> Tuple[List[Dict[str, Any]], Dict[str, Any]]:
    pages = int(math.ceil(bucket.total / page_size)) if bucket.total else 0
    page_data: Dict[int, List[Dict[str, Any]]] = {1: bucket.first_page}
    if pages > 1:
        with ThreadPoolExecutor(max_workers=max(1, workers)) as executor:
            future_map = {
                executor.submit(fetch_search_page, province_id, bucket.city_id, page, page_size, bucket.range): page
                for page in range(2, pages + 1)
            }
            for future in as_completed(future_map):
                page = future_map[future]
                data = future.result()
                page_data[page] = data.get("datas") or []
    rows: List[Dict[str, Any]] = []
    for page in range(1, pages + 1):
        items = page_data.get(page, [])
        configs = fetch_config([item.get("id") for item in items]) if with_config else {}
        rows.extend(
            map_row(
                item,
                province_label,
                bucket.city_id,
                bucket.city_name,
                bucket.range,
                page,
                configs.get(str(item.get("id"))) if configs else None,
            )
            for item in items
        )
    unique_links = len({row.get("链接") for row in rows if row.get("链接")})
    stat = {
        "城市": bucket.city_name,
        "城市ID": bucket.city_id,
        "时间桶": bucket.range.label,
        "接口总量": bucket.total,
        "抓取行数": len(rows),
        "唯一链接数": unique_links,
        "页数": pages,
    }
    return rows, stat


def selected_cities(city_arg: str, cities: Sequence[Tuple[str, str]]) -> List[Tuple[str, str]]:
    if not city_arg or city_arg.lower() == "all":
        return list(cities)
    wanted = {x.strip() for x in city_arg.split(",") if x.strip()}
    result = [(cid, name) for cid, name in cities if cid in wanted or name in wanted]
    missing = wanted - {cid for cid, _ in result} - {name for _, name in result}
    if missing:
        raise ValueError(f"Unknown city selector(s): {', '.join(sorted(missing))}")
    return result


def resolve_province(args: argparse.Namespace) -> Tuple[str, str, List[Tuple[str, str]]]:
    province_key = (args.province_key or "").strip().lower()
    preset = PROVINCE_PRESETS.get(province_key, {})
    province_id = str(args.province_id or preset.get("province_id") or "").strip()
    province_label = str(args.province_label or preset.get("province_label") or province_key or province_id).strip()
    cities = preset.get("cities") or [("", f"{province_label}市" if province_label and not province_label.endswith("市") else province_label)]
    if not province_id or not province_label:
        raise ValueError("Missing province config. Provide --province-key or both --province-id and --province-label.")
    return province_id, province_label, list(cities)


def default_paths(province_label: str) -> Dict[str, str]:
    return {
        "checkpoint": rf"output\京东法拍房_{province_label}_API城市索引.checkpoint.json",
        "merge_output": rf"output\京东法拍房_{province_label}_全量索引_API合并.xlsx",
        "stats_output": rf"output\京东法拍房_{province_label}_API城市索引_统计.xlsx",
    }


def merge_city_files(files: Iterable[Path], output_path: Path) -> int:
    frames = []
    for file in files:
        if file.exists() and file.stat().st_size > 0:
            frames.append(pd.read_excel(file).fillna("").astype(object))
    if not frames:
        with pd.ExcelWriter(
            output_path,
            engine="xlsxwriter",
            engine_kwargs={"options": {"strings_to_urls": False}},
        ) as writer:
            pd.DataFrame().to_excel(writer, index=False)
        return 0
    df = pd.concat(frames, ignore_index=True)
    if "链接" in df.columns:
        df["_non_empty"] = df.apply(lambda row: sum(str(v).strip() != "" for v in row), axis=1)
        df = df.sort_values(["链接", "_non_empty"], ascending=[True, False])
        df = df.drop_duplicates("链接", keep="first").drop(columns=["_non_empty"])
    sort_cols = [col for col in ["城市筛选", "endTime", "标的物ID"] if col in df.columns]
    if sort_cols:
        df = df.sort_values(sort_cols, kind="stable")
    output_path.parent.mkdir(parents=True, exist_ok=True)
    tmp_path = output_path.with_suffix(output_path.suffix + ".tmp.xlsx")
    with pd.ExcelWriter(
        tmp_path,
        engine="xlsxwriter",
        engine_kwargs={"options": {"strings_to_urls": False}},
    ) as writer:
        df.to_excel(writer, index=False)
    tmp_path.replace(output_path)
    return len(df)


def main() -> None:
    parser = argparse.ArgumentParser(description="Crawl JD auction list data through the public list API.")
    parser.add_argument("--province-key", default="gd", help="Province preset key, e.g. gd or bj.")
    parser.add_argument("--province-id", default="", help="Override JD provinceId.")
    parser.add_argument("--province-label", default="", help="Override human-readable province label.")
    parser.add_argument("--cities", default="all", help="Comma-separated city names/ids, or all.")
    parser.add_argument("--output-dir", default=r"output")
    parser.add_argument("--checkpoint", default="", help="Checkpoint JSON path. Defaults by province.")
    parser.add_argument("--merge-output", default="", help="Merged Excel path. Defaults by province.")
    parser.add_argument("--stats-output", default="", help="Stats Excel path. Defaults by province.")
    parser.add_argument("--page-size", type=int, default=1000)
    parser.add_argument("--max-bucket-size", type=int, default=9000)
    parser.add_argument("--workers", type=int, default=4)
    parser.add_argument("--start-year", type=int, default=2000)
    parser.add_argument("--end-year", type=int, default=2035)
    parser.add_argument("--with-config", action="store_true", help="Also fetch loan/label config for each asset.")
    args = parser.parse_args()

    province_id, province_label, province_cities = resolve_province(args)
    path_defaults = default_paths(province_label)
    output_dir = Path(args.output_dir)
    checkpoint_path = Path(args.checkpoint or path_defaults["checkpoint"])
    merge_output = Path(args.merge_output or path_defaults["merge_output"])
    stats_output = Path(args.stats_output or path_defaults["stats_output"])
    checkpoint = load_checkpoint(checkpoint_path)
    completed = set(checkpoint.get("completed_buckets") or [])
    all_stats: List[Dict[str, Any]] = checkpoint.get("stats") or []
    city_files: List[Path] = []
    ranges = initial_ranges(args.start_year, args.end_year)

    for city_id, city_name in selected_cities(args.cities, province_cities):
        city_output = output_dir / f"京东法拍房_{province_label}_API城市索引_{city_name}.xlsx"
        city_files.append(city_output)
        rows = load_existing(city_output)
        before_unique = len({row.get("链接") for row in rows if row.get("链接")})
        print(f"city_start {city_name} existing_unique={before_unique}", flush=True)
        buckets = resolve_buckets(
            province_id,
            city_id,
            city_name,
            ranges,
            args.page_size,
            args.max_bucket_size,
            args.workers,
        )
        total_by_buckets = sum(bucket.total for bucket in buckets)
        city_stat_header = {
            "城市": city_name,
            "城市ID": city_id,
            "时间桶": "ALL_BUCKETS",
            "接口总量": total_by_buckets,
            "抓取行数": "",
            "唯一链接数": "",
            "页数": "",
        }
        all_stats.append(city_stat_header)
        print(f"city_plan {city_name} buckets={len(buckets)} bucket_total={total_by_buckets}", flush=True)
        for bucket in buckets:
            if bucket.total == 0:
                continue
            if bucket.key in completed:
                print(f"bucket_skip {city_name} {bucket.range.label} total={bucket.total}", flush=True)
                continue
            started = time.time()
            bucket_rows, stat = crawl_bucket(
                bucket,
                province_id,
                province_label,
                args.page_size,
                args.workers,
                args.with_config,
            )
            rows.extend(bucket_rows)
            save_rows(rows, city_output)
            completed.add(bucket.key)
            all_stats.append(stat)
            checkpoint["completed_buckets"] = sorted(completed)
            checkpoint["stats"] = all_stats
            write_checkpoint(checkpoint_path, checkpoint)
            elapsed = time.time() - started
            city_unique = len({row.get("链接") for row in rows if row.get("链接")})
            print(
                f"bucket_done {city_name} {bucket.range.label} total={bucket.total} "
                f"rows={len(bucket_rows)} city_unique={city_unique} elapsed={elapsed:.1f}s",
                flush=True,
            )

    stats_output.parent.mkdir(parents=True, exist_ok=True)
    if all_stats:
        pd.DataFrame(all_stats).to_excel(stats_output, index=False)
    total_unique = merge_city_files(city_files, merge_output)
    print(f"merge_done output={merge_output} unique={total_unique}", flush=True)


if __name__ == "__main__":
    main()
