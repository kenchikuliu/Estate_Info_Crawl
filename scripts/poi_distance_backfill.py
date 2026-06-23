# -*- coding: utf-8 -*-
"""
Backfill surrounding POI distance fields with coordinate-level caching.

The public Overpass/OSM source can be slow or sparse in Mainland China, so this
script is deliberately resumable and sample-friendly. It processes unique
coordinates first, then joins the cached distances back to the full Excel.
"""
from __future__ import annotations

import argparse
import json
import math
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

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from utils.jd_detail_parser import (  # noqa: E402
    POI_COLUMN_SPECS,
    build_overpass_query,
    clean_text,
    format_distance,
    haversine_distance_meters,
)


DEFAULT_INPUT = r"output\京东法拍房_广东_详情回填_API_含附件索引.xlsx"
DEFAULT_FALLBACK_INPUT = r"output\京东法拍房_广东_详情回填_API.xlsx"
DEFAULT_OUTPUT = r"output\京东法拍房_广东_详情回填_API_含附件周边.xlsx"
DEFAULT_CACHE = r"output\京东法拍房_广东_POI周边距离缓存.jsonl"
DEFAULT_CHECKPOINT = r"output\京东法拍房_广东_POI周边距离.checkpoint.json"

OVERPASS_ENDPOINTS = [
    "https://overpass-api.de/api/interpreter",
    "https://overpass.kumi.systems/api/interpreter",
    "https://overpass.openstreetmap.ru/api/interpreter",
]

thread_local = threading.local()


def get_session() -> requests.Session:
    session = getattr(thread_local, "session", None)
    if session is None:
        session = requests.Session()
        session.headers.update({"User-Agent": "EstateInfoCrawl/1.0 (poi enrichment)"})
        thread_local.session = session
    return session


def clean_number(value: Any) -> Optional[float]:
    if value in ("", None):
        return None
    text = str(value).strip()
    if not text or text.lower() == "nan":
        return None
    try:
        number = float(text)
    except Exception:
        return None
    if not math.isfinite(number):
        return None
    return number


def coord_key(lat: float, lon: float, precision: int = 6) -> str:
    return f"{round(lat, precision):.{precision}f},{round(lon, precision):.{precision}f}"


def valid_guangdong_coord(lat: float, lon: float) -> bool:
    return 20.0 <= lat <= 25.8 and 109.0 <= lon <= 118.0


def find_column(columns: Sequence[str], candidates: Sequence[str]) -> str:
    for candidate in candidates:
        if candidate in columns:
            return candidate
    return ""


def read_coordinate_rows(path: Path, precision: int) -> pd.DataFrame:
    df = pd.read_excel(path).fillna("").astype(object)
    lat_col = find_column(df.columns, ["纬度", "lat", "latitude"])
    lon_col = find_column(df.columns, ["经度", "lng", "lon", "longitude"])
    if not lat_col or not lon_col:
        raise ValueError("Input Excel must contain 经度/纬度 columns.")

    keys: List[str] = []
    valid_mask: List[bool] = []
    for _, row in df.iterrows():
        lat = clean_number(row.get(lat_col))
        lon = clean_number(row.get(lon_col))
        valid = lat is not None and lon is not None and valid_guangdong_coord(lat, lon)
        valid_mask.append(valid)
        keys.append(coord_key(lat, lon, precision) if valid and lat is not None and lon is not None else "")
    df["_poi_coord_key"] = keys
    df["_poi_coord_valid"] = valid_mask
    return df


def unique_coords(df: pd.DataFrame) -> List[Tuple[str, float, float, int]]:
    counts = df.loc[df["_poi_coord_valid"], "_poi_coord_key"].value_counts()
    result: List[Tuple[str, float, float, int]] = []
    for key, count in counts.items():
        lat_text, lon_text = str(key).split(",", 1)
        result.append((str(key), float(lat_text), float(lon_text), int(count)))
    return result


def read_cache(path: Path) -> Dict[str, Dict[str, Any]]:
    if not path.exists():
        return {}
    result: Dict[str, Dict[str, Any]] = {}
    with path.open("r", encoding="utf-8") as handle:
        for line in handle:
            line = line.strip()
            if not line:
                continue
            try:
                item = json.loads(line)
            except Exception:
                continue
            if isinstance(item, dict) and item.get("坐标键"):
                result[str(item["坐标键"])] = item
    return result


def append_cache(path: Path, row: Dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("a", encoding="utf-8") as handle:
        handle.write(json.dumps(row, ensure_ascii=False, separators=(",", ":")) + "\n")


def write_checkpoint(path: Path, payload: Dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    payload = dict(payload)
    payload["updated_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def fetch_overpass_payload(lat: float, lon: float, timeout: int, endpoint_index: int) -> Tuple[List[Dict[str, Any]], str]:
    max_radius = max(int(spec["radius"]) for spec in POI_COLUMN_SPECS.values())
    tag_pairs = set()
    for spec in POI_COLUMN_SPECS.values():
        for tag in spec["tags"]:
            tag_pairs.add(tag)
    query = build_overpass_query(lat, lon, max_radius, list(tag_pairs))
    errors: List[str] = []
    endpoints = OVERPASS_ENDPOINTS[endpoint_index:] + OVERPASS_ENDPOINTS[:endpoint_index]
    for endpoint in endpoints:
        try:
            response = get_session().post(endpoint, data=query.encode("utf-8"), timeout=timeout)
            response.raise_for_status()
            payload = response.json()
            elements = payload.get("elements", [])
            return elements if isinstance(elements, list) else [], endpoint
        except Exception as exc:  # noqa: BLE001
            errors.append(f"{endpoint}:{exc}")
    raise RuntimeError(" | ".join(errors))


def element_coord(element: Dict[str, Any]) -> Tuple[Optional[float], Optional[float]]:
    lat = element.get("lat")
    lon = element.get("lon")
    center = element.get("center") or {}
    if lat is None:
        lat = center.get("lat")
    if lon is None:
        lon = center.get("lon")
    try:
        return float(lat), float(lon)
    except Exception:
        return None, None


def tag_matches(tags: Dict[str, Any], expected: Tuple[str, str]) -> bool:
    key, value = expected
    actual = tags.get(key)
    if actual == value:
        return True
    if value == "yes" and actual in {True, "true", "1", 1, "yes"}:
        return True
    return False


def poi_name(tags: Dict[str, Any]) -> str:
    return clean_text(str(tags.get("name:zh") or tags.get("name") or tags.get("name:en") or ""))


def compute_distances(lat: float, lon: float, elements: Sequence[Dict[str, Any]]) -> Dict[str, str]:
    min_distances: Dict[str, Optional[float]] = {column: None for column in POI_COLUMN_SPECS}
    for element in elements:
        if not isinstance(element, dict):
            continue
        tags = element.get("tags") or {}
        if not isinstance(tags, dict):
            continue
        poi_lat, poi_lon = element_coord(element)
        if poi_lat is None or poi_lon is None:
            continue
        distance = haversine_distance_meters(lat, lon, poi_lat, poi_lon)
        name = poi_name(tags)
        for column, spec in POI_COLUMN_SPECS.items():
            if distance > float(spec["radius"]):
                continue
            if not any(tag_matches(tags, tuple(tag)) for tag in spec["tags"]):
                continue
            keywords = [str(keyword) for keyword in spec.get("name_keywords", [])]
            if keywords and name and not any(keyword in name for keyword in keywords):
                continue
            current = min_distances[column]
            if current is None or distance < current:
                min_distances[column] = distance

    return {
        column: (format_distance(distance) if distance is not None else "")
        for column, distance in min_distances.items()
    }


def fetch_coord_poi(
    key: str,
    lat: float,
    lon: float,
    count: int,
    timeout: int,
    retries: int,
    sleep_seconds: float,
) -> Dict[str, Any]:
    last_error = ""
    for attempt in range(1, retries + 1):
        try:
            elements, endpoint = fetch_overpass_payload(lat, lon, timeout, attempt % len(OVERPASS_ENDPOINTS))
            distances = compute_distances(lat, lon, elements)
            non_empty = sum(1 for value in distances.values() if value)
            return {
                "坐标键": key,
                "纬度": lat,
                "经度": lon,
                "房源条数": count,
                "POI抓取时间": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "POI抓取状态": "成功",
                "POI抓取错误": "",
                "POI数据源": endpoint,
                "POI原始元素数": len(elements),
                "POI非空字段数": non_empty,
                **distances,
            }
        except Exception as exc:  # noqa: BLE001
            last_error = str(exc)
            time.sleep(sleep_seconds * attempt)
    return {
        "坐标键": key,
        "纬度": lat,
        "经度": lon,
        "房源条数": count,
        "POI抓取时间": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "POI抓取状态": "失败",
        "POI抓取错误": last_error,
        "POI数据源": "",
        "POI原始元素数": 0,
        "POI非空字段数": 0,
        **{column: "" for column in POI_COLUMN_SPECS},
    }


def export_excel(input_path: Path, cache_path: Path, output_path: Path, precision: int) -> int:
    df = read_coordinate_rows(input_path, precision)
    cache = read_cache(cache_path)
    for column in [
        "POI抓取状态", "POI抓取错误", "POI数据源", "POI原始元素数", "POI非空字段数",
        *POI_COLUMN_SPECS.keys(),
    ]:
        if column not in df.columns:
            df[column] = ""

    for row_index, key in df["_poi_coord_key"].items():
        if not key or key not in cache:
            continue
        item = cache[key]
        for column in ["POI抓取状态", "POI抓取错误", "POI数据源", "POI原始元素数", "POI非空字段数", *POI_COLUMN_SPECS.keys()]:
            df.at[row_index, column] = item.get(column, "")

    df = df.drop(columns=["_poi_coord_key", "_poi_coord_valid"], errors="ignore")
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


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Backfill surrounding POI distances with coordinate cache.")
    parser.add_argument("--input", default=DEFAULT_INPUT)
    parser.add_argument("--fallback-input", default=DEFAULT_FALLBACK_INPUT)
    parser.add_argument("--output", default=DEFAULT_OUTPUT)
    parser.add_argument("--cache", default=DEFAULT_CACHE)
    parser.add_argument("--checkpoint", default=DEFAULT_CHECKPOINT)
    parser.add_argument("--workers", type=int, default=2)
    parser.add_argument("--save-every", type=int, default=50)
    parser.add_argument("--start", type=int, default=0)
    parser.add_argument("--limit", type=int, default=0)
    parser.add_argument("--precision", type=int, default=6)
    parser.add_argument("--timeout", type=int, default=40)
    parser.add_argument("--retries", type=int, default=2)
    parser.add_argument("--sleep-seconds", type=float, default=0.8)
    parser.add_argument("--no-resume", action="store_true")
    parser.add_argument("--export-only", action="store_true")
    parser.add_argument("--no-export", action="store_true")
    return parser.parse_args()


def resolve_input(path: str, fallback: str) -> Path:
    primary = Path(path)
    if primary.exists():
        return primary
    fallback_path = Path(fallback)
    if fallback_path.exists():
        return fallback_path
    raise FileNotFoundError(path)


def main() -> None:
    args = parse_args()
    input_path = resolve_input(args.input, args.fallback_input)
    output_path = Path(args.output)
    cache_path = Path(args.cache)
    checkpoint_path = Path(args.checkpoint)

    if args.export_only:
        rows = export_excel(input_path, cache_path, output_path, args.precision)
        print(f"poi_export_done output={output_path} rows={rows}", flush=True)
        return

    df = read_coordinate_rows(input_path, args.precision)
    coords = unique_coords(df)
    selected = coords[args.start :]
    if args.limit and args.limit > 0:
        selected = selected[: args.limit]

    cache = {} if args.no_resume else read_cache(cache_path)
    todo = [item for item in selected if item[0] not in cache or cache[item[0]].get("POI抓取状态") == "失败"]
    valid_rows = int(df["_poi_coord_valid"].sum())
    print(
        f"poi_plan input={input_path} rows={len(df)} valid_coord_rows={valid_rows} "
        f"unique_coords={len(coords)} selected={len(selected)} completed_existing={len(cache)} "
        f"todo={len(todo)} workers={args.workers}",
        flush=True,
    )

    started_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    started = time.time()
    done = 0
    failed = 0
    non_empty = 0
    with ThreadPoolExecutor(max_workers=max(1, args.workers)) as executor:
        future_map = {
            executor.submit(
                fetch_coord_poi,
                key,
                lat,
                lon,
                count,
                args.timeout,
                args.retries,
                args.sleep_seconds,
            ): key
            for key, lat, lon, count in todo
        }
        for future in as_completed(future_map):
            row = future.result()
            append_cache(cache_path, row)
            done += 1
            failed += row.get("POI抓取状态") == "失败"
            non_empty += int(row.get("POI非空字段数") or 0) > 0
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
                        "cache": str(cache_path),
                        "output": str(output_path),
                        "rows": len(df),
                        "valid_coord_rows": valid_rows,
                        "unique_coords": len(coords),
                        "selected": len(selected),
                        "completed_existing": len(cache),
                        "completed_this_run": done,
                        "failed_this_run": failed,
                        "non_empty_this_run": non_empty,
                        "todo": len(todo),
                        "remaining": remaining,
                        "elapsed_seconds": round(elapsed, 1),
                        "items_per_second": round(rate, 3),
                        "eta_seconds": round(eta, 1),
                    },
                )
                print(
                    f"poi_progress done={done}/{len(todo)} failed={failed} "
                    f"non_empty={non_empty} rate={rate:.2f}/s eta={eta/60:.1f}m",
                    flush=True,
                )

    if not args.no_export:
        rows = export_excel(input_path, cache_path, output_path, args.precision)
        print(f"poi_done output={output_path} rows={rows} cache={cache_path}", flush=True)
    else:
        print(f"poi_cache_done cache={cache_path}", flush=True)


if __name__ == "__main__":
    main()
