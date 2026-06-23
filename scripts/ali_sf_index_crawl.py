# -*- coding: utf-8 -*-
"""
Alibaba judicial-auction real-estate index crawler based on the public H5 API.

Why this version exists:
1. The legacy HTML list page is now unstable and often redirects to login/captcha.
2. The mobile H5 mtop endpoint is significantly more stable for list/index data.
3. We keep the output surface close to the previous Guangdong workflow so the
   downstream Ali detail backfill can continue to read the same Excel outputs.

Current scope:
- Guangdong judicial-auction real-estate index data
- Stable list/index fields
- Optional public bid-detail enrichment for ended / paused / revoked assets

Out of scope:
- Full protected detail-page fields such as notice / warrant / owner text
  still belong to the browser-driven detail backfill step.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import math
import re
import threading
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass, replace
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

import pandas as pd
import requests


PROJECT_ROOT = Path(__file__).resolve().parents[1]
H5_GATEWAY = "https://h5api.m.taobao.com"
H5_APPKEY = "12574478"
HOME_URL = "https://sf.taobao.com/"
PAGE_SIZE = 10
ALI_H5_PAGE_CAP = 500
DEFAULT_SORT = "1"
DEFAULT_STATUS_ORDERS = ["0", "1", "2", "4", "5"]
DEFAULT_CIRC_GROUPS: List[Tuple[Tuple[str, ...], str]] = [
    (("1",), "一拍"),
    (("2",), "二拍"),
    (("4",), "变卖"),
    (("5",), "其他轮次5"),
    (("-1", "-2", "-4"), "其他轮次"),
]
DEFAULT_ZC_BIZ_TYPES = ["4", "6", "8", "10"]
PRICE_WINDOW_SORTS = ["501", "502"]
ILLEGAL_XML_RE = re.compile(r"[\x00-\x08\x0B\x0C\x0E-\x1F\x7F-\x9F]")
JSONP_RE = re.compile(r"callback\((.*)\)\s*$", re.S)
MOBILE_UA = (
    "Mozilla/5.0 (iPhone; CPU iPhone OS 17_5 like Mac OS X) "
    "AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.5 Mobile/15E148 Safari/604.1"
)

PROVINCE_PRESETS: Dict[str, Dict[str, Any]] = {
    "gd": {
        "province_label": "广东",
        "province_code": "440000",
        "cities": [
            ("440100", "广州市"),
            ("440300", "深圳市"),
            ("440400", "珠海市"),
            ("440500", "汕头市"),
            ("440200", "韶关市"),
            ("441600", "河源市"),
            ("441400", "梅州市"),
            ("441300", "惠州市"),
            ("441500", "汕尾市"),
            ("441900", "东莞市"),
            ("442000", "中山市"),
            ("440700", "江门市"),
            ("440600", "佛山市"),
            ("441700", "阳江市"),
            ("440800", "湛江市"),
            ("440900", "茂名市"),
            ("441200", "肇庆市"),
            ("441800", "清远市"),
            ("445100", "潮州市"),
            ("445200", "揭阳市"),
            ("445300", "云浮市"),
        ],
    },
    "bj": {
        "province_label": "北京",
        "province_code": "110000",
        "cities": [
            ("110000", "北京市"),
        ],
    },
    "sh": {
        "province_label": "上海",
        "province_code": "310000",
        "cities": [
            ("310000", "上海市"),
        ],
    },
    "zj": {
        "province_label": "浙江",
        "province_code": "330000",
        "cities": [
            ("330100", "杭州市"),
            ("330200", "宁波市"),
            ("330300", "温州市"),
            ("330400", "嘉兴市"),
            ("330500", "湖州市"),
            ("330600", "绍兴市"),
            ("330700", "金华市"),
            ("330800", "衢州市"),
            ("330900", "舟山市"),
            ("331000", "台州市"),
            ("331100", "丽水市"),
        ],
    },
    "hb": {
        "province_label": "湖北",
        "province_code": "420000",
        "cities": [
            ("420100", "武汉市"),
            ("420200", "黄石市"),
            ("420300", "十堰市"),
            ("420500", "宜昌市"),
            ("420600", "襄阳市"),
            ("420700", "鄂州市"),
            ("420800", "荆门市"),
            ("420900", "孝感市"),
            ("421000", "荆州市"),
            ("421100", "黄冈市"),
            ("421200", "咸宁市"),
            ("421300", "随州市"),
            ("422800", "恩施土家族苗族自治州"),
            ("429004", "仙桃市"),
            ("429005", "潜江市"),
            ("429006", "天门市"),
            ("429021", "神农架林区"),
        ],
    },
}

CATEGORY_PRESETS: Dict[str, Dict[str, Any]] = {
    "50025969": {"name": "住宅用房", "fcat_v4_ids": ["206060601"], "scope": "city"},
    "200782003": {"name": "商业用房", "fcat_v4_ids": ["206057102"], "scope": "city"},
    "200788003": {"name": "工业用房", "fcat_v4_ids": ["206051702"], "scope": "city"},
    "50025970": {"name": "土地", "fcat_v4_ids": ["206067101"], "scope": "province"},
    "50025975": {"name": "工程", "fcat_v4_ids": ["206146002"], "scope": "province"},
    "200798003": {"name": "其他用房", "fcat_v4_ids": ["206060701"], "scope": "city"},
}

STATUS_TEXT_MAP = {
    "before": "待开始",
    "ing": "进行中",
    "end": "已结束",
    "pause": "已中止",
    "revoke": "已撤回",
}
STATUS_ORDER_LABELS = {
    "0": "进行中",
    "1": "即将开始",
    "2": "已结束",
    "4": "已中止",
    "5": "已撤回",
}
CIRC_LABELS = {
    "1": "一拍",
    "2": "二拍",
    "4": "变卖",
    "5": "其他轮次5",
    "-1": "其他轮次-1",
    "-2": "其他轮次-2",
    "-4": "其他轮次-4",
}
ZC_BIZ_TYPE_LABELS = {
    "4": "涉刑资产",
    "6": "诉讼资产",
    "8": "破产资产",
    "10": "自行处置",
}
SORT_LABELS = {
    "1": "按时间排序",
    "500": "默认排序",
    "501": "当前价格由高到低",
    "502": "当前价格由低到高",
    "503": "出价次数由高到低",
    "507": "出价次数由低到高",
    "504": "最新发布",
}

JD_COMPAT_PRIORITY_COLUMNS = [
    "标的物ID",
    "链接",
    "当前价",
    "起拍价",
    "保证金",
    "评估价",
    "加价幅度",
    "竞价周期",
    "延时周期",
    "标题",
    "完整地址",
    "处置法院",
    "城市",
    "标的物详情描述",
    "标的物名称",
    "权利来源",
    "权证情况",
    "被执行人",
    "钥匙",
    "户籍注册",
    "欠费情况",
    "提供文件",
    "建筑面积",
    "房屋类型",
    "房屋用途",
    "总层数",
    "核准日期",
    "所有权来源",
    "土地用途",
    "土地性质",
    "使用期限",
    "所在层",
    "坐标",
    "省",
    "市",
    "区县",
    "格式化地址",
    "交通-地铁站",
    "交通-公交站",
    "教育-幼儿园",
    "教育-小学",
    "教育-中学",
    "购物-商场",
    "购物-超市",
    "医疗-综合医院",
    "医疗-诊所",
    "公园-公园",
    "大家都在问_QA",
    "图片链接",
    "图片数量",
    "竞买公告",
    "竞买须知",
    "拍卖公告",
    "详情页截图路径",
    "正文截图路径",
    "附件索引",
    "资源目录",
    "解析状态",
]

ALI_OUTPUT_PRIORITY_COLUMNS = [
    "平台",
    "详情页URL",
    "列表页码",
    "抓取时间",
    "省份筛选",
    "城市筛选",
    "查询区域",
    "查询区域编码",
    "查询区域级别",
    "状态筛选",
    "轮次筛选",
    "资产类型筛选",
    "排序筛选",
    "省份",
    "城市",
    "locationCode",
    "类目ID",
    "类目",
    "标的名称",
    "列表原文",
    "标的物介绍原文",
    "成交价/获拍价_元",
    "成交时间",
    "结束时间",
    "成交时间/结束时间",
    "是否成交",
    "是否流拍",
    "是否变卖",
    "是否支持贷款",
    "是否有金融服务",
    "金融机构",
    "最高可贷比例",
    "参考利率",
    "金融其他费用",
    "startTime",
    "endTime",
    "status",
    "statusOrder",
    "tags",
    "标签",
    "fcatV4Ids",
    "hasBid",
    "govStatus",
    "h5BidStatus",
    "public_detail_json",
    "hArea_raw",
    "incrementnum_raw",
    "回填时间",
]

LEGACY_GB2260_CANDIDATES = [
    PROJECT_ROOT / r"data\gb2260_200712.json",
    Path(r"C:\Users\Administrator\Downloads\_repo_inspect\auction-mcp\gb2260_200712.json"),
]
AREAS_JSON_CANDIDATES = [
    PROJECT_ROOT / r"data\areas.json",
    Path(r"C:\Users\Administrator\Downloads\_repo_inspect\Administrative-divisions-of-China\dist\areas.json"),
]

SPECIAL_GD_TOWNS: Dict[str, List[Tuple[str, str]]] = {
    "441900": [
        ("441901103", "茶山镇"),
        ("441901110", "常平镇"),
        ("441901119", "长安镇"),
        ("441901113", "大朗镇"),
        ("441901118", "大岭山镇"),
        ("441901124", "道滘镇"),
        ("441901003", "东城街道"),
        ("441901109", "东坑镇"),
        ("441901403", "东莞生态园"),
        ("441901117", "凤岗镇"),
        ("441901129", "高埗镇"),
        ("441901106", "横沥镇"),
        ("441901125", "洪梅镇"),
        ("441901122", "厚街镇"),
        ("441901402", "虎门港管委会"),
        ("441901121", "虎门镇"),
        ("441901114", "黄江镇"),
        ("441901126", "麻涌镇"),
        ("441901004", "南城街道"),
        ("441901105", "企石镇"),
        ("441901107", "桥头镇"),
        ("441901115", "清溪镇"),
        ("441901123", "沙田镇"),
        ("441901102", "石龙镇"),
        ("441901104", "石排镇"),
        ("441901101", "石碣镇"),
        ("441901401", "松山湖管委会"),
        ("441901116", "塘厦镇"),
        ("441901005", "万江街道"),
        ("441901127", "望牛墩镇"),
        ("441901108", "谢岗镇"),
        ("441901112", "樟木头镇"),
        ("441901128", "中堂镇"),
        ("441901006", "莞城街道"),
        ("441901111", "寮步镇"),
    ],
    "442000": [
        ("442001115", "板芙镇"),
        ("442001116", "大涌镇"),
        ("442001103", "东凤镇"),
        ("442001002", "东区街道"),
        ("442001104", "东升镇"),
        ("442001112", "阜沙镇"),
        ("442001108", "港口镇"),
        ("442001105", "古镇镇"),
        ("442001110", "横栏镇"),
        ("442001101", "黄圃镇"),
        ("442001003", "火炬开发区街道"),
        ("442001102", "民众镇"),
        ("442001113", "南朗镇"),
        ("442001005", "南区街道"),
        ("442001111", "南头镇"),
        ("442001109", "三角镇"),
        ("442001114", "三乡镇"),
        ("442001106", "沙溪镇"),
        ("442001117", "神湾镇"),
        ("442001001", "石岐区街道"),
        ("442001107", "坦洲镇"),
        ("442001006", "五桂山街道"),
        ("442001004", "西区街道"),
        ("442001100", "小榄镇"),
    ],
}

thread_local = threading.local()
CITY_CODE_TO_LABEL = {
    code: label
    for province in PROVINCE_PRESETS.values()
    for code, label in province.get("cities", [])
}


@dataclass(frozen=True)
class Partition:
    category_id: str
    category_name: str
    fcat_v4_ids: Tuple[str, ...]
    root_location_code: str
    root_location_label: str
    location_code: str
    location_label: str
    scope: str
    location_level: str = "city"
    expected_prefix: str = ""
    status_orders: Tuple[str, ...] = ()
    circs: Tuple[str, ...] = ()
    zc_biz_types: Tuple[str, ...] = ()
    sort_order: str = ""

    @property
    def key(self) -> str:
        parts = [
            self.category_id,
            self.root_location_code,
            self.location_code,
            ",".join(self.status_orders),
            ",".join(self.circs),
            ",".join(self.zc_biz_types),
        ]
        if self.sort_order:
            parts.append(self.sort_order)
        return "|".join(parts)


def resolve_project_path(value: str | Path) -> Path:
    path = Path(value)
    if path.is_absolute():
        return path
    return PROJECT_ROOT / path


def clean_text(value: Any) -> str:
    text = str(value or "")
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def safe_label(value: str) -> str:
    cleaned = clean_text(value)
    return re.sub(r'[\\/:*?"<>|]+', "_", cleaned)


def format_filter_values(values: Sequence[str], value_map: Dict[str, str], empty: str = "全部") -> str:
    if not values:
        return empty
    labels = [value_map.get(value, value) for value in values]
    return "、".join(labels)


def expected_prefix_for_location(location_code: str) -> str:
    code = clean_text(location_code)
    if len(code) == 9 and (code.startswith("441901") or code.startswith("442001")):
        return code
    if len(code) >= 6 and code.endswith("0000"):
        return code[:2]
    if len(code) >= 6 and code.endswith("00"):
        return code[:4]
    if len(code) >= 6:
        return code[:6]
    return code


def validate_location_scoped(
    items: Sequence[Dict[str, Any]],
    expected_prefix: str,
    min_ratio: float = 0.8,
) -> Dict[str, Any]:
    if not expected_prefix or not items:
        return {"ok": True, "matched": 0, "total": len(items), "sample_off_prefix": []}
    matched = 0
    off_prefix: List[str] = []
    for item in items:
        extra = item.get("extraMap") or {}
        location_code = clean_text(extra.get("locationCode") or item.get("locationCode"))
        if not location_code:
            continue
        if location_code.startswith(expected_prefix):
            matched += 1
        elif len(off_prefix) < 10:
            off_prefix.append(location_code)
    total = len(items)
    ok = (matched / total) >= min_ratio if total else True
    return {"ok": ok, "matched": matched, "total": total, "sample_off_prefix": off_prefix}


def first_existing_path(paths: Sequence[Path]) -> Optional[Path]:
    for path in paths:
        if path.exists():
            return path
    return None


def normalize_division_token(value: Any) -> str:
    text = clean_text(value)
    for suffix in [
        "特别合作区",
        "自治县",
        "开发区街道",
        "开发区",
        "街道",
        "区",
        "县",
        "市",
        "旗",
    ]:
        if text.endswith(suffix):
            text = text[: -len(suffix)]
    return clean_text(text)


def load_legacy_gb2260(path_override: str | Path = "") -> List[Dict[str, Any]]:
    path = resolve_project_path(path_override) if path_override else first_existing_path(LEGACY_GB2260_CANDIDATES)
    if not path or not path.exists():
        tried = [str(path) for path in LEGACY_GB2260_CANDIDATES]
        raise FileNotFoundError(
            "legacy gb2260 path not found. Provide --legacy-gb2260 or place file at one of: "
            + "; ".join(tried)
        )
    with path.open("r", encoding="utf-8") as fp:
        payload = json.load(fp)
    if not isinstance(payload, list):
        raise ValueError(f"legacy gb2260 payload is not a list: {path}")
    return payload


def resolve_legacy_district_code(
    legacy_gb2260: Sequence[Dict[str, Any]],
    *,
    province_code: str,
    city_code: str,
    district_name: str,
) -> str:
    province_prefix = clean_text(province_code)[:2]
    city_prefix = clean_text(city_code)[:4]
    district_token = normalize_division_token(district_name)
    if not district_token:
        return ""

    province_node = next(
        (node for node in legacy_gb2260 if clean_text(node.get("code")) == province_prefix),
        None,
    )
    if not province_node:
        return ""
    city_node = next(
        (node for node in province_node.get("children") or [] if clean_text(node.get("code")) == city_prefix),
        None,
    )
    if not city_node:
        return ""

    for child in city_node.get("children") or []:
        child_code = clean_text(child.get("code"))
        child_name = clean_text(child.get("name"))
        child_token = normalize_division_token(child_name)
        if not child_code or not child_token:
            continue
        if child_token == district_token or district_token in child_token or child_token in district_token:
            return child_code
    return ""


def load_modern_areas(
    path_override: str | Path = "",
    legacy_path_override: str | Path = "",
) -> List[Dict[str, Any]]:
    path = resolve_project_path(path_override) if path_override else first_existing_path(AREAS_JSON_CANDIDATES)
    if not path or not path.exists():
        tried = [str(path) for path in AREAS_JSON_CANDIDATES]
        raise FileNotFoundError(
            "modern areas path not found. Provide --areas-json or place file at one of: "
            + "; ".join(tried)
        )
    with path.open("r", encoding="utf-8") as fp:
        payload = json.load(fp)
    if not isinstance(payload, list):
        raise ValueError(f"modern areas payload is not a list: {path}")
    try:
        legacy_gb2260 = load_legacy_gb2260(legacy_path_override)
    except Exception:
        legacy_gb2260 = []

    enriched: List[Dict[str, Any]] = []
    for item in payload:
        if not isinstance(item, dict):
            continue
        row = dict(item)
        code = clean_text(row.get("code"))
        ali_code = resolve_legacy_district_code(
            legacy_gb2260,
            province_code=clean_text(row.get("provinceCode")),
            city_code=clean_text(row.get("cityCode")),
            district_name=clean_text(row.get("name")),
        )
        row["aliCode"] = ali_code or code
        enriched.append(row)
    return enriched


def legacy_city_children(
    location_seed_data: Sequence[Dict[str, Any]],
    province_code: str,
    province_label: str,
    city_code: str,
) -> List[Tuple[str, str, str]]:
    if city_code in SPECIAL_GD_TOWNS:
        return [(code, label, "town") for code, label in SPECIAL_GD_TOWNS[city_code]]

    if location_seed_data and "cityCode" in (location_seed_data[0] or {}):
        city_prefix = clean_text(city_code)[:4]
        results = []
        for item in location_seed_data:
            if clean_text(item.get("cityCode")) != city_prefix:
                continue
            code = clean_text(item.get("aliCode") or item.get("code"))
            label = clean_text(item.get("name"))
            if len(code) != 6 or code == city_code or not label:
                continue
            results.append((code, label, "district"))
        return results

    province_prefix = clean_text(province_code)[:2]
    province_short = clean_text(province_label).rstrip("省市自治区")
    province_node = next(
        (
            node
            for node in location_seed_data
            if clean_text(node.get("code")) == province_prefix
            or province_short in clean_text(node.get("name"))
        ),
        None,
    )
    if not province_node:
        return []

    target_city_prefix = clean_text(city_code)[:4]
    city_children = province_node.get("children") or []
    city_node = next(
        (
            node
            for node in city_children
            if clean_text(node.get("code")) == target_city_prefix
        ),
        None,
    )
    if not city_node:
        return []

    results: List[Tuple[str, str, str]] = []
    for child in city_node.get("children") or []:
        code = clean_text(child.get("code"))
        label = clean_text(child.get("name"))
        if len(code) != 6 or not label or label in {"市辖区", "县"}:
            continue
        results.append((code, label, "district"))
    return results


@dataclass(frozen=True)
class ProbeResult:
    total: int
    page: int
    page_size: int
    actual_total_pages: int
    total_pages: int
    first_items: Tuple[Dict[str, Any], ...]


def compact_json(value: Any) -> str:
    if value in ("", None, [], {}):
        return ""
    return json.dumps(value, ensure_ascii=False, separators=(",", ":"))


def to_int(value: Any, default: int = 0) -> int:
    try:
        if value in ("", None):
            return default
        return int(float(value))
    except Exception:
        return default


def cents_to_yuan(value: Any) -> Any:
    try:
        if value in ("", None):
            return ""
        number = float(value)
    except Exception:
        return ""
    if math.isnan(number):
        return ""
    return round(number / 100.0, 2)


def ts_to_text(value: Any) -> str:
    if value in ("", None):
        return ""
    try:
        number = float(value)
    except Exception:
        return ""
    if number <= 0:
        return ""
    if number > 10_000_000_000:
        number = number / 1000.0
    return datetime.fromtimestamp(number).strftime("%Y-%m-%d %H:%M:%S")


def sanitize_excel_value(value: Any) -> Any:
    if isinstance(value, str):
        return ILLEGAL_XML_RE.sub("", value)
    return value


def sanitize_dataframe_for_excel(df: pd.DataFrame) -> pd.DataFrame:
    return df.apply(lambda col: col.map(sanitize_excel_value))


def normalize_url(value: Any) -> str:
    text = clean_text(value)
    if not text:
        return ""
    if text.startswith("//"):
        return "https:" + text
    if text.startswith("/"):
        return "https://sf.taobao.com" + text
    return text


def normalize_item_id(value: Any) -> str:
    text = clean_text(value)
    if not text:
        return ""
    if text.endswith(".0") and text[:-2].isdigit():
        return text[:-2]
    if text.isdigit():
        return text
    match = re.search(r"sf_item/(\d+)\.htm", text)
    if match:
        return match.group(1)
    return ""


def tag_aliases(tags: Sequence[Dict[str, Any]]) -> List[str]:
    aliases: List[str] = []
    for tag in tags:
        alias = clean_text(tag.get("alias"))
        if alias:
            aliases.append(alias)
    return aliases


def infer_has_finance(tags: Sequence[Dict[str, Any]]) -> bool:
    aliases = tag_aliases(tags)
    keywords = ("支持贷款", "可配资", "赊钱报名", "金融", "贷款")
    return any(any(keyword in alias for keyword in keywords) for alias in aliases)


def location_city_label(location_code: Any, fallback: str = "") -> str:
    code = clean_text(location_code)
    if not code.isdigit():
        return fallback
    if len(code) >= 6:
        city_code = code[:4] + "00"
        if city_code in CITY_CODE_TO_LABEL:
            return CITY_CODE_TO_LABEL[city_code]
    return fallback


def coordinate_fields(value: Any) -> Tuple[str, str, str]:
    if not isinstance(value, list) or len(value) < 2:
        return "", "", ""
    try:
        lng = float(value[0])
        lat = float(value[1])
    except Exception:
        return "", "", ""
    return str(lng), str(lat), f"{lng},{lat}"


def infer_sale_flags(
    list_status: str,
    bid_count: int,
    gov_status: Any,
    h5_bid_status: Any,
) -> Tuple[bool, bool]:
    gov = to_int(gov_status, default=-1)
    h5 = to_int(h5_bid_status, default=-1)
    if gov == 2 or h5 == 2:
        return True, False
    if gov == 4 or h5 == 6:
        return False, True
    if list_status == "end" and bid_count > 0:
        return True, False
    if list_status == "end" and bid_count == 0:
        return False, True
    return False, False


def build_listing_digest(row: Dict[str, Any]) -> str:
    lines = [
        row.get("标题", ""),
        row.get("类目", ""),
        row.get("城市", ""),
        f"当前价:{row.get('当前价_元')}" if row.get("当前价_元") != "" else "",
        f"起拍价:{row.get('起拍价_元')}" if row.get("起拍价_元") != "" else "",
        f"保证金:{row.get('保证金_元')}" if row.get("保证金_元") != "" else "",
        f"围观:{row.get('围观次数')}" if row.get("围观次数") != "" else "",
        f"出价:{row.get('出价次数')}" if row.get("出价次数") != "" else "",
        row.get("竞价状态", ""),
        row.get("结束时间", ""),
    ]
    return "\n".join(str(x) for x in lines if str(x).strip())


def apply_compatibility_aliases(row: Dict[str, Any]) -> Dict[str, Any]:
    normalized = dict(row)
    normalized["当前价"] = normalized.get("当前价", normalized.get("当前价_元", ""))
    normalized["起拍价"] = normalized.get("起拍价", normalized.get("起拍价_元", ""))
    normalized["保证金"] = normalized.get("保证金", normalized.get("保证金_元", ""))
    normalized["评估价"] = normalized.get("评估价", normalized.get("评估价_元", ""))
    normalized["成交价/获拍价"] = normalized.get("成交价/获拍价", normalized.get("成交价/获拍价_元", ""))
    normalized["成交状态"] = normalized.get("成交状态", normalized.get("竞价状态", ""))
    normalized["成交时间/结束时间"] = normalized.get("成交时间/结束时间", normalized.get("结束时间", ""))
    normalized["成交时间"] = normalized.get("成交时间", normalized.get("成交时间/结束时间", "") if normalized.get("是否成交") == "是" else "")
    normalized["处置法院"] = normalized.get("处置法院", normalized.get("法院", ""))
    normalized["标的物详情描述"] = normalized.get("标的物详情描述", normalized.get("列表原文", ""))
    normalized["标的物介绍原文"] = normalized.get("标的物介绍原文", normalized.get("列表原文", ""))
    normalized["图片数量"] = normalized.get("图片数量", 1 if clean_text(normalized.get("图片链接")) else "")
    normalized["详情页截图路径"] = normalized.get("详情页截图路径", "")
    normalized["正文截图路径"] = normalized.get("正文截图路径", "")
    normalized["附件索引"] = normalized.get("附件索引", "")
    normalized["资源目录"] = normalized.get("资源目录", "")
    normalized["解析状态"] = normalized.get("解析状态", "已索引")
    return normalized


def ordered_output_columns(rows: Iterable[Dict[str, Any]]) -> List[str]:
    seen = set()
    ordered: List[str] = []
    for column in JD_COMPAT_PRIORITY_COLUMNS + ALI_OUTPUT_PRIORITY_COLUMNS:
        if column not in seen:
            ordered.append(column)
            seen.add(column)
    for row in rows:
        for column in row.keys():
            if column not in seen:
                ordered.append(column)
                seen.add(column)
    return ordered


def finalize_output_frame(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    df = df.fillna("").astype(object)
    return df.reindex(columns=ordered_output_columns(df.to_dict("records")), fill_value="")


class AliH5Client:
    _TOKEN_ERROR_MARKERS = (
        "TOKEN_EMPTY",
        "TOKEN_EXPIRED",
        "ILLEGAL_ACCESS::Sign Error!",
        "ILLEGAL_REQUEST",
    )

    def __init__(self) -> None:
        self.session = requests.Session()
        self.session.headers.update(
            {
                "User-Agent": MOBILE_UA,
                "Accept": "application/json",
                "Referer": HOME_URL,
            }
        )
        self._tk_token: Optional[str] = None

    def _bootstrap_token(self) -> None:
        if self._tk_token:
            return
        url = f"{H5_GATEWAY}/h5/mtop.taobao.datafront.invoke.auctionwalle/1.0/"
        params = {
            "jsv": "2.7.5",
            "appKey": H5_APPKEY,
            "t": str(int(time.time() * 1000)),
            "sign": "0" * 32,
            "api": "mtop.taobao.datafront.invoke.auctionwalle",
            "v": "1.0",
            "type": "originaljson",
            "dataType": "json",
        }
        self.session.get(url, params=params, timeout=20)
        tk_full = self.session.cookies.get("_m_h5_tk")
        if not tk_full or "_" not in tk_full:
            self.session.get(HOME_URL, timeout=20)
            tk_full = self.session.cookies.get("_m_h5_tk")
        if not tk_full or "_" not in tk_full:
            raise RuntimeError("failed to obtain _m_h5_tk cookie")
        self._tk_token = tk_full.split("_", 1)[0]

    def _sign(self, t_ms: str, data_str: str) -> str:
        raw = f"{self._tk_token}&{t_ms}&{H5_APPKEY}&{data_str}"
        return hashlib.md5(raw.encode("utf-8")).hexdigest()

    def _do_call(self, api: str, version: str, data_str: str) -> Dict[str, Any]:
        t_ms = str(int(time.time() * 1000))
        sign = self._sign(t_ms, data_str)
        url = f"{H5_GATEWAY}/h5/{api}/{version}/"
        params = {
            "jsv": "2.7.5",
            "appKey": H5_APPKEY,
            "t": t_ms,
            "sign": sign,
            "api": api,
            "v": version,
            "type": "originaljson",
            "dataType": "json",
        }
        response = self.session.post(url, params=params, data={"data": data_str}, timeout=25)
        try:
            return response.json()
        except Exception:
            preview = response.text[:200] if response.text else ""
            return {
                "ret": ["LOCAL_NON_JSON::响应非 JSON, 多为风控页或网关异常"],
                "_status": response.status_code,
                "_raw_preview": preview,
            }

    def call_mtop(self, api: str, version: str, data: Dict[str, Any]) -> Dict[str, Any]:
        self._bootstrap_token()
        data_str = json.dumps(data, separators=(",", ":"), ensure_ascii=False)
        response = self._do_call(api, version, data_str)
        ret_first = ""
        if isinstance(response.get("ret"), list) and response["ret"]:
            ret_first = str(response["ret"][0])
        if any(marker in ret_first for marker in self._TOKEN_ERROR_MARKERS):
            self._tk_token = None
            try:
                self.session.cookies.delete("_m_h5_tk")
                self.session.cookies.delete("_m_h5_tk_enc")
            except Exception:
                pass
            self._bootstrap_token()
            response = self._do_call(api, version, data_str)
            response["_token_refreshed"] = True
        return response

    def search_judicial(
        self,
        *,
        page: int,
        location_code: str,
        fcat_v4_ids: Sequence[str],
        sort: str,
        status_orders: Sequence[str],
        circs: Sequence[str] = (),
        zc_biz_types: Sequence[str] = (),
    ) -> Dict[str, Any]:
        filters: Dict[str, Any] = {"sort": sort}
        if status_orders:
            filters["statusOrders"] = list(status_orders)
        if fcat_v4_ids:
            filters["fcatV4Ids"] = list(fcat_v4_ids)
        if location_code:
            filters["locationCodes"] = [location_code]
        if circs:
            filters["circs"] = list(circs)
        if zc_biz_types:
            filters["zcBizTypes"] = list(zc_biz_types)

        filters_str = json.dumps(filters, separators=(",", ":"), ensure_ascii=False)
        user_info = json.dumps(
            {"prov": "", "city": "", "locationCode": location_code},
            separators=(",", ":"),
            ensure_ascii=False,
        )
        df_variables = {
            "page": page,
            "pageSpmb": "sf-home",
            "pageSpmcs": "searchlistsf-items",
            "context": {
                "_c_searchlistsf-items": filters_str,
                "prov": "",
                "city": "",
                "locationCode": location_code,
                "userInfo": user_info,
                "piPageType": "original",
            },
        }
        data = {
            "dfApp": "auctionwalle",
            "dfApiName": "auctionwalle.page.getScenes",
            "dfVariables": json.dumps(df_variables, separators=(",", ":"), ensure_ascii=False),
            "dfUniqueId": "sf-home_searchlistsf-items",
            "dfVariablesRecover": "{}",
        }
        return self.call_mtop("mtop.taobao.datafront.invoke.auctionwalle", "1.0", data)


def make_h5_client() -> AliH5Client:
    return AliH5Client()


def get_h5_client() -> AliH5Client:
    client = getattr(thread_local, "ali_h5_client", None)
    if client is None:
        client = make_h5_client()
        thread_local.ali_h5_client = client
    return client


def reset_h5_client() -> None:
    thread_local.ali_h5_client = make_h5_client()


def extract_scheme_list(payload: Dict[str, Any]) -> Tuple[int, int, int, List[Dict[str, Any]]]:
    ret_first = ""
    if isinstance(payload.get("ret"), list) and payload["ret"]:
        ret_first = str(payload["ret"][0])
    if ret_first != "SUCCESS::调用成功":
        raise RuntimeError(
            f"mtop_call_failed ret={payload.get('ret')} preview={payload.get('_raw_preview', '')}"
        )
    scenes = (((payload.get("data") or {}).get("data") or {}).get("scenes") or [])
    if not scenes:
        return 0, 1, PAGE_SIZE, []
    scheme = (scenes[0].get("schemeList") or [{}])[0]
    items = scheme.get("contentList") or []
    total = to_int(scheme.get("totalCount"), default=0)
    page = to_int(scheme.get("page"), default=1)
    page_size = to_int(scheme.get("pageSize"), default=PAGE_SIZE) or PAGE_SIZE
    return total, page, page_size, [item for item in items if isinstance(item, dict)]


def fetch_list_page(
    partition: Partition,
    page: int,
    sort: str,
    retries: int = 5,
    sleep_seconds: float = 0.8,
) -> Tuple[int, int, int, List[Dict[str, Any]]]:
    last_error: Optional[Exception] = None
    for attempt in range(1, retries + 1):
        try:
            effective_sort = partition.sort_order or sort
            payload = get_h5_client().search_judicial(
                page=page,
                location_code=partition.location_code,
                fcat_v4_ids=partition.fcat_v4_ids,
                sort=effective_sort,
                status_orders=partition.status_orders,
                circs=partition.circs,
                zc_biz_types=partition.zc_biz_types,
            )
            total, page_back, page_size, items = extract_scheme_list(payload)
            validation = validate_location_scoped(items, partition.expected_prefix)
            if not validation["ok"]:
                raise RuntimeError(
                    "ali_returned_unscoped_results "
                    f"expected_prefix={partition.expected_prefix} matched={validation['matched']} "
                    f"total={validation['total']} sample_off_prefix={validation['sample_off_prefix']}"
                )
            return total, page_back, page_size, items
        except Exception as exc:  # noqa: BLE001
            last_error = exc
            message = str(exc)
            if "LOCAL_NON_JSON" in message or "mtop_call_failed" in message:
                reset_h5_client()
            time.sleep(sleep_seconds * attempt)
    raise RuntimeError(
        f"fetch_list_page failed partition={partition.key} page={page}: {last_error}"
    )


def probe_partition(partition: Partition, sort: str, page_limit: int) -> ProbeResult:
    total, page, page_size, items = fetch_list_page(partition, 1, sort)
    actual_total_pages = int(math.ceil(total / page_size)) if total else 0
    total_pages = min(actual_total_pages, page_limit) if page_limit > 0 else actual_total_pages
    return ProbeResult(
        total=total,
        page=page,
        page_size=page_size,
        actual_total_pages=actual_total_pages,
        total_pages=total_pages,
        first_items=tuple(items),
    )


def parse_jsonp_payload(text: str) -> Dict[str, Any]:
    match = JSONP_RE.search(text)
    if not match:
        return {}
    try:
        data = json.loads(match.group(1))
    except Exception:
        return {}
    return data if isinstance(data, dict) else {}


def fetch_public_bid_detail(item_id: str, retries: int = 3) -> Dict[str, Any]:
    if not item_id:
        return {}
    url = f"https://sf-item.taobao.com/json/get_bid_detail.htm?id={item_id}"
    headers = {"User-Agent": "Mozilla/5.0", "Referer": HOME_URL}
    last_error: Optional[Exception] = None
    for attempt in range(1, retries + 1):
        try:
            response = requests.get(url, headers=headers, timeout=25)
            response.raise_for_status()
            payload = parse_jsonp_payload(response.text)
            data = payload.get("data") or {}
            if not isinstance(data, dict):
                return {}
            return data
        except Exception as exc:  # noqa: BLE001
            last_error = exc
            time.sleep(0.5 * attempt)
    print(f"public_bid_detail_failed item_id={item_id} error={last_error}", flush=True)
    return {}


def enrich_with_public_bid_detail(
    items: Sequence[Dict[str, Any]],
    workers: int,
) -> Dict[str, Dict[str, Any]]:
    targets = [
        normalize_item_id(item.get("itemId"))
        for item in items
        if clean_text(item.get("status")) in {"end", "pause", "revoke"}
    ]
    targets = [item_id for item_id in targets if item_id]
    if not targets:
        return {}
    results: Dict[str, Dict[str, Any]] = {}
    with ThreadPoolExecutor(max_workers=max(1, workers)) as executor:
        future_map = {
            executor.submit(fetch_public_bid_detail, item_id): item_id
            for item_id in targets
        }
        for future in as_completed(future_map):
            item_id = future_map[future]
            try:
                results[item_id] = future.result()
            except Exception:
                results[item_id] = {}
    return results


def map_row(
    item: Dict[str, Any],
    *,
    partition: Partition,
    province_label: str,
    page_no: int,
    public_bid_detail: Optional[Dict[str, Any]] = None,
) -> Dict[str, Any]:
    public_bid_detail = public_bid_detail or {}
    extra = item.get("extraMap") or {}
    item_id = normalize_item_id(item.get("itemId"))
    list_status = clean_text(item.get("status"))
    bid_count = to_int(public_bid_detail.get("bidCount"), default=to_int(item.get("bidCnt"), default=0))
    apply_count = public_bid_detail.get("applyCnt", "")
    delay_count = public_bid_detail.get("delayCount", "")
    current_price = cents_to_yuan(public_bid_detail.get("currentPrice", item.get("currentPrice")))
    start_price = cents_to_yuan(item.get("initialPrice"))
    bail_price = cents_to_yuan(extra.get("bail"))
    market_price = cents_to_yuan(item.get("marketPrice"))
    consult_price = cents_to_yuan(item.get("consultPrice"))
    location_code = clean_text(item.get("locationCode"))
    city_label = location_city_label(location_code, fallback=partition.location_label if partition.scope == "city" else "")
    tags = item.get("tags") or []
    has_finance = infer_has_finance(tags)
    gov_status = public_bid_detail.get("govStatus", "")
    h5_bid_status = public_bid_detail.get("h5BidStatus", "")
    sold, unsold = infer_sale_flags(list_status, bid_count, gov_status, h5_bid_status)
    end_time = ts_to_text(item.get("endTime"))
    lng, lat, coordinate_text = coordinate_fields(item.get("coordinate"))
    title = clean_text(item.get("title"))

    row = {
        "平台": "阿里法拍",
        "标的物ID": item_id,
        "链接": f"https://sf-item.taobao.com/sf_item/{item_id}.htm" if item_id else "",
        "详情页URL": f"https://sf-item.taobao.com/sf_item/{item_id}.htm" if item_id else "",
        "列表页码": page_no,
        "抓取时间": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "省份筛选": province_label,
        "城市筛选": partition.root_location_label,
        "查询区域": partition.location_label,
        "查询区域编码": partition.location_code,
        "查询区域级别": partition.location_level,
        "状态筛选": format_filter_values(partition.status_orders, STATUS_ORDER_LABELS),
        "轮次筛选": format_filter_values(partition.circs, CIRC_LABELS),
        "资产类型筛选": format_filter_values(partition.zc_biz_types, ZC_BIZ_TYPE_LABELS),
        "排序筛选": SORT_LABELS.get(partition.sort_order, partition.sort_order) if partition.sort_order else "",
        "省份": province_label,
        "城市": city_label or partition.root_location_label,
        "locationCode": location_code,
        "类目ID": partition.category_id,
        "类目": partition.category_name,
        "标题": title,
        "标的名称": title,
        "标的物名称": title,
        "当前价": current_price,
        "当前价_元": current_price,
        "起拍价": start_price,
        "起拍价_元": start_price,
        "保证金": bail_price,
        "保证金_元": bail_price,
        "评估价": consult_price,
        "评估价_元": consult_price,
        "市场价_元": market_price,
        "成交价/获拍价": current_price if sold and current_price != "" else "",
        "成交价/获拍价_元": current_price if sold and current_price != "" else "",
        "围观次数": to_int(item.get("pv"), default=0) or "",
        "报名人数": apply_count,
        "提醒人数": to_int(item.get("subscribeCnt"), default=0) or "",
        "出价次数": bid_count,
        "延时次数": delay_count,
        "竞价状态": STATUS_TEXT_MAP.get(list_status, list_status),
        "成交状态": STATUS_TEXT_MAP.get(list_status, list_status),
        "status": list_status,
        "statusOrder": item.get("statusOrder", ""),
        "是否成交": "是" if sold else "否",
        "是否流拍": "是" if unsold else "否",
        "成交时间": end_time if sold else "",
        "成交时间/结束时间": end_time,
        "开始时间": ts_to_text(item.get("startTime")),
        "结束时间": end_time,
        "startTime": item.get("startTime", ""),
        "endTime": item.get("endTime", ""),
        "是否变卖": "是" if "变卖" in tag_aliases(tags) else "否",
        "是否有金融服务": "是" if has_finance else "否",
        "是否支持贷款": "是" if any("贷款" in alias for alias in tag_aliases(tags)) else "",
        "是否支持机构贷款": "",
        "金融机构": "",
        "最高可贷比例": "",
        "参考利率": "",
        "金融其他费用": "",
        "竣工时间": "",
        "购买时间": "",
        "建筑面积": "",
        "房屋用途": "",
        "房屋类型": "",
        "所在层": "",
        "总层数": "",
        "土地性质": "",
        "土地用途": "",
        "使用期限": "",
        "权利来源": "",
        "所有权来源": "",
        "权证情况": "",
        "标的所有人": "",
        "产权证号": "",
        "法院": clean_text(item.get("shopName")),
        "courtName": clean_text(item.get("shopName")),
        "处置机构": clean_text(item.get("shopName")),
        "图片链接": normalize_url(item.get("picURL")),
        "图片数量": 1 if normalize_url(item.get("picURL")) else "",
        "坐标": coordinate_text,
        "经度": lng,
        "纬度": lat,
        "displayCurrentPrice": clean_text(item.get("displayCurrentPrice")),
        "displayCurrentPriceUnit": clean_text(item.get("displayCurrentPriceUnit")),
        "displayInitialPrice": clean_text(item.get("displayInitialPrice")),
        "displayInitialPriceUnit": clean_text(item.get("displayInitialPriceUnit")),
        "fcatV4Ids": compact_json(item.get("fcatV4Ids")),
        "标签": "、".join(dict.fromkeys(tag_aliases(tags))),
        "tags": compact_json(tags),
        "govStatus": gov_status,
        "h5BidStatus": h5_bid_status,
        "hasBid": public_bid_detail.get("hasBid", ""),
        "hArea_raw": extra.get("hArea", ""),
        "incrementnum_raw": extra.get("incrementnum", ""),
        "public_detail_json": compact_json(public_bid_detail) if public_bid_detail else "",
    }
    row["列表原文"] = build_listing_digest(row)
    row["标的物介绍原文"] = row["列表原文"]
    row["标的物详情描述"] = row["列表原文"]
    row["处置法院"] = row["法院"]
    row = apply_compatibility_aliases(row)
    return row


def dedupe_rows(rows: List[Dict[str, Any]]) -> pd.DataFrame:
    if not rows:
        return pd.DataFrame()
    df = pd.DataFrame(rows).fillna("").astype(object)
    if "标的物ID" in df.columns:
        df["_non_empty"] = df.apply(lambda row: sum(str(v).strip() != "" for v in row), axis=1)
        df = df.sort_values(["标的物ID", "_non_empty"], ascending=[True, False], kind="stable")
        df = df.drop_duplicates("标的物ID", keep="first").drop(columns=["_non_empty"])
    sort_cols = [col for col in ["类目ID", "城市筛选", "endTime", "标的物ID"] if col in df.columns]
    if sort_cols:
        df = df.sort_values(sort_cols, kind="stable")
    return finalize_output_frame(df)


def save_rows(rows: List[Dict[str, Any]], path: Path) -> int:
    path.parent.mkdir(parents=True, exist_ok=True)
    df = dedupe_rows(rows)
    clean_df = sanitize_dataframe_for_excel(finalize_output_frame(df))
    tmp_path = path.with_suffix(path.suffix + ".tmp.xlsx")
    with pd.ExcelWriter(
        tmp_path,
        engine="xlsxwriter",
        engine_kwargs={"options": {"strings_to_urls": False}},
    ) as writer:
        clean_df.to_excel(writer, index=False)
    tmp_path.replace(path)
    return len(clean_df)


def save_stats(rows: List[Dict[str, Any]], path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    df = sanitize_dataframe_for_excel(pd.DataFrame(rows))
    tmp_path = path.with_suffix(path.suffix + ".tmp.xlsx")
    with pd.ExcelWriter(
        tmp_path,
        engine="xlsxwriter",
        engine_kwargs={"options": {"strings_to_urls": False}},
    ) as writer:
        df.to_excel(writer, index=False)
    tmp_path.replace(path)


def load_existing_rows(path: Path) -> List[Dict[str, Any]]:
    if not path.exists():
        return []
    return pd.read_excel(path).fillna("").astype(object).to_dict("records")


def load_checkpoint(path: Path) -> Dict[str, Any]:
    if not path.exists():
        return {"partitions": {}, "stats": []}
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
        if not isinstance(payload, dict):
            raise ValueError("checkpoint must be dict")
        payload.setdefault("partitions", {})
        payload.setdefault("stats", [])
        return payload
    except Exception:
        return {"partitions": {}, "stats": []}


def write_checkpoint(path: Path, payload: Dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    copied = dict(payload)
    copied["updated_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    path.write_text(json.dumps(copied, ensure_ascii=False, indent=2), encoding="utf-8")


def partition_output_name(province_label: str, partition: Partition) -> str:
    parts = [
        "阿里法拍房",
        province_label,
        "索引",
        safe_label(partition.category_name),
        safe_label(partition.root_location_label),
    ]
    if partition.location_code != partition.root_location_code or partition.location_label != partition.root_location_label:
        parts.extend([partition.location_code, safe_label(partition.location_label)])
    if partition.status_orders and tuple(partition.status_orders) != tuple(DEFAULT_STATUS_ORDERS):
        parts.append(safe_label(format_filter_values(partition.status_orders, STATUS_ORDER_LABELS)))
    if partition.circs:
        parts.append(safe_label(format_filter_values(partition.circs, CIRC_LABELS)))
    if partition.zc_biz_types:
        parts.append(safe_label(format_filter_values(partition.zc_biz_types, ZC_BIZ_TYPE_LABELS)))
    if partition.sort_order:
        parts.append(safe_label(SORT_LABELS.get(partition.sort_order, partition.sort_order)))
    return "_".join(parts) + ".xlsx"


def merge_partition_files(files: Iterable[Path], output_path: Path) -> int:
    frames: List[pd.DataFrame] = []
    for file in files:
        if file.exists() and file.stat().st_size > 0:
            frames.append(pd.read_excel(file).fillna("").astype(object))
    if not frames:
        save_stats([], output_path)
        return 0
    merged = pd.concat(frames, ignore_index=True)
    final_df = dedupe_rows(merged.to_dict("records"))
    clean_df = sanitize_dataframe_for_excel(finalize_output_frame(final_df))
    output_path.parent.mkdir(parents=True, exist_ok=True)
    tmp_path = output_path.with_suffix(output_path.suffix + ".tmp.xlsx")
    with pd.ExcelWriter(
        tmp_path,
        engine="xlsxwriter",
        engine_kwargs={"options": {"strings_to_urls": False}},
    ) as writer:
        clean_df.to_excel(writer, index=False)
    tmp_path.replace(output_path)
    return len(clean_df)


def select_categories(category_arg: str) -> List[Tuple[str, Dict[str, Any]]]:
    if not category_arg or category_arg.lower() == "all":
        return list(CATEGORY_PRESETS.items())
    wanted = {x.strip() for x in category_arg.split(",") if x.strip()}
    selected = [
        (category_id, meta)
        for category_id, meta in CATEGORY_PRESETS.items()
        if category_id in wanted or meta["name"] in wanted
    ]
    resolved = {category_id for category_id, _ in selected} | {meta["name"] for _, meta in selected}
    missing = wanted - resolved
    if missing:
        raise ValueError(f"Unknown categories: {', '.join(sorted(missing))}")
    return selected


def select_cities(city_arg: str, cities: Sequence[Tuple[str, str]]) -> List[Tuple[str, str]]:
    if not city_arg or city_arg.lower() == "all":
        return list(cities)
    wanted = {x.strip() for x in city_arg.split(",") if x.strip()}
    selected = [(code, label) for code, label in cities if code in wanted or label in wanted]
    resolved = {code for code, _ in selected} | {label for _, label in selected}
    missing = wanted - resolved
    if missing:
        raise ValueError(f"Unknown cities: {', '.join(sorted(missing))}")
    return selected


def build_seed_partitions(
    province_label: str,
    province_code: str,
    cities: Sequence[Tuple[str, str]],
    categories: Sequence[Tuple[str, Dict[str, Any]]],
    legacy_gb2260: Sequence[Dict[str, Any]],
    status_orders: Sequence[str],
) -> List[Partition]:
    partitions: List[Partition] = []
    status_tuple = tuple(status_orders)
    for category_id, meta in categories:
        scope = meta.get("scope", "city")
        for city_code, city_label in cities:
            if scope == "province":
                partitions.append(
                    Partition(
                        category_id=category_id,
                        category_name=meta["name"],
                        fcat_v4_ids=tuple(meta["fcat_v4_ids"]),
                        root_location_code=city_code,
                        root_location_label=city_label,
                        location_code=city_code,
                        location_label=city_label,
                        scope="province",
                        location_level="city",
                        expected_prefix=expected_prefix_for_location(city_code),
                        status_orders=status_tuple,
                    )
                )
                continue

            children = legacy_city_children(legacy_gb2260, province_code, province_label, city_code)
            if not children:
                partitions.append(
                    Partition(
                        category_id=category_id,
                        category_name=meta["name"],
                        fcat_v4_ids=tuple(meta["fcat_v4_ids"]),
                        root_location_code=city_code,
                        root_location_label=city_label,
                        location_code=city_code,
                        location_label=city_label,
                        scope="city",
                        location_level="city",
                        expected_prefix=expected_prefix_for_location(city_code),
                        status_orders=status_tuple,
                    )
                )
                continue

            for location_code, location_label, location_level in children:
                partitions.append(
                    Partition(
                        category_id=category_id,
                        category_name=meta["name"],
                        fcat_v4_ids=tuple(meta["fcat_v4_ids"]),
                        root_location_code=city_code,
                        root_location_label=city_label,
                        location_code=location_code,
                        location_label=location_label,
                        scope="city",
                        location_level=location_level,
                        expected_prefix=expected_prefix_for_location(location_code),
                        status_orders=status_tuple,
                    )
                )
    return partitions


def default_paths(province_label: str) -> Dict[str, Path]:
    return {
        "checkpoint": Path(rf"output\阿里法拍房_{province_label}_索引.checkpoint.json"),
        "merge_output": Path(rf"output\阿里法拍房_{province_label}_全量索引.xlsx"),
        "stats_output": Path(rf"output\阿里法拍房_{province_label}_索引统计.xlsx"),
    }


def parse_status_orders(value: str) -> List[str]:
    parts = [x.strip() for x in clean_text(value).split(",") if x.strip()]
    return parts or list(DEFAULT_STATUS_ORDERS)


def partition_stats_meta(partition: Partition) -> Dict[str, Any]:
    return {
        "类目ID": partition.category_id,
        "类目": partition.category_name,
        "城市": partition.root_location_label,
        "查询区域": partition.location_label,
        "查询区域编码": partition.location_code,
        "查询区域级别": partition.location_level,
        "状态筛选": format_filter_values(partition.status_orders, STATUS_ORDER_LABELS),
        "轮次筛选": format_filter_values(partition.circs, CIRC_LABELS),
        "资产类型筛选": format_filter_values(partition.zc_biz_types, ZC_BIZ_TYPE_LABELS),
        "排序筛选": SORT_LABELS.get(partition.sort_order, partition.sort_order) if partition.sort_order else "",
    }


def build_child_partitions_for_split(
    partition: Partition,
    *,
    province_label: str,
    province_code: str,
    legacy_gb2260: Sequence[Dict[str, Any]],
) -> Tuple[str, List[Partition]]:
    if partition.scope == "province" and partition.location_level == "city":
        children = legacy_city_children(legacy_gb2260, province_code, province_label, partition.location_code)
        if children:
            return (
                "地域拆分",
                [
                    replace(
                        partition,
                        location_code=location_code,
                        location_label=location_label,
                        location_level=location_level,
                        expected_prefix=expected_prefix_for_location(location_code),
                    )
                    for location_code, location_label, location_level in children
                ],
            )

    allowed_statuses = [value for value in DEFAULT_STATUS_ORDERS if value in set(partition.status_orders)]
    if len(allowed_statuses) > 1:
        return (
            "状态拆分",
            [
                replace(
                    partition,
                    status_orders=(status_order,),
                    circs=(),
                    zc_biz_types=(),
                )
                for status_order in allowed_statuses
            ],
        )

    if partition.status_orders == ("2",) and not partition.circs:
        return (
            "轮次拆分",
            [
                replace(
                    partition,
                    circs=circ_group,
                    zc_biz_types=(),
                    sort_order="",
                )
                for circ_group, _ in DEFAULT_CIRC_GROUPS
            ],
        )

    if not partition.zc_biz_types:
        return (
            "资产类型拆分",
            [
                replace(
                    partition,
                    zc_biz_types=(zc_biz_type,),
                    sort_order="",
                )
                for zc_biz_type in DEFAULT_ZC_BIZ_TYPES
            ],
        )

    if not partition.sort_order:
        return (
            "排序窗口拆分",
            [
                replace(
                    partition,
                    sort_order=sort_order,
                )
                for sort_order in PRICE_WINDOW_SORTS
            ],
        )

    return "", []


def crawl_leaf_partition(
    *,
    partition: Partition,
    probe: ProbeResult,
    province_label: str,
    partition_path: Path,
    checkpoint: Dict[str, Any],
    stats_rows: List[Dict[str, Any]],
    checkpoint_path: Path,
    stats_output: Path,
    sort: str,
    workers: int,
    public_bid_detail: bool,
    public_bid_detail_workers: int,
    page_chunk_size: int,
    save_every_pages: int,
) -> None:
    progress = checkpoint["partitions"].get(partition.key) or {}
    if progress.get("done") and partition_path.exists():
        print(
            f"partition_skip category={partition.category_name} city={partition.root_location_label} "
            f"query={partition.location_label} reason=done",
            flush=True,
        )
        return

    existing_rows = load_existing_rows(partition_path)
    existing_unique = len(
        {
            normalize_item_id(row.get("标的物ID"))
            for row in existing_rows
            if normalize_item_id(row.get("标的物ID"))
        }
    )
    print(
        f"partition_start category={partition.category_name} city={partition.root_location_label} "
        f"query={partition.location_label} status={format_filter_values(partition.status_orders, STATUS_ORDER_LABELS)} "
        f"circ={format_filter_values(partition.circs, CIRC_LABELS)} "
        f"existing_unique={existing_unique}",
        flush=True,
    )

    stats_rows.append(
        {
            **partition_stats_meta(partition),
            "阶段": "PLAN",
            "时间桶": "ALL_PAGES",
            "接口总量": probe.total,
            "抓取行数": "",
            "唯一标的数": existing_unique,
            "页数": probe.actual_total_pages,
            "错误": "",
        }
    )
    save_stats(stats_rows, stats_output)
    print(
        f"partition_plan category={partition.category_name} city={partition.root_location_label} "
        f"query={partition.location_label} total={probe.total} total_pages={probe.actual_total_pages} "
        f"effective_pages={probe.total_pages}",
        flush=True,
    )

    if probe.total_pages == 0:
        checkpoint["partitions"][partition.key] = {
            "next_page": 1,
            "total_pages": 0,
            "actual_total_pages": 0,
            "total": probe.total,
            "done": True,
            "output_file": str(partition_path),
        }
        checkpoint["stats"] = stats_rows
        write_checkpoint(checkpoint_path, checkpoint)
        return

    next_page = max(1, to_int(progress.get("next_page"), default=1))
    rows = existing_rows
    pages_since_save = 0

    for chunk_start in range(next_page, probe.total_pages + 1, max(1, page_chunk_size)):
        chunk_end = min(probe.total_pages, chunk_start + max(1, page_chunk_size) - 1)
        started = time.time()
        page_items: Dict[int, List[Dict[str, Any]]] = {}
        if chunk_start == 1:
            page_items[1] = list(probe.first_items)

        page_numbers = [page for page in range(chunk_start, chunk_end + 1) if page not in page_items]
        if page_numbers:
            with ThreadPoolExecutor(max_workers=max(1, workers)) as executor:
                future_map = {
                    executor.submit(fetch_list_page, partition, page, sort): page
                    for page in page_numbers
                }
                for future in as_completed(future_map):
                    page = future_map[future]
                    page_total, page_back, _, items = future.result()
                    if page_total and probe.total and page_total != probe.total:
                        print(
                            f"page_meta_changed partition={partition.key} page={page} "
                            f"first_total={probe.total} page_total={page_total} page_back={page_back}",
                            flush=True,
                        )
                    page_items[page] = items

        bid_detail_map: Dict[str, Dict[str, Any]] = {}
        if public_bid_detail:
            all_items = [item for page in range(chunk_start, chunk_end + 1) for item in page_items.get(page, [])]
            bid_detail_map = enrich_with_public_bid_detail(all_items, public_bid_detail_workers)

        chunk_rows: List[Dict[str, Any]] = []
        for page in range(chunk_start, chunk_end + 1):
            items = page_items.get(page, [])
            for item in items:
                item_id = normalize_item_id(item.get("itemId"))
                chunk_rows.append(
                    map_row(
                        item,
                        partition=partition,
                        province_label=province_label,
                        page_no=page,
                        public_bid_detail=bid_detail_map.get(item_id),
                    )
                )

        rows.extend(chunk_rows)
        pages_since_save += (chunk_end - chunk_start + 1)
        unique_after = len(
            {
                normalize_item_id(row.get("标的物ID"))
                for row in rows
                if normalize_item_id(row.get("标的物ID"))
            }
        )

        checkpoint["partitions"][partition.key] = {
            "next_page": chunk_end + 1,
            "total_pages": probe.total_pages,
            "actual_total_pages": probe.actual_total_pages,
            "total": probe.total,
            "done": chunk_end >= probe.total_pages,
            "output_file": str(partition_path),
            "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        }

        if pages_since_save >= max(1, save_every_pages) or chunk_end >= probe.total_pages:
            unique_after = save_rows(rows, partition_path)
            pages_since_save = 0

        stats_rows.append(
            {
                **partition_stats_meta(partition),
                "阶段": "CRAWL",
                "时间桶": f"{chunk_start}-{chunk_end}",
                "接口总量": probe.total,
                "抓取行数": len(chunk_rows),
                "唯一标的数": unique_after,
                "页数": chunk_end - chunk_start + 1,
                "错误": "",
            }
        )
        checkpoint["stats"] = stats_rows
        write_checkpoint(checkpoint_path, checkpoint)
        save_stats(stats_rows, stats_output)
        elapsed = time.time() - started
        print(
            f"chunk_done category={partition.category_name} city={partition.root_location_label} "
            f"query={partition.location_label} "
            f"pages={chunk_start}-{chunk_end} rows={len(chunk_rows)} unique={unique_after} "
            f"elapsed={elapsed:.1f}s",
            flush=True,
        )


def process_partition(
    *,
    partition: Partition,
    province_label: str,
    province_code: str,
    legacy_gb2260: Sequence[Dict[str, Any]],
    output_dir: Path,
    checkpoint: Dict[str, Any],
    stats_rows: List[Dict[str, Any]],
    checkpoint_path: Path,
    stats_output: Path,
    sort: str,
    workers: int,
    public_bid_detail: bool,
    public_bid_detail_workers: int,
    page_chunk_size: int,
    save_every_pages: int,
    max_pages_per_partition: int,
    partition_files: List[Path],
) -> None:
    partition_path = output_dir / partition_output_name(province_label, partition)
    if partition_path not in partition_files:
        partition_files.append(partition_path)

    progress = checkpoint.get("partitions", {}).get(partition.key) or {}
    if progress.get("done") and partition_path.exists():
        print(
            f"partition_skip category={partition.category_name} city={partition.root_location_label} "
            f"query={partition.location_label} reason=done",
            flush=True,
        )
        return

    effective_page_limit = max_pages_per_partition if max_pages_per_partition > 0 else ALI_H5_PAGE_CAP
    probe = probe_partition(partition, sort, effective_page_limit)

    if (
        max_pages_per_partition <= 0
        and probe.actual_total_pages > ALI_H5_PAGE_CAP
        and not partition.sort_order
    ):
        split_reason, child_partitions = build_child_partitions_for_split(
            partition,
            province_label=province_label,
            province_code=province_code,
            legacy_gb2260=legacy_gb2260,
        )
        if child_partitions:
            stats_rows.append(
                {
                    **partition_stats_meta(partition),
                    "阶段": "SPLIT",
                    "时间桶": split_reason,
                    "接口总量": probe.total,
                    "抓取行数": "",
                    "唯一标的数": "",
                    "页数": probe.actual_total_pages,
                    "错误": "",
                }
            )
            checkpoint["stats"] = stats_rows
            write_checkpoint(checkpoint_path, checkpoint)
            save_stats(stats_rows, stats_output)
            print(
                f"partition_split category={partition.category_name} city={partition.root_location_label} "
                f"query={partition.location_label} total_pages={probe.actual_total_pages} strategy={split_reason} "
                f"children={len(child_partitions)}",
                flush=True,
            )
            for child_partition in child_partitions:
                process_partition(
                    partition=child_partition,
                    province_label=province_label,
                    province_code=province_code,
                    legacy_gb2260=legacy_gb2260,
                    output_dir=output_dir,
                    checkpoint=checkpoint,
                    stats_rows=stats_rows,
                    checkpoint_path=checkpoint_path,
                    stats_output=stats_output,
                    sort=sort,
                    workers=workers,
                    public_bid_detail=public_bid_detail,
                    public_bid_detail_workers=public_bid_detail_workers,
                    page_chunk_size=page_chunk_size,
                    save_every_pages=save_every_pages,
                    max_pages_per_partition=max_pages_per_partition,
                    partition_files=partition_files,
                )
            return
        raise RuntimeError(
            f"partition_exceeds_page_cap category={partition.category_name} city={partition.root_location_label} "
            f"query={partition.location_label} actual_total_pages={probe.actual_total_pages}"
        )

    crawl_leaf_partition(
        partition=partition,
        probe=probe,
        province_label=province_label,
        partition_path=partition_path,
        checkpoint=checkpoint,
        stats_rows=stats_rows,
        checkpoint_path=checkpoint_path,
        stats_output=stats_output,
        sort=sort,
        workers=workers,
        public_bid_detail=public_bid_detail,
        public_bid_detail_workers=public_bid_detail_workers,
        page_chunk_size=page_chunk_size,
        save_every_pages=save_every_pages,
    )


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Crawl Alibaba judicial-auction real-estate index data via H5 API.")
    parser.add_argument("--province-key", default="gd", help="Province preset key. Currently validated for gd.")
    parser.add_argument("--cities", default="all", help="Comma-separated city codes or labels, or all.")
    parser.add_argument("--categories", default="all", help="Comma-separated category ids or names, or all.")
    parser.add_argument("--output-dir", default=r"output")
    parser.add_argument("--checkpoint", default="", help="Checkpoint JSON path.")
    parser.add_argument("--merge-output", default="", help="Merged province Excel path.")
    parser.add_argument("--stats-output", default="", help="Stats Excel path.")
    parser.add_argument("--workers", type=int, default=8)
    parser.add_argument("--page-chunk-size", type=int, default=100)
    parser.add_argument("--save-every-pages", type=int, default=1000)
    parser.add_argument("--sort", default=DEFAULT_SORT, help="Ali H5 sort value. Default 1=按时间排序.")
    parser.add_argument(
        "--status-orders",
        default=",".join(DEFAULT_STATUS_ORDERS),
        help="Comma-separated Ali H5 statusOrders. Default 0,1,2,4,5.",
    )
    parser.add_argument("--with-public-bid-detail", action="store_true", help="Call public get_bid_detail for ended/paused/revoked assets.")
    parser.add_argument("--public-bid-detail-workers", type=int, default=8)
    parser.add_argument("--areas-json", default="", help="Modern areas.json path for current district codes.")
    parser.add_argument("--legacy-gb2260", default="", help="Legacy gb2260_200712.json path for Ali-compatible district codes.")
    parser.add_argument("--max-pages-per-partition", type=int, default=0, help="Limit pages per partition for smoke tests.")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    province_key = clean_text(args.province_key).lower()
    province = PROVINCE_PRESETS.get(province_key)
    if not province:
        raise ValueError(f"Unknown province preset: {province_key}")

    province_label = province["province_label"]
    province_code = province["province_code"]
    cities = select_cities(args.cities, province["cities"])
    categories = select_categories(args.categories)

    path_defaults = default_paths(province_label)
    output_dir = resolve_project_path(args.output_dir)
    checkpoint_path = resolve_project_path(args.checkpoint or path_defaults["checkpoint"])
    merge_output = resolve_project_path(args.merge_output or path_defaults["merge_output"])
    stats_output = resolve_project_path(args.stats_output or path_defaults["stats_output"])
    status_orders = parse_status_orders(args.status_orders)
    try:
        location_seed_data = load_modern_areas(args.areas_json, args.legacy_gb2260)
    except FileNotFoundError:
        location_seed_data = load_legacy_gb2260(args.legacy_gb2260)
    partitions = build_seed_partitions(
        province_label,
        province_code,
        cities,
        categories,
        location_seed_data,
        status_orders,
    )

    checkpoint = load_checkpoint(checkpoint_path)
    stats_rows: List[Dict[str, Any]] = list(checkpoint.get("stats") or [])
    partition_files: List[Path] = []

    for partition in partitions:
        try:
            process_partition(
                partition=partition,
                province_label=province_label,
                province_code=province_code,
                legacy_gb2260=location_seed_data,
                output_dir=output_dir,
                checkpoint=checkpoint,
                stats_rows=stats_rows,
                checkpoint_path=checkpoint_path,
                stats_output=stats_output,
                sort=args.sort,
                workers=args.workers,
                public_bid_detail=args.with_public_bid_detail,
                public_bid_detail_workers=args.public_bid_detail_workers,
                page_chunk_size=args.page_chunk_size,
                save_every_pages=args.save_every_pages,
                max_pages_per_partition=args.max_pages_per_partition,
                partition_files=partition_files,
            )
        except Exception as exc:  # noqa: BLE001
            stats_rows.append(
                {
                    **partition_stats_meta(partition),
                    "阶段": "ERROR",
                    "时间桶": "ERROR",
                    "接口总量": "",
                    "抓取行数": "",
                    "唯一标的数": "",
                    "页数": "",
                    "错误": str(exc),
                }
            )
            checkpoint["stats"] = stats_rows
            write_checkpoint(checkpoint_path, checkpoint)
            save_stats(stats_rows, stats_output)
            print(
                f"partition_error category={partition.category_name} city={partition.root_location_label} "
                f"query={partition.location_label} error={exc}",
                flush=True,
            )

    checkpoint["stats"] = stats_rows
    write_checkpoint(checkpoint_path, checkpoint)
    total_unique = merge_partition_files(partition_files, merge_output)
    print(f"merge_done output={merge_output} unique={total_unique}", flush=True)


if __name__ == "__main__":
    main()
