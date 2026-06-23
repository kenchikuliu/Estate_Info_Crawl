# -*- coding: utf-8 -*-
"""
Backfill Alibaba judicial-auction detail fields through a real browser session.

Why browser-driven:
1. The public detail HTML often redirects to login / anti-bot pages.
2. Once a real logged-in browser opens the item page, lazy tabs such as
   description / notice / warrant / loan can be populated inside the DOM.
3. Many key fields only appear in announcement text or downloadable files.

This script therefore:
- reads Ali index Excel rows
- opens each item page in Selenium Chrome
- warms lazy detail sections inside the page
- extracts page text / hidden coordinates / attachment links
- optionally downloads attachments with the browser cookies
- reuses the JD text parser for structured field extraction
- writes resumable JSONL rows before exporting merged Excel
"""
from __future__ import annotations

import argparse
import json
import math
import re
import sys
import time
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Set, Tuple
from urllib.parse import urljoin, urlparse

import pandas as pd
import requests
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.common.exceptions import TimeoutException, WebDriverException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
try:
    import undetected_chromedriver as uc
except Exception:  # noqa: BLE001
    uc = None

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from utils.jd_detail_parser import (  # noqa: E402
    clean_text,
    extract_attachment_text,
    extract_coordinates,
    extract_labeled_fields,
    extract_region_fields,
    extract_rights_status_text,
    parse_intro_sections,
    parse_survey_table_rows,
    postprocess_structured_fields,
)


DEFAULT_OUTPUT = r"output\阿里法拍房_广东_详情回填.xlsx"
DEFAULT_JSONL = r"output\阿里法拍房_广东_详情回填.jsonl"
DEFAULT_CHECKPOINT = r"output\阿里法拍房_广东_详情回填.checkpoint.json"
DEFAULT_PROFILE = r"output\browser_profiles\ali_sf"
DEFAULT_ATTACHMENT_CACHE = r"output\ali_sf_attachment_cache"
ILLEGAL_XML_RE = re.compile(r"[\x00-\x08\x0B\x0C\x0E-\x1F\x7F-\x9F]")
JSONP_PAYLOAD_RE = re.compile(r"^[^(]*\((.*)\)\s*;?\s*$", re.S)

DETAIL_COLUMNS = [
    "详情抓取时间",
    "详情抓取状态",
    "详情抓取错误",
    "详情页URL",
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
    "是否成交",
    "是否流拍",
    "成交价/获拍价_元",
    "成交时间",
    "当前价_详情_元",
    "起拍价_详情_元",
    "保证金_详情_元",
    "评估价_详情_元",
    "市场价_详情_元",
    "出价次数_详情",
    "围观次数_详情",
    "报名人数_详情",
    "延时次数_详情",
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
    "附件抓取成功数",
    "附件解析成功数",
    "附件本地路径",
    "附件索引原文",
    "详情懒加载URL",
    "详情懒加载错误",
    "标的物介绍文本",
    "竞买公告文本",
    "竞买须知文本",
]


def sanitize_excel_value(value: Any) -> Any:
    if isinstance(value, str):
        return ILLEGAL_XML_RE.sub("", value)
    return value


def sanitize_dataframe_for_excel(df: pd.DataFrame) -> pd.DataFrame:
    return df.apply(lambda col: col.map(sanitize_excel_value))


def latest_ali_index_file() -> Path:
    candidates = sorted(
        PROJECT_ROOT.glob(r"output\**\阿里法拍房_*全量索引*.xlsx"),
        key=lambda path: path.stat().st_mtime,
        reverse=True,
    )
    if candidates:
        return candidates[0]

    candidates = sorted(
        PROJECT_ROOT.glob(r"output\**\阿里法拍房_*索引_*.xlsx"),
        key=lambda path: path.stat().st_mtime,
        reverse=True,
    )
    if candidates:
        return candidates[0]
    return PROJECT_ROOT / r"output\阿里法拍房_广东_全量索引.xlsx"


def resolve_project_path(value: str | Path) -> Path:
    path = Path(value)
    if path.is_absolute():
        return path
    return PROJECT_ROOT / path


def normalize_id(value: Any) -> str:
    if value in ("", None):
        return ""
    text = str(value).strip()
    if not text or text.lower() == "nan":
        return ""
    if text.endswith(".0") and text[:-2].isdigit():
        text = text[:-2]
    match = re.search(r"sf_item/(\d+)\.htm", text)
    if match:
        return match.group(1)
    return text if text.isdigit() else ""


def normalize_url(value: Any, base_url: str = "https://sf-item.taobao.com/") -> str:
    text = clean_text(str(value or ""))
    if not text:
        return ""
    if text.startswith("//"):
        return "https:" + text
    if text.startswith("/"):
        return urljoin(base_url, text)
    return text


def dump_json(value: Any) -> str:
    if value in ("", None, [], {}):
        return ""
    return json.dumps(value, ensure_ascii=False, separators=(",", ":"))


def first_non_empty(*values: Any) -> Any:
    for value in values:
        if value in (None, ""):
            continue
        if isinstance(value, float) and pd.isna(value):
            continue
        return value
    return ""


def compact(value: Any) -> str:
    return re.sub(r"\s+", " ", clean_text(str(value or ""))).strip()


def to_int(value: Any, default: int = 0) -> int:
    try:
        if value in ("", None):
            return default
        return int(float(value))
    except Exception:
        return default


def to_float_or_blank(value: Any) -> Any:
    try:
        if value in ("", None):
            return ""
        number = float(str(value).replace(",", ""))
    except Exception:
        return ""
    if math.isnan(number):
        return ""
    return number


def parse_money_text(text: str) -> Any:
    normalized = compact(text).replace(",", "")
    if not normalized:
        return ""
    match = re.search(r"([0-9]+(?:\.[0-9]+)?)(亿|万)?", normalized)
    if not match:
        return ""
    number = float(match.group(1))
    unit = match.group(2) or ""
    if unit == "万":
        number *= 10_000
    elif unit == "亿":
        number *= 100_000_000
    return number


def parse_text_money(body_text: str, labels: Sequence[str]) -> Any:
    for label in labels:
        match = re.search(
            rf"{label}\s*[：:]?\s*[¥￥]?\s*([0-9][0-9,]*(?:\.[0-9]+)?(?:亿|万)?)",
            body_text,
        )
        if match:
            value = parse_money_text(match.group(1))
            if value != "":
                return value
    return ""


def parse_text_int(body_text: str, labels: Sequence[str]) -> Any:
    for label in labels:
        match = re.search(rf"{label}\s*[：:]?\s*([0-9][0-9,]*)", body_text)
        if match:
            try:
                return int(match.group(1).replace(",", ""))
            except Exception:
                continue
    return ""


def truncate_text(text: str, limit: int) -> str:
    normalized = clean_text(text)
    if limit <= 0 or len(normalized) <= limit:
        return normalized
    return normalized[:limit] + "..."


def safe_name(value: str, default: str = "file") -> str:
    text = clean_text(value)
    text = re.sub(r'[\\/:*?"<>|]+', "_", text)
    text = text.strip(" .")
    return text or default


def ordered_unique_ids(values: Iterable[Any]) -> List[str]:
    result: List[str] = []
    seen: Set[str] = set()
    for value in values:
        item_id = normalize_id(value)
        if not item_id or item_id in seen:
            continue
        seen.add(item_id)
        result.append(item_id)
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
        if normalize_id(row.get("标的物ID")) and row.get("详情抓取状态") == "成功"
    }


def append_jsonl(path: Path, row: Dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("a", encoding="utf-8") as handle:
        handle.write(json.dumps(row, ensure_ascii=False, separators=(",", ":")) + "\n")


def write_checkpoint(path: Path, payload: Dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    copied = dict(payload)
    copied["updated_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    path.write_text(json.dumps(copied, ensure_ascii=False, indent=2), encoding="utf-8")


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


def read_input_rows(path: Path, start: int = 0, limit: int = 0) -> List[Dict[str, Any]]:
    read_kwargs: Dict[str, Any] = {}
    if start > 0:
        read_kwargs["skiprows"] = range(1, start + 1)
    if limit and limit > 0:
        read_kwargs["nrows"] = limit
    df = pd.read_excel(path, **read_kwargs).fillna("").astype(object)
    rows = df.to_dict("records")
    result: List[Dict[str, Any]] = []
    seen: Set[str] = set()
    for row in rows:
        item_id = normalize_id(row.get("标的物ID") or row.get("链接"))
        if not item_id or item_id in seen:
            continue
        row = dict(row)
        row["标的物ID"] = item_id
        seen.add(item_id)
        result.append(row)
    return result


def connect_debug_browser(debug_address: str) -> webdriver.Chrome:
    options = Options()
    options.add_experimental_option("debuggerAddress", debug_address)
    driver = webdriver.Chrome(options=options)
    driver.set_page_load_timeout(60)
    driver.implicitly_wait(2)
    return driver


def create_normal_browser(user_data_dir: Path, headless: bool) -> webdriver.Chrome:
    user_data_dir.mkdir(parents=True, exist_ok=True)
    user_agent = (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/148.0.0.0 Safari/537.36"
    )
    if uc is not None:
        options = uc.ChromeOptions()
        options.add_argument(f"--user-data-dir={user_data_dir}")
        options.add_argument("--disable-popup-blocking")
        options.add_argument("--disable-notifications")
        options.add_argument("--disable-infobars")
        options.add_argument("--disable-extensions")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--window-size=1600,1200")
        options.add_argument(f"--user-agent={user_agent}")
        driver = uc.Chrome(options=options, headless=headless)
        driver.set_page_load_timeout(60)
        driver.implicitly_wait(2)
        return driver

    options = Options()
    options.add_argument(f"--user-data-dir={user_data_dir}")
    options.add_argument("--disable-popup-blocking")
    options.add_argument("--disable-notifications")
    options.add_argument("--disable-infobars")
    options.add_argument("--disable-extensions")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--window-size=1600,1200")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)
    options.add_argument(f"user-agent={user_agent}")
    if headless:
        options.add_argument("--headless=new")
    driver = webdriver.Chrome(options=options)
    driver.set_page_load_timeout(60)
    driver.implicitly_wait(2)
    try:
        driver.execute_cdp_cmd(
            "Page.addScriptToEvaluateOnNewDocument",
            {"source": 'Object.defineProperty(navigator, "webdriver", {get: () => undefined})'},
        )
    except Exception:
        pass
    return driver


def ensure_work_tab(driver: webdriver.Chrome) -> None:
    if not driver.window_handles:
        raise RuntimeError("Browser has no available tabs.")
    driver.switch_to.window(driver.window_handles[-1])


def open_detail_page(driver: webdriver.Chrome, item_id: str, detail_url: str) -> None:
    ensure_work_tab(driver)
    driver.get(detail_url or f"https://sf-item.taobao.com/sf_item/{item_id}.htm")


def page_state(driver: webdriver.Chrome) -> Dict[str, Any]:
    return driver.execute_script(
        """
        const bodyText = document.body ? (document.body.innerText || '') : '';
        const title = document.title || '';
        const href = location.href || '';
        return {
          url: href,
          title: title,
          is_login: /login\\.taobao\\.com/.test(href) || title.includes('登录'),
          is_captcha: bodyText.includes('验证码') || bodyText.includes('滑动验证') || bodyText.includes('请完成验证'),
          ready: !!document.querySelector('#J_ItemId') || !!document.querySelector('#J_desc') || !!document.querySelector('#J_DetailTabMenu'),
          body_text: bodyText.slice(0, 1200)
        };
        """
    )


def wait_for_detail_ready(driver: webdriver.Chrome, timeout: int, interactive: bool) -> Dict[str, Any]:
    deadline = time.time() + timeout
    last_state: Dict[str, Any] = {}
    login_prompt_sent = False
    captcha_prompt_sent = False
    while time.time() < deadline:
        try:
            last_state = page_state(driver)
        except WebDriverException:
            time.sleep(1)
            continue
        if last_state.get("ready"):
            return last_state
        if interactive and last_state.get("is_login") and not login_prompt_sent:
            print("ali_detail_waiting login required; please finish Taobao login in the opened browser window.", flush=True)
            login_prompt_sent = True
        if interactive and last_state.get("is_captcha") and not captcha_prompt_sent:
            print("ali_detail_waiting captcha required; please finish the slider verification in the opened browser window.", flush=True)
            captcha_prompt_sent = True
        time.sleep(2)
    marker = "login_required" if last_state.get("is_login") else "captcha_required" if last_state.get("is_captcha") else "not_ready"
    raise RuntimeError(
        "Ali detail page not ready. "
        f"marker={marker} "
        f"last_url={last_state.get('url', '')} "
        f"login={last_state.get('is_login', False)} captcha={last_state.get('is_captcha', False)}"
    )


def warm_page_sections(driver: webdriver.Chrome) -> None:
    driver.execute_script("window.scrollTo(0, 0);")
    time.sleep(0.8)

    try:
        anchors = driver.find_elements(By.CSS_SELECTOR, "#J_DetailTabMenu a")
    except Exception:
        anchors = []
    for anchor in anchors:
        try:
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", anchor)
            driver.execute_script("arguments[0].click();", anchor)
            time.sleep(1.0)
        except Exception:
            continue

    try:
        total_height = driver.execute_script("return Math.max(document.body.scrollHeight, document.documentElement.scrollHeight);")
    except Exception:
        total_height = 0
    if isinstance(total_height, (int, float)) and total_height > 0:
        step = 700
        position = 0
        while position <= total_height:
            try:
                driver.execute_script("window.scrollTo(0, arguments[0]);", position)
            except Exception:
                break
            time.sleep(0.6)
            position += step

    for selector in [
        "#J_desc",
        "#J_ItemNotice",
        "#J_NoticeDetail",
        "#J_WarrantContent",
        "#J_LoanContent",
        "#J_DownLoadFirst",
        "#J_DownLoadSecond",
    ]:
        try:
            element = driver.find_element(By.CSS_SELECTOR, selector)
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", element)
            time.sleep(0.9)
        except Exception:
            continue
    try:
        driver.execute_script("window.scrollTo(0, 0);")
    except Exception:
        pass
    time.sleep(0.5)


def collect_page_payload(driver: webdriver.Chrome) -> Dict[str, Any]:
    return driver.execute_script(
        """
        function txt(sel) {
          const el = document.querySelector(sel);
          return el ? (el.innerText || el.textContent || '').trim() : '';
        }
        function html(sel) {
          const el = document.querySelector(sel);
          return el ? (el.innerHTML || '') : '';
        }
        function val(sel) {
          const el = document.querySelector(sel);
          return el ? (el.value || el.getAttribute('value') || '').trim() : '';
        }
        function attr(sel, name) {
          const el = document.querySelector(sel);
          return el ? (el.getAttribute(name) || '').trim() : '';
        }
        function links(sel) {
          return Array.from(document.querySelectorAll(sel)).map((el) => ({
            text: (el.innerText || el.textContent || '').trim(),
            href: el.href || el.getAttribute('href') || ''
          })).filter((item) => item.text || item.href);
        }
        function firstText(selectors) {
          for (const selector of selectors) {
            const value = txt(selector);
            if (value) return value;
          }
          return '';
        }
        return {
          url: location.href || '',
          title: document.title || '',
          body_text: document.body ? (document.body.innerText || '').trim() : '',
          title_text: firstText(['h1', '.pm-main h1', '.title h1', '.pm-title', '.main-title']),
          desc_text: txt('#J_desc'),
          desc_html: html('#J_desc'),
          desc_from: attr('#J_desc', 'data-from'),
          item_notice_text: txt('#J_ItemNotice'),
          item_notice_html: html('#J_ItemNotice'),
          item_notice_from: attr('#J_ItemNotice', 'data-from'),
          notice_detail_text: txt('#J_NoticeDetail'),
          notice_detail_html: html('#J_NoticeDetail'),
          notice_detail_from: attr('#J_NoticeDetail', 'data-from'),
          loan_text: txt('#J_LoanContent'),
          loan_html: html('#J_LoanContent'),
          loan_from: attr('#J_LoanContent', 'data-from'),
          warrant_text: txt('#J_WarrantContent'),
          warrant_html: html('#J_WarrantContent'),
          warrant_from: attr('#J_WarrantContent', 'data-from'),
          coordinate: val('#J_Coordinate'),
          province_name: val('#J_ProvinceName'),
          city_name: val('#J_CityName'),
          area_name: val('#J_AreaName'),
          item_id: val('#J_ItemId'),
          version: val('#J_Version'),
          download_first: links('#J_DownLoadFirst a'),
          download_first_from: attr('#J_DownLoadFirst', 'data-from'),
          download_first_url: attr('#J_DownLoadFirst', 'dowload-url'),
          download_second: links('#J_DownLoadSecond a'),
          download_second_from: attr('#J_DownLoadSecond', 'data-from'),
          download_second_url: attr('#J_DownLoadSecond', 'dowload-url')
        };
        """
    )


def html_table_rows(html: str) -> List[List[str]]:
    if not clean_text(html):
        return []
    rows: List[List[str]] = []
    soup = BeautifulSoup(html, "html.parser")
    for table in soup.find_all("table"):
        for tr in table.find_all("tr"):
            cells = [clean_text(cell.get_text(" ", strip=True)) for cell in tr.find_all(["th", "td"])]
            cells = [cell for cell in cells if cell]
            if cells:
                rows.append(cells)
    return rows


def merge_prefer_first(*mappings: Dict[str, Any]) -> Dict[str, str]:
    result: Dict[str, str] = {}
    for mapping in mappings:
        if not isinstance(mapping, dict):
            continue
        for key, value in mapping.items():
            text = clean_text(str(value or ""))
            if text and not result.get(key):
                result[key] = text
    return result


def page_requests_session(driver: webdriver.Chrome) -> requests.Session:
    session = requests.Session()
    session.headers.update(
        {
            "User-Agent": driver.execute_script("return navigator.userAgent;") or (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/148.0.0.0 Safari/537.36"
            ),
            "Accept": "*/*",
            "Referer": driver.current_url,
        }
    )
    for cookie in driver.get_cookies():
        try:
            session.cookies.set(
                cookie.get("name", ""),
                cookie.get("value", ""),
                domain=cookie.get("domain"),
                path=cookie.get("path", "/"),
            )
        except Exception:
            continue
    return session


def normalize_lazy_url(value: Any, current_url: str) -> str:
    return normalize_url(value, base_url=current_url or "https://sf-item.taobao.com/")


def parse_jsonp_response(text: str) -> Any:
    payload = clean_text(text)
    if not payload:
        return {}
    match = JSONP_PAYLOAD_RE.match(payload)
    if match:
        payload = match.group(1)
    try:
        return json.loads(payload)
    except Exception:
        return {}


def lazy_content_text(value: Any) -> str:
    if isinstance(value, str):
        return value
    if not isinstance(value, dict):
        return ""
    for key in ["content", "desc", "html", "data", "result"]:
        item = value.get(key)
        if isinstance(item, str) and clean_text(item):
            return item
    return ""


def lazy_attaches(value: Any, download_url: str) -> List[Dict[str, str]]:
    if not isinstance(value, dict):
        return []
    raw = value.get("attaches") or value.get("attachments") or []
    if not isinstance(raw, list):
        return []
    result: List[Dict[str, str]] = []
    base = normalize_url(download_url, base_url="https://sf-item.taobao.com/")
    for item in raw:
        if not isinstance(item, dict):
            continue
        title = clean_text(str(item.get("title") or item.get("name") or ""))
        href = normalize_url(item.get("href") or item.get("url") or "")
        attach_id = clean_text(str(item.get("id") or item.get("attach_id") or ""))
        if not href and base and attach_id:
            href = base + ("&" if "?" in base else "?") + "attach_id=" + attach_id
        if title or href:
            result.append({"text": title, "href": href})
    return result


def fetch_lazy_jsonp(session: requests.Session, url: str) -> Any:
    response = session.get(url, timeout=30, allow_redirects=True)
    response.raise_for_status()
    return parse_jsonp_response(response.text)


def enrich_payload_from_lazy_sources(driver: webdriver.Chrome, page_payload: Dict[str, Any]) -> Dict[str, Any]:
    session = page_requests_session(driver)
    current_url = str(page_payload.get("url") or driver.current_url or "")
    source_map = {
        "desc": ("desc_from", "desc_text", "desc_html"),
        "item_notice": ("item_notice_from", "item_notice_text", "item_notice_html"),
        "notice_detail": ("notice_detail_from", "notice_detail_text", "notice_detail_html"),
        "loan": ("loan_from", "loan_text", "loan_html"),
        "warrant": ("warrant_from", "warrant_text", "warrant_html"),
    }
    errors: List[str] = []
    fetched_urls: Dict[str, str] = {}

    for label, (url_key, text_key, html_key) in source_map.items():
        lazy_url = normalize_lazy_url(page_payload.get(url_key), current_url)
        if not lazy_url or clean_text(page_payload.get(text_key)):
            continue
        try:
            payload = fetch_lazy_jsonp(session, lazy_url)
            html = lazy_content_text(payload)
            if html:
                soup = BeautifulSoup(html, "html.parser")
                page_payload[text_key] = clean_text(soup.get_text("\n", strip=True))
                page_payload[html_key] = html
                fetched_urls[label] = lazy_url
        except Exception as exc:  # noqa: BLE001
            errors.append(f"{label}:{exc}")

    for prefix in ["download_first", "download_second"]:
        lazy_url = normalize_lazy_url(page_payload.get(f"{prefix}_from"), current_url)
        if not lazy_url or page_payload.get(prefix):
            continue
        try:
            payload = fetch_lazy_jsonp(session, lazy_url)
            attaches = lazy_attaches(payload, str(page_payload.get(f"{prefix}_url") or ""))
            if attaches:
                page_payload[prefix] = attaches
                fetched_urls[prefix] = lazy_url
        except Exception as exc:  # noqa: BLE001
            errors.append(f"{prefix}:{exc}")

    page_payload["lazy_fetched_urls"] = dump_json(fetched_urls)
    page_payload["lazy_fetch_errors"] = "；".join(errors)
    return page_payload


def looks_like_html(path: Path) -> bool:
    try:
        prefix = path.read_bytes()[:512]
    except Exception:
        return False
    lowered = prefix.lower()
    return b"<html" in lowered or b"<!doctype html" in lowered


def attachment_file_name(item_id: str, index: int, title: str, href: str) -> str:
    parsed = urlparse(href or "")
    suffix = Path(parsed.path).suffix
    base = safe_name(title or f"attachment_{index}")
    if suffix and not base.lower().endswith(suffix.lower()):
        return base + suffix
    return base


def download_attachments(
    driver: webdriver.Chrome,
    item_id: str,
    attachments: Sequence[Dict[str, str]],
    cache_root: Path,
    limit: int,
    timeout: int,
) -> Tuple[List[Dict[str, Any]], List[str]]:
    session = page_requests_session(driver)
    item_dir = cache_root / item_id
    item_dir.mkdir(parents=True, exist_ok=True)
    downloaded: List[Dict[str, Any]] = []
    texts: List[str] = []
    seen_urls: Set[str] = set()

    for index, attachment in enumerate(attachments, start=1):
        href = normalize_url(attachment.get("href"))
        title = clean_text(attachment.get("text", ""))
        if not href or href in seen_urls:
            continue
        seen_urls.add(href)
        if limit > 0 and len(downloaded) >= limit:
            break

        file_name = attachment_file_name(item_id, index, title, href)
        local_path = item_dir / file_name
        status = "已存在"
        error = ""
        if not local_path.exists():
            try:
                response = session.get(href, timeout=timeout, allow_redirects=True)
                response.raise_for_status()
                local_path.write_bytes(response.content)
                if looks_like_html(local_path):
                    local_path.unlink(missing_ok=True)
                    raise RuntimeError("attachment returned html page")
                status = "下载成功"
            except Exception as exc:  # noqa: BLE001
                error = str(exc)
                status = "下载失败"
                local_path.unlink(missing_ok=True)
        extracted_text = ""
        if local_path.exists():
            try:
                extracted_text = extract_attachment_text(str(local_path))
            except Exception as exc:  # noqa: BLE001
                error = str(exc)
        if extracted_text:
            texts.append(extracted_text)
        downloaded.append(
            {
                "text": title,
                "href": href,
                "local_path": str(local_path) if local_path.exists() else "",
                "download_status": status,
                "error": error,
                "parsed_text_length": len(clean_text(extracted_text)),
            }
        )
    return downloaded, texts


def parse_finance_fields(loan_text: str, body_text: str) -> Dict[str, str]:
    text = clean_text("\n".join([loan_text, body_text]))
    if not text:
        return {
            "是否有金融服务_详情": "",
            "金融机构": "",
            "最高可贷比例": "",
            "参考利率": "",
            "金融其他费用": "",
            "金融服务原文": "",
        }

    def pick(patterns: Sequence[str]) -> str:
        for pattern in patterns:
            match = re.search(pattern, text)
            if match:
                return clean_text(match.group(1))
        return ""

    loan_flag = "是" if ("贷款" in text or "金融服务" in text or "利率" in text) else ""
    return {
        "是否有金融服务_详情": loan_flag,
        "金融机构": pick(
            [
                r"(?:金融机构|贷款机构|服务机构)[:：]?\s*([^\n]+)",
                r"(?:机构名称)[:：]?\s*([^\n]+)",
            ]
        ),
        "最高可贷比例": pick(
            [
                r"(?:最高可贷比例|可贷比例|最高贷款比例)[:：]?\s*([^\n]+)",
            ]
        ),
        "参考利率": pick(
            [
                r"(?:参考利率|贷款利率|年利率|月利率)[:：]?\s*([^\n]+)",
            ]
        ),
        "金融其他费用": pick(
            [
                r"(?:服务费|担保费|评估费|其他费用)[:：]?\s*([^\n]+)",
            ]
        ),
        "金融服务原文": text if loan_flag else "",
    }


def parse_page_metrics(body_text: str, page_title: str, base_row: Dict[str, Any]) -> Dict[str, Any]:
    status_text = ""
    for status in ["已成交", "已结束", "进行中", "即将开始", "流拍", "中止", "撤回", "变卖"]:
        if status in body_text or status in page_title:
            status_text = status
            break
    if not status_text:
        status_text = clean_text(str(base_row.get("竞价状态") or ""))

    sold = "已成交" in status_text
    unsold = "流拍" in status_text
    current_price = first_non_empty(
        parse_text_money(body_text, ["当前价"]),
        to_float_or_blank(base_row.get("当前价_元")),
    )
    start_price = first_non_empty(
        parse_text_money(body_text, ["起拍价", "起始价", "变卖价"]),
        to_float_or_blank(base_row.get("起拍价_元")),
    )
    deposit = parse_text_money(body_text, ["保证金"])
    assess_price = first_non_empty(
        parse_text_money(body_text, ["评估价"]),
        to_float_or_blank(base_row.get("评估价_元")),
    )
    market_price = first_non_empty(
        parse_text_money(body_text, ["市场价"]),
        to_float_or_blank(base_row.get("市场价_元")),
    )
    bid_count = first_non_empty(
        parse_text_int(body_text, ["出价次数", "出价"]),
        to_int(base_row.get("出价次数"), 0) or "",
    )
    viewer_count = first_non_empty(
        parse_text_int(body_text, ["围观次数", "围观"]),
        to_int(base_row.get("围观次数"), 0) or "",
    )
    apply_count = first_non_empty(
        parse_text_int(body_text, ["报名人数", "报名"]),
        to_int(base_row.get("报名人数"), 0) or "",
    )
    delay_count = first_non_empty(
        parse_text_int(body_text, ["延时次数", "延时"]),
        to_int(base_row.get("延时次数"), 0) or "",
    )

    time_match = None
    for pattern in [
        r"(?:成交时间|结束时间|竞价结束时间)[:：]?\s*([0-9]{4}[-/年][0-9]{1,2}[-/月][0-9]{1,2}(?:日)?\s*[0-9]{0,2}:?[0-9]{0,2}:?[0-9]{0,2})",
        r"(?:结束)\s*([0-9]{4}[-/年][0-9]{1,2}[-/月][0-9]{1,2}(?:日)?\s*[0-9]{0,2}:?[0-9]{0,2}:?[0-9]{0,2})",
    ]:
        time_match = re.search(pattern, body_text)
        if time_match:
            break
    deal_time = clean_text(time_match.group(1)) if time_match and sold else clean_text(str(base_row.get("成交时间") or ""))

    deal_price = current_price if sold and current_price != "" else to_float_or_blank(base_row.get("成交价/获拍价_元"))
    return {
        "拍卖状态_详情": status_text,
        "是否成交": "是" if sold else clean_text(str(base_row.get("是否成交") or "否")),
        "是否流拍": "是" if unsold else clean_text(str(base_row.get("是否流拍") or "否")),
        "成交价/获拍价_元": deal_price,
        "成交时间": deal_time,
        "当前价_详情_元": current_price,
        "起拍价_详情_元": start_price,
        "保证金_详情_元": deposit,
        "评估价_详情_元": assess_price,
        "市场价_详情_元": market_price,
        "出价次数_详情": bid_count,
        "围观次数_详情": viewer_count,
        "报名人数_详情": apply_count,
        "延时次数_详情": delay_count,
    }


def address_fields(page_payload: Dict[str, Any], structured_fields: Dict[str, str], base_row: Dict[str, Any]) -> Dict[str, str]:
    province = clean_text(page_payload.get("province_name", ""))
    city = clean_text(page_payload.get("city_name", ""))
    area = clean_text(page_payload.get("area_name", ""))
    title = clean_text(page_payload.get("title_text", "")) or clean_text(base_row.get("标题", ""))
    detail_address = clean_text(structured_fields.get("标的物名称", "")) or title
    region = extract_region_fields(detail_address)
    return {
        "详情地址": detail_address,
        "详情省": province or region.get("省", ""),
        "详情市": city or region.get("市", ""),
        "详情区县": area or region.get("区县", ""),
    }


def export_excel(input_path: Path, jsonl_path: Path, output_path: Path) -> int:
    base = pd.read_excel(input_path).fillna("").astype(object)
    if "标的物ID" in base.columns:
        base["标的物ID"] = base["标的物ID"].map(normalize_id)
    else:
        raise ValueError("Input Excel must contain 标的物ID.")

    detail_rows = read_jsonl(jsonl_path)
    if not detail_rows:
        clean_base = sanitize_dataframe_for_excel(base)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        with pd.ExcelWriter(
            output_path,
            engine="xlsxwriter",
            engine_kwargs={"options": {"strings_to_urls": False}},
        ) as writer:
            clean_base.to_excel(writer, index=False)
        return len(clean_base)

    detail_df = pd.DataFrame(detail_rows).fillna("").astype(object)
    detail_df["标的物ID"] = detail_df["标的物ID"].map(normalize_id)
    detail_df["_order"] = range(len(detail_df))
    detail_df = detail_df.sort_values("_order").drop_duplicates("标的物ID", keep="last").drop(columns=["_order"])
    drop_cols = [col for col in DETAIL_COLUMNS if col in base.columns]
    base = base.drop(columns=drop_cols, errors="ignore")
    overlap = [col for col in detail_df.columns if col in base.columns and col != "标的物ID"]
    detail_df = detail_df.drop(columns=overlap, errors="ignore")
    merged = base.merge(detail_df, on="标的物ID", how="left")
    clean_df = sanitize_dataframe_for_excel(merged)

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


def build_detail_row(
    base_row: Dict[str, Any],
    page_payload: Dict[str, Any],
    attachments: Sequence[Dict[str, Any]],
    attachment_texts: Sequence[str],
    raw_text_limit: int,
) -> Dict[str, Any]:
    section_texts = [
        page_payload.get("desc_text", ""),
        page_payload.get("item_notice_text", ""),
        page_payload.get("notice_detail_text", ""),
        page_payload.get("loan_text", ""),
        page_payload.get("warrant_text", ""),
    ]
    combined_text = clean_text("\n".join(text for text in section_texts if clean_text(text)))
    attachment_text = clean_text("\n".join(text for text in attachment_texts if clean_text(text)))
    all_text = clean_text("\n".join([combined_text, attachment_text]))

    table_rows: List[List[str]] = []
    for html in [
        page_payload.get("desc_html", ""),
        page_payload.get("item_notice_html", ""),
        page_payload.get("notice_detail_html", ""),
        page_payload.get("loan_html", ""),
    ]:
        table_rows.extend(html_table_rows(str(html or "")))

    table_fields = parse_survey_table_rows(table_rows)
    text_fields = extract_labeled_fields(all_text)
    intro_sections = parse_intro_sections(all_text)
    intro_fields = {
        "权证情况": intro_sections.get("权证情况", ""),
        "被执行人": intro_sections.get("拍品所有人", ""),
        "权利限制状况及抵押状况": intro_sections.get("权利限制状况及抵押状况", ""),
        "房屋权属状况": intro_sections.get("房屋权属状况", ""),
        "土地权属状况": intro_sections.get("土地权属状况", ""),
    }
    rights_fields = extract_rights_status_text(all_text)
    structured_fields = postprocess_structured_fields(
        merge_prefer_first(table_fields, text_fields, intro_fields, rights_fields)
    )

    coordinate = clean_text(page_payload.get("coordinate", ""))
    lng, lat = extract_coordinates(coordinate)
    finance_fields = parse_finance_fields(
        clean_text(page_payload.get("loan_text", "")),
        clean_text(page_payload.get("body_text", "")),
    )
    page_metrics = parse_page_metrics(
        clean_text(page_payload.get("body_text", "")),
        clean_text(page_payload.get("title", "")),
        base_row,
    )
    addr_fields = address_fields(page_payload, structured_fields, base_row)
    owner = first_non_empty(
        structured_fields.get("被执行人"),
        clean_text(base_row.get("标的所有人", "")),
    )
    detail_title = first_non_empty(page_payload.get("title_text"), page_payload.get("title"), base_row.get("标题"))
    attachment_success = len([item for item in attachments if clean_text(item.get("local_path", ""))])
    parsed_success = len([item for item in attachments if to_int(item.get("parsed_text_length"), 0) > 0])

    row = {
        "标的物ID": normalize_id(base_row.get("标的物ID")),
        "详情抓取时间": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "详情抓取状态": "成功",
        "详情抓取错误": "",
        "详情页URL": normalize_url(page_payload.get("url", "")),
        "详情标题": clean_text(str(detail_title or "")),
        "标的名称": clean_text(str(first_non_empty(structured_fields.get("标的物名称"), detail_title) or "")),
        "标的物名称": clean_text(str(first_non_empty(structured_fields.get("标的物名称"), detail_title) or "")),
        "权证情况": structured_fields.get("权证情况", ""),
        "标的所有人": owner,
        "被执行人/标的所有人": owner,
        **addr_fields,
        "经度": lng if lng is not None else "",
        "纬度": lat if lat is not None else "",
        "经纬度来源": "阿里详情页坐标" if lng is not None and lat is not None else "",
        "法院/处置机构": clean_text(str(base_row.get("法院", "") or base_row.get("courtName", "") or "")),
        **page_metrics,
        **finance_fields,
        "建筑面积": structured_fields.get("建筑面积", ""),
        "房屋用途": structured_fields.get("房屋用途", ""),
        "房屋类型": structured_fields.get("房屋类型", ""),
        "所在层": structured_fields.get("所在层", ""),
        "总层数": structured_fields.get("总层数", ""),
        "竣工时间": structured_fields.get("竣工时间", ""),
        "购买时间": structured_fields.get("购买时间", ""),
        "土地性质": structured_fields.get("土地性质", ""),
        "土地用途": structured_fields.get("土地用途", ""),
        "使用期限": structured_fields.get("使用期限", ""),
        "权利来源": structured_fields.get("权利来源", ""),
        "所有权来源": structured_fields.get("所有权来源", ""),
        "钥匙/占用情况": structured_fields.get("钥匙", ""),
        "腾空情况": structured_fields.get("腾空情况", ""),
        "户籍/工商注册": structured_fields.get("户籍注册", ""),
        "欠费情况": structured_fields.get("欠费情况", ""),
        "提供文件": structured_fields.get("提供文件", ""),
        "权利限制状况及抵押状况": structured_fields.get("权利限制状况及抵押状况", ""),
        "房屋权属状况": structured_fields.get("房屋权属状况", ""),
        "土地权属状况": structured_fields.get("土地权属状况", ""),
        "附件数量": len(attachments),
        "附件名称": join_unique(item.get("text") for item in attachments),
        "附件链接": join_unique(item.get("href") for item in attachments),
        "附件抓取成功数": attachment_success,
        "附件解析成功数": parsed_success,
        "附件本地路径": join_unique(item.get("local_path") for item in attachments),
        "附件索引原文": dump_json(list(attachments)),
        "详情懒加载URL": page_payload.get("lazy_fetched_urls", ""),
        "详情懒加载错误": page_payload.get("lazy_fetch_errors", ""),
        "标的物介绍文本": truncate_text(page_payload.get("desc_text", ""), raw_text_limit),
        "竞买公告文本": truncate_text(
            clean_text("\n".join([page_payload.get("item_notice_text", ""), page_payload.get("notice_detail_text", "")])),
            raw_text_limit,
        ),
        "竞买须知文本": truncate_text(page_payload.get("warrant_text", ""), raw_text_limit),
    }
    return row


def failure_row(base_row: Dict[str, Any], message: str) -> Dict[str, Any]:
    status = "需登录" if "marker=login_required" in message else "需验证" if "marker=captcha_required" in message else "失败"
    return {
        "标的物ID": normalize_id(base_row.get("标的物ID")),
        "详情抓取时间": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "详情抓取状态": status,
        "详情抓取错误": clean_text(message),
        "详情页URL": normalize_url(base_row.get("链接", "")) or f"https://sf-item.taobao.com/sf_item/{normalize_id(base_row.get('标的物ID'))}.htm",
    }


def fetch_detail_for_row(
    driver: webdriver.Chrome,
    base_row: Dict[str, Any],
    manual_wait_seconds: int,
    interactive: bool,
    raw_text_limit: int,
    include_attachments: bool,
    attachment_cache_root: Path,
    attachment_limit: int,
    attachment_timeout: int,
) -> Dict[str, Any]:
    item_id = normalize_id(base_row.get("标的物ID"))
    if not item_id:
        return failure_row(base_row, "missing item id")
    detail_url = normalize_url(base_row.get("链接", "")) or f"https://sf-item.taobao.com/sf_item/{item_id}.htm"

    try:
        open_detail_page(driver, item_id, detail_url)
        wait_for_detail_ready(driver, manual_wait_seconds, interactive=interactive)
        warm_page_sections(driver)
        page_payload = collect_page_payload(driver)
        page_payload = enrich_payload_from_lazy_sources(driver, page_payload)
        attachments = []
        attachment_texts: List[str] = []
        if include_attachments:
            raw_attachments = list(page_payload.get("download_first", []) or []) + list(page_payload.get("download_second", []) or [])
            attachments, attachment_texts = download_attachments(
                driver=driver,
                item_id=item_id,
                attachments=raw_attachments,
                cache_root=attachment_cache_root,
                limit=attachment_limit,
                timeout=attachment_timeout,
            )
        return build_detail_row(
            base_row=base_row,
            page_payload=page_payload,
            attachments=attachments,
            attachment_texts=attachment_texts,
            raw_text_limit=raw_text_limit,
        )
    except Exception as exc:  # noqa: BLE001
        return failure_row(base_row, str(exc))


def run_login_check(
    *,
    input_path: Path,
    start: int,
    user_data_dir: Path,
    browser_mode: str,
    debug_address: str,
    manual_wait_seconds: int,
) -> None:
    rows = read_input_rows(input_path, start=start, limit=1)
    if not rows:
        raise RuntimeError("login check input has no valid item rows")
    item_id = normalize_id(rows[0].get("标的物ID"))
    detail_url = normalize_url(rows[0].get("链接", "")) or f"https://sf-item.taobao.com/sf_item/{item_id}.htm"
    if browser_mode == "debug":
        driver = connect_debug_browser(debug_address)
    else:
        driver = create_normal_browser(user_data_dir, headless=False)
    try:
        print(f"ali_login_check_open item_id={item_id} url={detail_url}", flush=True)
        open_detail_page(driver, item_id, detail_url)
        state = wait_for_detail_ready(driver, manual_wait_seconds, interactive=True)
        payload = collect_page_payload(driver)
        payload = enrich_payload_from_lazy_sources(driver, payload)
        print(
            "ali_login_check_ready "
            f"url={state.get('url', '')} "
            f"desc_len={len(clean_text(payload.get('desc_text', '')))} "
            f"notice_len={len(clean_text(payload.get('item_notice_text', '')) + clean_text(payload.get('notice_detail_text', '')))} "
            f"lazy_urls={payload.get('lazy_fetched_urls', '')}",
            flush=True,
        )
    finally:
        try:
            driver.quit()
        except Exception:
            pass


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Backfill Alibaba judicial-auction detail fields through browser DOM.")
    parser.add_argument("--input", default=str(latest_ali_index_file()))
    parser.add_argument("--output", default=DEFAULT_OUTPUT)
    parser.add_argument("--jsonl-output", default=DEFAULT_JSONL)
    parser.add_argument("--checkpoint", default=DEFAULT_CHECKPOINT)
    parser.add_argument("--browser-mode", choices=["normal", "debug"], default="normal")
    parser.add_argument("--debug-address", default="127.0.0.1:9222")
    parser.add_argument("--user-data-dir", default=DEFAULT_PROFILE)
    parser.add_argument("--headless", action="store_true")
    parser.add_argument("--manual-wait-seconds", type=int, default=180)
    parser.add_argument("--start", type=int, default=0)
    parser.add_argument("--limit", type=int, default=0)
    parser.add_argument("--save-every", type=int, default=20)
    parser.add_argument("--raw-text-limit", type=int, default=4000)
    parser.add_argument("--attachment-cache", default=DEFAULT_ATTACHMENT_CACHE)
    parser.add_argument("--attachment-limit", type=int, default=8)
    parser.add_argument("--attachment-timeout", type=int, default=60)
    parser.add_argument("--skip-attachments", action="store_true")
    parser.add_argument("--no-resume", action="store_true")
    parser.add_argument("--export-only", action="store_true")
    parser.add_argument("--login-check", action="store_true", help="Open one Ali detail page and wait for manual login/captcha completion.")
    parser.add_argument("--stop-on-auth-block", action="store_true", help="Stop immediately without appending the row when login/captcha blocks detail loading.")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    input_path = resolve_project_path(args.input)
    output_path = resolve_project_path(args.output)
    jsonl_path = resolve_project_path(args.jsonl_output)
    checkpoint_path = resolve_project_path(args.checkpoint)
    attachment_cache_root = resolve_project_path(args.attachment_cache)

    if args.export_only:
        rows = export_excel(input_path, jsonl_path, output_path)
        print(f"ali_detail_export_done output={output_path} rows={rows}", flush=True)
        return

    if args.login_check:
        run_login_check(
            input_path=input_path,
            start=args.start,
            user_data_dir=resolve_project_path(args.user_data_dir),
            browser_mode=args.browser_mode,
            debug_address=args.debug_address,
            manual_wait_seconds=args.manual_wait_seconds,
        )
        return

    input_rows = read_input_rows(input_path, start=args.start, limit=args.limit)
    selected = input_rows
    completed = load_completed_ids(jsonl_path) if not args.no_resume else set()
    todo = [row for row in selected if normalize_id(row.get("标的物ID")) not in completed]
    print(
        f"ali_detail_plan input_rows={len(input_rows)} selected={len(selected)} "
        f"completed_existing={len(completed)} todo={len(todo)} browser_mode={args.browser_mode}",
        flush=True,
    )
    if not todo:
        rows = export_excel(input_path, jsonl_path, output_path)
        print(f"ali_detail_done output={output_path} rows={rows}", flush=True)
        return

    if args.browser_mode == "debug":
        driver = connect_debug_browser(args.debug_address)
    else:
        driver = create_normal_browser(resolve_project_path(args.user_data_dir), headless=args.headless)

    done = 0
    failed = 0
    started = time.time()
    try:
        for index, row in enumerate(todo, start=1):
            detail_row = fetch_detail_for_row(
                driver=driver,
                base_row=row,
                manual_wait_seconds=args.manual_wait_seconds,
                interactive=not args.headless and args.browser_mode == "normal",
                raw_text_limit=args.raw_text_limit,
                include_attachments=not args.skip_attachments,
                attachment_cache_root=attachment_cache_root,
                attachment_limit=args.attachment_limit,
                attachment_timeout=args.attachment_timeout,
            )
            if args.stop_on_auth_block and detail_row.get("详情抓取状态") in {"需登录", "需验证"}:
                checkpoint = {
                    "input": str(input_path),
                    "output": str(output_path),
                    "jsonl_output": str(jsonl_path),
                    "done": done,
                    "failed": failed,
                    "processed": index - 1,
                    "total": len(todo),
                    "elapsed_seconds": round(time.time() - started, 1),
                    "rows_exported": 0,
                    "last_item_id": detail_row.get("标的物ID", ""),
                    "last_status": detail_row.get("详情抓取状态", ""),
                    "auth_block_error": detail_row.get("详情抓取错误", ""),
                }
                write_checkpoint(checkpoint_path, checkpoint)
                print(
                    f"ali_detail_auth_block status={detail_row.get('详情抓取状态')} "
                    f"item_id={detail_row.get('标的物ID', '')}; run --login-check then resume.",
                    flush=True,
                )
                break
            append_jsonl(jsonl_path, detail_row)
            if detail_row.get("详情抓取状态") == "成功":
                done += 1
            else:
                failed += 1

            if index == 1 or index % max(1, args.save_every) == 0 or index == len(todo):
                rows = export_excel(input_path, jsonl_path, output_path)
                checkpoint = {
                    "input": str(input_path),
                    "output": str(output_path),
                    "jsonl_output": str(jsonl_path),
                    "done": done,
                    "failed": failed,
                    "processed": index,
                    "total": len(todo),
                    "elapsed_seconds": round(time.time() - started, 1),
                    "rows_exported": rows,
                    "last_item_id": detail_row.get("标的物ID", ""),
                    "last_status": detail_row.get("详情抓取状态", ""),
                }
                write_checkpoint(checkpoint_path, checkpoint)
                print(
                    f"ali_detail_progress processed={index}/{len(todo)} done={done} failed={failed} "
                    f"last_id={detail_row.get('标的物ID', '')} last_status={detail_row.get('详情抓取状态', '')}",
                    flush=True,
                )
    finally:
        try:
            driver.quit()
        except Exception:
            pass

    rows = export_excel(input_path, jsonl_path, output_path)
    print(f"ali_detail_done output={output_path} rows={rows} jsonl={jsonl_path}", flush=True)


if __name__ == "__main__":
    main()
