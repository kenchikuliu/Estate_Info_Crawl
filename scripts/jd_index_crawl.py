# -*- coding: utf-8 -*-
"""
Fast JD auction list index crawler.

This crawler only scans the search result list. It does not open detail pages,
download attachments, take screenshots, or enrich POI data.
"""
import argparse
import json
import re
import time
from datetime import datetime
from pathlib import Path
from typing import Dict, List

import pandas as pd
from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait


BASE_URL = "https://pmsearch.jd.com/?publishSource=7&childrenCateId=12728"


def clean_text(value: str) -> str:
    value = value or ""
    value = value.replace("\r", "\n").replace("\u3000", " ")
    value = re.sub(r"[ \t]+", " ", value)
    value = re.sub(r"\n{3,}", "\n\n", value)
    return value.strip()


def compact(value: str) -> str:
    return re.sub(r"\s+", " ", clean_text(value)).strip()


def parse_card_text(raw_text: str) -> Dict[str, str]:
    lines = [line.strip() for line in clean_text(raw_text).splitlines() if line.strip()]
    joined = "\n".join(lines)
    result: Dict[str, str] = {
        "标题": lines[0] if lines else "",
        "城市": "",
        "当前价": "",
        "评估价": "",
        "市场价": "",
        "出价次数": "",
        "竞价状态": "",
        "剩余时间": "",
        "是否有金融服务": "是" if ("一键贷款" in joined or "金融服务" in joined) else "否",
        "列表原文": joined,
    }

    for line in lines[1:5]:
        if re.fullmatch(r"[\u4e00-\u9fa5]{2,12}市", line):
            result["城市"] = line
            break

    current_match = re.search(r"当前价[:：]?\s*¥?\s*([0-9.]+万?)", joined)
    if current_match:
        result["当前价"] = current_match.group(1)

    eval_match = re.search(r"评估价[:：]?\s*¥?\s*([0-9.]+万?)", joined)
    if eval_match:
        result["评估价"] = eval_match.group(1)

    market_match = re.search(r"市场价[:：]?\s*¥?\s*([0-9.]+万?)", joined)
    if market_match:
        result["市场价"] = market_match.group(1)

    bid_match = re.search(r"(\d+)次出价", joined)
    if bid_match:
        result["出价次数"] = bid_match.group(1)

    for status in ["进行中", "即将开始", "已结束", "已成交", "流拍", "中止", "撤回", "变卖"]:
        if status in joined:
            result["竞价状态"] = status
            break

    remaining_match = re.search(r"预计剩余[:：]?([^\n]+)", joined)
    if remaining_match:
        result["剩余时间"] = compact(remaining_match.group(1))

    return result


def connect_debug_browser(debug_address: str) -> webdriver.Chrome:
    options = Options()
    options.add_experimental_option("debuggerAddress", debug_address)
    driver = webdriver.Chrome(options=options)
    driver.set_page_load_timeout(45)
    driver.implicitly_wait(3)
    return driver


def create_normal_browser(headless: bool = False) -> webdriver.Chrome:
    options = Options()
    options.add_argument("--disable-popup-blocking")
    options.add_argument("--disable-notifications")
    options.add_argument("--disable-infobars")
    options.add_argument("--disable-extensions")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--window-size=1920,1080")
    options.add_argument(
        "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/148.0.0.0 Safari/537.36"
    )
    if headless:
        options.add_argument("--headless=new")
    driver = webdriver.Chrome(options=options)
    driver.set_page_load_timeout(45)
    driver.implicitly_wait(3)
    return driver


def open_work_tab(driver: webdriver.Chrome, url: str) -> str:
    original_handle = driver.current_window_handle
    driver.execute_script("window.open('about:blank', '_blank');")
    driver.switch_to.window(driver.window_handles[-1])
    driver.get(url)
    return original_handle


def wait_for_list_page(driver: webdriver.Chrome, timeout: int = 45) -> None:
    WebDriverWait(driver, timeout).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
    WebDriverWait(driver, timeout).until(
        EC.presence_of_all_elements_located((By.XPATH, "//a[contains(@href, 'paimai.jd.com/')]"))
    )
    time.sleep(1.5)


def select_location(driver: webdriver.Chrome, province_label: str, city_label: str = "") -> str:
    location_root = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located(
            (By.XPATH, "//div[contains(@class,'s-location')][.//*[contains(normalize-space(.),'标的物所在地')]]")
        )
    )
    province = location_root.find_element(By.CSS_SELECTOR, "dl.province")
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", location_root)
    time.sleep(0.5)
    driver.execute_script("arguments[0].click();", province)
    time.sleep(1)
    option = province.find_element(By.XPATH, f".//a[normalize-space()='{province_label}']")
    driver.execute_script("arguments[0].click();", option)
    time.sleep(3)

    selected_label = province_label
    if city_label:
        location_root = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located(
                (By.XPATH, "//div[contains(@class,'s-location')][.//*[contains(normalize-space(.),'标的物所在地')]]")
            )
        )
        city = WebDriverWait(location_root, 20).until(
            lambda root: root.find_element(By.CSS_SELECTOR, "dl.city")
        )
        driver.execute_script("arguments[0].click();", city)
        time.sleep(1)
        city_option = city.find_element(By.XPATH, f".//a[normalize-space()='{city_label}']")
        driver.execute_script("arguments[0].click();", city_option)
        time.sleep(3)
        selected_label = city_label

    return selected_label


def get_current_page(driver: webdriver.Chrome) -> int:
    try:
        current = driver.find_element(By.CLASS_NAME, "ui-pager-current").text
        return int(re.search(r"\d+", current).group(0))
    except Exception:
        return 1


def extract_page_items(driver: webdriver.Chrome, province_label: str, city_label: str, page_no: int) -> List[Dict[str, str]]:
    script = r"""
const anchors = Array.from(document.querySelectorAll("a[href*='paimai.jd.com/']"));
const items = [];
const seen = new Set();
for (const a of anchors) {
  const href = (a.href || '').split('?')[0];
  const match = href.match(/^https:\/\/paimai\.jd\.com\/(\d+)\/?$/);
  if (!match || seen.has(href)) continue;
  seen.add(href);
  const text = (a.innerText || a.textContent || '').trim();
  items.push({href, asset_id: match[1], text});
}
return items;
"""
    raw_items = driver.execute_script(script) or []
    rows: List[Dict[str, str]] = []
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    for item in raw_items:
        parsed = parse_card_text(item.get("text", ""))
        text_blob = parsed.get("列表原文", "")
        if city_label and city_label not in text_blob:
            continue
        row = {
            "标的物ID": item.get("asset_id", ""),
            "链接": item.get("href", ""),
            "列表页码": page_no,
            "省份筛选": province_label,
            "城市筛选": city_label,
            "抓取时间": now,
        }
        row.update(parsed)
        rows.append(row)
    return rows


def click_next_page(driver: webdriver.Chrome, current_page: int) -> bool:
    try:
        next_button = driver.find_element(By.CLASS_NAME, "ui-pager-next")
    except Exception:
        return False

    classes = next_button.get_attribute("class") or ""
    if "disable" in classes.lower() or "disabled" in classes.lower():
        return False

    old_url = driver.current_url
    old_page = current_page
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", next_button)
    time.sleep(0.3)
    driver.execute_script("arguments[0].click();", next_button)

    deadline = time.time() + 25
    while time.time() < deadline:
        time.sleep(1)
        new_page = get_current_page(driver)
        if new_page != old_page or driver.current_url != old_url:
            time.sleep(1.5)
            return True
    return False


def advance_to_page(driver: webdriver.Chrome, target_page: int) -> int:
    current_page = get_current_page(driver)
    while current_page < target_page:
        if not click_next_page(driver, current_page):
            return get_current_page(driver)
        current_page = get_current_page(driver)
        print(f"advanced_to_page={current_page}", flush=True)
    return current_page


def write_outputs(rows: List[Dict[str, str]], output_path: Path, checkpoint_path: Path, last_page: int) -> None:
    if not rows:
        return
    df = pd.DataFrame(rows)
    if "链接" in df.columns:
        df["_non_empty"] = df.apply(lambda row: sum(str(v).strip() != "" for v in row), axis=1)
        df = df.sort_values(["链接", "_non_empty"], ascending=[True, False])
        df = df.drop_duplicates("链接", keep="first").drop(columns=["_non_empty"])
        df = df.sort_values(["列表页码", "标的物ID"], kind="stable")
    tmp_path = output_path.with_suffix(output_path.suffix + ".tmp.xlsx")
    df.to_excel(tmp_path, index=False)
    last_error = None
    for _ in range(20):
        try:
            tmp_path.replace(output_path)
            last_error = None
            break
        except PermissionError as exc:
            last_error = exc
            time.sleep(1)
    if last_error is not None:
        fallback_path = output_path.with_name(
            f"{output_path.stem}_{datetime.now().strftime('%Y%m%d_%H%M%S')}{output_path.suffix}"
        )
        tmp_path.replace(fallback_path)
        output_path = fallback_path
    checkpoint_path.write_text(
        json.dumps(
            {
                "last_page": last_page,
                "unique_links": int(df["链接"].nunique()) if "链接" in df.columns else len(df),
                "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "output": str(output_path),
            },
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )


def load_existing(output_path: Path) -> List[Dict[str, str]]:
    if not output_path.exists():
        return []
    return pd.read_excel(output_path).fillna("").astype(object).to_dict("records")


def main() -> None:
    parser = argparse.ArgumentParser(description="Fast JD Guangdong index crawler")
    parser.add_argument("--browser-mode", choices=["debug", "normal"], default="debug")
    parser.add_argument("--debug-address", default="127.0.0.1:9222")
    parser.add_argument("--headless", action="store_true")
    parser.add_argument("--province-label", default="广东")
    parser.add_argument("--city-label", default="")
    parser.add_argument("--output", default=r"output\京东法拍房_广东_全量索引_live.xlsx")
    parser.add_argument("--checkpoint", default=r"output\京东法拍房_广东_全量索引_live.checkpoint.json")
    parser.add_argument("--start-page", type=int, default=1)
    parser.add_argument("--max-pages", type=int, default=999)
    parser.add_argument("--save-every", type=int, default=1)
    parser.add_argument(
        "--quit-browser",
        action="store_true",
        help="Also quit the attached debug browser. By default only the working tab is closed.",
    )
    args = parser.parse_args()

    output_path = Path(args.output).resolve()
    checkpoint_path = Path(args.checkpoint).resolve()
    output_path.parent.mkdir(parents=True, exist_ok=True)

    rows = load_existing(output_path)
    driver = (
        connect_debug_browser(args.debug_address)
        if args.browser_mode == "debug"
        else create_normal_browser(headless=args.headless)
    )
    original_handle = None
    try:
        if args.browser_mode == "debug":
            original_handle = open_work_tab(driver, BASE_URL)
        else:
            driver.get(BASE_URL)
        wait_for_list_page(driver)
        select_location(driver, args.province_label, args.city_label)
        wait_for_list_page(driver)
        if args.start_page > 1:
            reached_page = advance_to_page(driver, args.start_page)
            if reached_page < args.start_page:
                raise RuntimeError(f"无法跳转到起始页 {args.start_page}，当前页 {reached_page}")

        no_new_pages = 0
        for _ in range(args.max_pages):
            page_no = get_current_page(driver)
            page_rows = extract_page_items(driver, args.province_label, args.city_label, page_no)
            before = len({str(row.get("链接", "")).strip() for row in rows if str(row.get("链接", "")).strip()})
            rows.extend(page_rows)
            after = len({str(row.get("链接", "")).strip() for row in rows if str(row.get("链接", "")).strip()})
            print(
                f"[{datetime.now().strftime('%H:%M:%S')}] page={page_no} "
                f"items={len(page_rows)} unique={after} new={after-before}",
                flush=True,
            )

            if page_no % args.save_every == 0:
                write_outputs(rows, output_path, checkpoint_path, page_no)

            if after == before:
                no_new_pages += 1
            else:
                no_new_pages = 0

            if no_new_pages >= 5:
                print("No new links for 5 consecutive pages; stopping.", flush=True)
                break

            if not click_next_page(driver, page_no):
                print("No next page or next page did not change; stopping.", flush=True)
                break
            time.sleep(1.5)

        write_outputs(rows, output_path, checkpoint_path, get_current_page(driver))
    finally:
        try:
            driver.close()
        except Exception:
            pass
        try:
            if original_handle and original_handle in driver.window_handles:
                driver.switch_to.window(original_handle)
        except Exception:
            pass
        if args.browser_mode == "normal" or args.quit_browser:
            try:
                driver.quit()
            except Exception:
                pass


if __name__ == "__main__":
    main()
