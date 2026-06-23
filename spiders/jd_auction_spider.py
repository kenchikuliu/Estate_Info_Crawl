# -*- coding: utf-8 -*-
"""
京东法拍房爬虫
"""
import re
import time
import random
import sys
from time import sleep
import subprocess
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from bs4 import BeautifulSoup
import pandas as pd
from typing import Dict, Any, Optional
import undetected_chromedriver as uc
from spiders.base_spider import BaseSpider
from utils.data_storage import DataStorage
from utils.jd_detail_parser import (
    clean_text,
    dump_json,
    extract_attachment_text,
    extract_labeled_fields,
    extract_pdf_fields,
    extract_rights_status_text,
    extract_region_fields,
    filter_image_urls,
    list_attachment_files,
    list_pdf_files,
    normalize_field_value,
    parse_survey_table_rows,
    parse_intro_sections,
    postprocess_structured_fields,
)
from config import Config
import os
from datetime import datetime
import winreg
from pathlib import Path
import json

class JDAuctionSpider(BaseSpider):
    """京东法拍房爬虫"""

    EXTRA_PARSE_COLUMNS = [
        "拍品名称",
        "拍品所有人",
        "拍品现状",
        "租赁情况",
        "钥匙/占用情况",
        "户籍/工商注册",
        "权利限制状况及抵押状况",
        "成交后提供的文件",
        "拍品介绍",
        "房屋权属状况",
        "土地权属状况",
        "附件文本",
        "回填时间",
    ]

    PROVINCE_LABELS = {
        "gd": "广东",
        "zj": "浙江",
        "bj": "北京",
        "sh": "上海",
        "sc": "四川",
        "hb": "湖北",
    }

    CITY_LABELS = {
        "sz": "深圳市",
        "hz": "杭州市",
        "cd": "成都市",
        "wh": "武汉市",
    }

    OUTPUT_COLUMNS = [
        "标的物ID", "链接", "当前价", "起拍价", "保证金", "评估价", "加价幅度", "竞价周期", "延时周期",
        "标题", "完整地址", "处置法院", "城市",
        "标的物详情描述",
        "标的物名称", "权利来源", "权证情况", "被执行人", "钥匙", "户籍注册", "欠费情况", "提供文件",
        "建筑面积", "房屋类型", "房屋用途", "总层数", "核准日期", "所有权来源", "土地用途", "土地性质", "使用期限", "所在层",
        "坐标", "省", "市", "区县", "格式化地址",
        "交通-地铁站", "交通-公交站", "教育-幼儿园", "教育-小学", "教育-中学", "购物-商场", "购物-超市",
        "医疗-综合医院", "医疗-诊所", "公园-公园",
        "大家都在问_QA", "图片链接", "图片数量", "竞买公告", "竞买须知", "拍卖公告",
        "详情页截图路径", "正文截图路径", "附件索引", "资源目录", "解析状态",
    ]

    def __init__(self, start_page: int = 1, max_pages: int = None, province: str = None, city: str = None, cutoff_time: str = None, resume_from_archive: bool = False, max_items: int = None, output_suffix: str = "", crawl_mode: str = "fast"):
        """
        初始化京东法拍房爬虫

        Args:
            start_page: 开始页码
            max_pages: 最大爬取页数
            province: 要爬取的省份
            city: 要爬取的城市
            cutoff_time: 截止时间，格式为"YYYY年MM月DD日 HH:MM:SS"，当拍卖结束时间早于此时间时停止爬取
            resume_from_archive: 是否从存档恢复爬取
        """
        super().__init__("京东法拍房")
        self.start_page = start_page or 1
        self.max_pages = max_pages or Config.JD_AUCTION_CONFIG["max_pages"]
        self.config = Config.JD_AUCTION_CONFIG
        self.cutoff_time = cutoff_time
        self.should_stop = False  # 控制爬取停止的标志
        self.resume_from_archive = resume_from_archive
        self.max_items = max_items
        self.output_suffix = output_suffix.strip()
        self.crawl_mode = crawl_mode
        self.current_asset_folder = ""
        self.location_filter_applied = False
        self.processed_items = 0
        self.last_crawled_asset_name = None  # 存档中最后一条记录的资产名称
        self.should_start_crawling = True  # 是否开始正式爬取的标志

        # 设置省份和城市
        self.province = province or Config.JD_AUCTION_CONFIG["default_province"]
        self.city = city
        if self.city is None:
            default_city = Config.JD_AUCTION_CONFIG.get("default_city")
            available_cities = Config.JD_AUCTION_CONFIG["province_city_mapping"].get(self.province, [])
            if default_city in available_cities:
                self.city = default_city

        # 验证省份和城市是否有效
        self._validate_location()

        # 如果设置了截止时间，验证格式并记录
        if self.cutoff_time:
            self._validate_cutoff_time()
            self.logger.info(f"设置截止时间: {self.cutoff_time}")

        # 如果启用存档恢复，获取最后一条记录的资产名称
        if self.resume_from_archive:
            self.last_crawled_asset_name = self._get_last_asset_name_from_archive()
            if self.last_crawled_asset_name:
                self.should_start_crawling = False  # 需要先找到对应记录
                self.logger.info(f"启用存档恢复模式，最后爬取的资产名称: {self.last_crawled_asset_name}")
            else:
                self.logger.warning("未找到存档文件或存档文件为空，将从头开始爬取")

    def _validate_cutoff_time(self) -> None:
        """
        验证截止时间格式是否正确
        """
        try:
            self._parse_time_string(self.cutoff_time)
            self.logger.info(f"截止时间格式验证通过: {self.cutoff_time}")
        except ValueError as e:
            raise ValueError(f"截止时间格式错误: {e}。正确格式示例: '2024年01月01日 12:00:00'")

    def _parse_time_string(self, time_str: str) -> datetime:
        """
        解析时间字符串为datetime对象

        Args:
            time_str: 时间字符串，格式为"YYYY年MM月DD日 HH:MM:SS"

        Returns:
            datetime: 解析后的datetime对象
        """
        try:
            return datetime.strptime(time_str, "%Y-%m-%d %H:%M:%S")
        except ValueError:
            raise ValueError(f"时间格式错误: {time_str}，正确格式: 'YYYY年MM月DD日 HH:MM:SS'")

    def _is_end_time_before_cutoff(self, end_time: str) -> bool:
        """
        检查结束时间是否早于截止时间

        Args:
            end_time: 拍卖结束时间字符串

        Returns:
            bool: 如果结束时间早于截止时间返回True，否则返回False
        """
        if not self.cutoff_time or not end_time:
            return False

        try:
            end_datetime = self._parse_time_string(end_time)
            cutoff_datetime = self._parse_time_string(self.cutoff_time)

            is_before = end_datetime < cutoff_datetime
            if is_before:
                self.logger.info(f"拍卖结束时间 {end_time} 早于截止时间 {self.cutoff_time}，准备停止爬取")

            return is_before
        except ValueError as e:
            self.logger.warning(f"时间比较失败: {e}")
            return False

    def _get_last_asset_name_from_archive(self) -> Optional[str]:
        """
        从存档文件中获取最后一条记录的资产名称

        Returns:
            Optional[str]: 最后一条记录的资产名称，如果没有找到则返回None
        """
        try:
            output_dir = Config.OUTPUT_DIR
            if not os.path.isdir(output_dir):
                self.logger.info("输出目录不存在，未找到任何xlsx存档文件")
                return None

            archive_files = [
                os.path.join(output_dir, file_name)
                for file_name in os.listdir(output_dir)
                if file_name.startswith("京东法拍房_数据_错误保存") and file_name.endswith(".xlsx")
            ]

            if not archive_files:
                self.logger.info("未找到任何xlsx存档文件")
                return None

            archive_files.sort(key=os.path.getmtime, reverse=True)
            latest_archive = archive_files[0]
            self.logger.info(f"读取最新存档文件: {latest_archive}")

            # 读取Excel文件
            df = pd.read_excel(latest_archive)

            if df.empty:
                self.logger.info("存档文件为空")
                return None

            # 检查是否包含"资产名称"列
            if "资产名称" not in df.columns:
                self.logger.warning("存档文件中未找到'资产名称'列")
                return None

            # 获取最后一条记录的资产名称
            last_asset_name = df["资产名称"].iloc[-1]

            if pd.isna(last_asset_name) or str(last_asset_name).strip() == "":
                self.logger.warning("最后一条记录的资产名称为空")
                return None

            self.logger.info(f"从存档文件中读取到最后一条记录的资产名称: {last_asset_name}")
            if "】" in str(last_asset_name):
                return str(last_asset_name.split("】", 1)[1]).strip()
            return str(last_asset_name).strip()

        except Exception as e:
            self.logger.error(f"读取存档文件失败: {e}")
            return None

    def _validate_location(self) -> None:
        """
        验证省份和城市是否有效
        """
        available_provinces = list(self.config["province_xpath_mapping"].keys())
        if self.province not in available_provinces:
            raise ValueError(f"不支持的省份: {self.province}。支持的省份: {', '.join(available_provinces)}")

        available_cities = self.config["province_city_mapping"].get(self.province, [])
        if self.city and self.city not in available_cities:
            raise ValueError(f"省份 {self.province} 不支持城市 {self.city}。支持的城市: {', '.join(available_cities)}")

    def setup_driver(self) -> None:
        """
        设置浏览器驱动（使用undetected-chromedriver）
        """
        try:
            # 配置 undetected-chromedriver 选项
            options = uc.ChromeOptions()
            options.add_argument("--disable-popup-blocking")
            options.add_argument("--disable-notifications")
            options.add_argument("--disable-infobars")
            options.add_argument("--disable-extensions")
            options.add_argument("--no-sandbox")
            options.add_argument("--disable-dev-shm-usage")
            options.add_argument("--disable-blink-features=AutomationControlled")

            version_main = self._detect_local_chrome_major_version()
            self._cleanup_uc_driver_artifact()

            # 创建 undetected-chromedriver 实例
            self.driver = uc.Chrome(options=options, version_main=version_main)
            self.driver.implicitly_wait(Config.BROWSER_CONFIG["implicit_wait"])
            self.driver.set_page_load_timeout(Config.BROWSER_CONFIG["page_load_timeout"])
            self.driver.set_script_timeout(Config.BROWSER_CONFIG["script_timeout"])
            self.logger.info(f"成功创建 undetected-chromedriver 浏览器实例，Chrome主版本: {version_main}")

        except Exception as e:
            self.logger.error(f"创建浏览器驱动失败: {e}")
            raise Exception(f"无法创建浏览器驱动: {e}")

    def _cleanup_uc_driver_artifact(self) -> None:
        """
        清理 undetected-chromedriver 在 Windows 上遗留的驱动副本，避免 WinError 183。
        """
        candidate = Path.home() / "appdata" / "roaming" / "undetected_chromedriver" / "undetected_chromedriver.exe"
        if candidate.exists():
            try:
                candidate.unlink()
                self.logger.info(f"已清理残留驱动文件: {candidate}")
            except PermissionError:
                self.logger.warning(f"残留驱动文件正在被占用，跳过清理: {candidate}")
            except OSError as e:
                self.logger.warning(f"清理残留驱动文件失败: {e}")

    def _find_listing_elements(self):
        """
        基于真实拍卖链接定位列表项，避免依赖易变的页面层级 XPath。
        """
        anchor_xpath = "//a[contains(@href, 'paimai.jd.com/')]"
        anchors = self.driver.find_elements(By.XPATH, anchor_xpath)
        listing_elements = []
        seen = set()

        for anchor in anchors:
            href = anchor.get_attribute("href") or anchor.get_property("href") or ""
            if not href or href in seen:
                continue
            if not re.search(r"https://paimai\.jd\.com/\d+(?:$|[/?#])", href):
                continue
            seen.add(href)
            try:
                container = anchor.find_element(By.XPATH, "./ancestor::li[1]")
            except Exception:
                container = anchor
            listing_elements.append(container)

        return listing_elements

    def _detect_local_chrome_major_version(self) -> Optional[int]:
        """
        自动探测本机 Chrome 主版本，避免驱动和浏览器版本不匹配。
        """
        version_sources = [
            (winreg.HKEY_CURRENT_USER, r"Software\Google\Chrome\BLBeacon"),
            (winreg.HKEY_LOCAL_MACHINE, r"Software\Google\Chrome\BLBeacon"),
            (winreg.HKEY_LOCAL_MACHINE, r"Software\WOW6432Node\Google\Chrome\BLBeacon"),
        ]

        for root, path in version_sources:
            try:
                key = winreg.OpenKey(root, path)
                version, _ = winreg.QueryValueEx(key, "version")
                if version:
                    return int(str(version).split(".")[0])
            except OSError:
                continue

        chrome_paths = [
            r"C:\Program Files\Google\Chrome\Application\chrome.exe",
            r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
            os.path.join(os.environ.get("LOCALAPPDATA", ""), r"Google\Chrome\Application\chrome.exe"),
        ]

        for chrome_path in chrome_paths:
            if not chrome_path or not os.path.exists(chrome_path):
                continue
            try:
                completed = subprocess.run(
                    [chrome_path, "--version"],
                    capture_output=True,
                    text=True,
                    timeout=10,
                    check=True,
                )
                version_text = completed.stdout.strip() or completed.stderr.strip()
                match = re.search(r"(\d+)\.", version_text)
                if match:
                    return int(match.group(1))
            except Exception:
                continue

        self.logger.warning("未能探测到本机 Chrome 版本，将使用 undetected-chromedriver 默认版本")
        return None

    def get_output_filename(self) -> str:
        province_label = self.PROVINCE_LABELS.get(self.province, self.province or "all")
        city_label = self.CITY_LABELS.get(self.city, self.city or "")
        suffix = f"_{city_label}" if city_label else ""
        output_suffix = f"_{self.output_suffix}" if self.output_suffix else ""
        return f"京东法拍房_{province_label}{suffix}{output_suffix}_数据.xlsx"

    def get_error_output_filename(self, timestamp: str) -> str:
        province_label = self.PROVINCE_LABELS.get(self.province, self.province or "all")
        city_label = self.CITY_LABELS.get(self.city, self.city or "")
        suffix = f"_{city_label}" if city_label else ""
        output_suffix = f"_{self.output_suffix}" if self.output_suffix else ""
        return f"京东法拍房_{province_label}{suffix}{output_suffix}_数据_错误保存_{timestamp}.xlsx"

    def _make_safe_name(self, value: str, fallback: str = "asset") -> str:
        cleaned = clean_text(value)
        cleaned = re.sub(r'[<>:"/\\|?*]+', "_", cleaned).strip(" .")
        return (cleaned[:80] if cleaned else fallback) or fallback

    def _build_asset_folder(self, detail_info: Dict[str, Any]) -> str:
        asset_id = detail_info.get("标的物ID", "") or "unknown"
        title = self._make_safe_name(detail_info.get("标题", "") or detail_info.get("资产名称", ""), fallback=asset_id)
        return self.data_storage.create_folder(f"京东法拍/{asset_id}_{title}")

    def _resolve_asset_folder(self, asset_name: str) -> str:
        if self.current_asset_folder:
            os.makedirs(self.current_asset_folder, exist_ok=True)
            return self.current_asset_folder
        return self.data_storage.create_folder(f"京东法拍/{self._make_safe_name(asset_name)}")

    def _capture_detail_snapshots(self, folder_path: str) -> Dict[str, str]:
        snapshot_info = {
            "详情页截图路径": "",
            "正文截图路径": "",
        }
        if not folder_path:
            return snapshot_info

        try:
            full_path = os.path.join(folder_path, "detail_page.png")
            self.driver.save_screenshot(full_path)
            snapshot_info["详情页截图路径"] = full_path
        except Exception:
            pass

        try:
            content_element = self.driver.find_element(By.XPATH, "//*[@id='pmMainFloor']/ul/li[1]")
            content_path = os.path.join(folder_path, "detail_intro.png")
            content_element.screenshot(content_path)
            snapshot_info["正文截图路径"] = content_path
        except Exception:
            pass

        return snapshot_info

    def _collect_attachment_index(self) -> str:
        attachments = []
        try:
            file_list = self.driver.find_elements(By.XPATH, "//*[@id='pmMainFloor']/ul/li[1]/div[1]/div/div/div[1]/ul/li")
            for item in file_list:
                try:
                    link = item.find_element(By.XPATH, ".//*[@id='openAttachmentTag']")
                    attachments.append({
                        "name": clean_text(link.text),
                        "url": clean_text(link.get_property("href") or ""),
                    })
                except Exception:
                    continue
        except Exception:
            pass
        return json.dumps(attachments, ensure_ascii=False)

    def run(self) -> None:
        """
        运行爬虫逻辑
        """
        try:
            # 设置浏览器驱动
            self.setup_driver()

            # 自动打开京东法拍页面
            self.wait_for_manual_page_open()

            # 等待手动登录
            self.wait_for_manual_login()

            # 选择地区
            self.select_location()

            # 手动进行更多筛选
            self._wait_for_enter("如需进行更多筛选，请手动操作，按回车键继续...")

            # 开始爬取
            self.crawl_auction_data()

        except Exception as e:
            self.logger.error(f"爬取过程中出错: {e}")
            raise
        finally:
            # 关闭浏览器
            if hasattr(self, 'driver') and self.driver:
                try:
                    self.driver.quit()
                    self.logger.info("浏览器已关闭")
                except:
                    pass

    def wait_for_manual_page_open(self) -> None:
        """
        自动打开京东法拍页面
        """
        self.logger.info("=" * 50)
        self.logger.info("正在自动打开京东法拍页面...")
        self.logger.info(f"目标URL: {self.config['base_url']}")
        self.logger.info("=" * 50)

        last_error = None
        for attempt in range(1, 4):
            try:
                if attempt > 1:
                    self.logger.warning(f"第 {attempt - 1} 次打开页面失败，准备重试第 {attempt} 次")
                    try:
                        self.driver.get("about:blank")
                        sleep(random.uniform(1, 2))
                    except Exception:
                        pass

                # 直接导航到京东法拍页面
                self.driver.get(self.config['base_url'])

                # 等待页面加载完成
                self.logger.info("等待页面加载完成...")
                WebDriverWait(self.driver, 30).until(
                    EC.presence_of_element_located((By.CLASS_NAME, "province"))
                )

                # 增加额外等待，确保页面完全加载
                sleep(random.uniform(3, 5))

                self.logger.info("京东法拍页面加载完成，开始执行后续逻辑...")
                return

            except Exception as e:
                last_error = e
                self.logger.error(f"打开京东法拍页面失败(第 {attempt} 次): {e}")

        raise Exception(f"无法打开京东法拍页面: {last_error}")


    def wait_for_manual_login(self) -> None:
        """
        等待手动登录网站
        """
        self.logger.info("=" * 50)

        # 首先检查是否已经登录
        if self.check_login_status():
            self.logger.info("检测到已登录状态，无需重新登录")
            return

        self.logger.info("请手动登录京东法拍网站")
        self.logger.info("登录完成后，请按回车键继续...")
        self.logger.info("=" * 50)

        try:
            # 等待用户输入
            self._wait_for_enter("按回车键继续...")

            # 验证是否已登录
            if self.check_login_status():
                self.logger.info("登录验证完成，继续执行后续逻辑...")
            else:
                self.logger.warning("登录状态验证失败，将继续执行。若后续详情页受限，请手动登录后重试。")

        except KeyboardInterrupt:
            self.logger.info("用户中断操作")
            raise Exception("用户中断登录过程")
        except Exception as e:
            self.logger.error(f"等待登录过程中出错: {e}")
            raise

    def check_login_status(self) -> bool:
        """
        检查登录状态

        Returns:
            bool: 是否已登录
        """
        try:
            # 等待一段时间让页面稳定
            sleep(2)

            # 检查URL是否包含登录相关信息
            current_url = self.driver.current_url
            if "login" in current_url.lower() or "signin" in current_url.lower():
                return False

            # 检查页面中是否存在登录相关元素
            try:
                # 查找登录按钮或登录链接
                login_elements = self.driver.find_elements(By.XPATH,
                    "//*[contains(text(), '登录') or contains(text(), '登录') or contains(@class, 'login')]")

                # 查找用户相关元素（如用户头像、用户名等）
                user_elements = self.driver.find_elements(By.XPATH,
                    "//*[contains(@class, 'user') or contains(@class, 'avatar') or contains(@class, 'nickname')]")

                # 如果找到用户相关元素且没有登录按钮，认为已登录
                if user_elements and not login_elements:
                    self.logger.info("检测到用户相关元素，判断为已登录状态")
                    return True

                # 如果找到登录按钮，认为未登录
                if login_elements:
                    self.logger.info("检测到登录按钮，判断为未登录状态")
                    return False

                # 其他情况，尝试访问需要登录的页面来验证
                return self.verify_login_by_access()

            except Exception as e:
                self.logger.debug(f"检查登录状态时出错: {e}")
                return self.verify_login_by_access()

        except Exception as e:
            self.logger.debug(f"检查登录状态时出错: {e}")
            return False


    def verify_login_by_access(self) -> bool:
        """
        通过访问需要登录的页面来验证登录状态

        Returns:
            bool: 是否已登录
        """
        try:
            # 记录当前URL
            original_url = self.driver.current_url

            # 尝试访问一个需要登录的页面
            test_url = "https://pmsearch.jd.com/user/center"
            self.driver.get(test_url)
            sleep(3)

            # 检查是否被重定向到登录页面
            current_url = self.driver.current_url
            if "login" in current_url.lower() or "signin" in current_url.lower():
                self.logger.info("访问用户中心被重定向到登录页面，判断为未登录")
                # 回到原页面
                self.driver.get(original_url)
                return False

            # 检查页面内容是否包含用户信息
            page_source = self.driver.page_source
            if "用户中心" in page_source or "个人中心" in page_source:
                self.logger.info("成功访问用户中心，判断为已登录")
                # 回到原页面
                self.driver.get(original_url)
                return True

            # 回到原页面
            self.driver.get(original_url)
            return False

        except Exception as e:
            self.logger.debug(f"通过访问验证登录状态时出错: {e}")
            return False

    def handle_verification_popup(self) -> None:
        """
        处理验证弹窗（支持多种弹窗类型）
        """
        max_retries = 3
        popup_selectors = [
            # 基于类名的选择器（处理空格问题）
            "//div[contains(@class, 'alert-popup-button-confirm')]",
            # 基于文本内容的选择器（忽略前后空格）
            "//div[normalize-space(text())='我已知晓并同意']",
            # 组合选择器
            "//div[@class=' alert-popup-button-confirm']",
            # 更宽泛的选择器
            "//div[contains(@class, 'alert-popup-buttons')]//div[contains(text(), '我已知晓')]",
            # 基于父容器的选择器
            "//div[contains(@class, 'alert-popup-overlay')]//div[contains(text(), '我已知晓并同意')]"
        ]

        for attempt in range(max_retries):
            try:
                self.logger.info(f"检查验证弹窗 (第{attempt + 1}次)")

                # 首先检查弹窗是否存在
                popup_overlay = None
                try:
                    popup_overlay = WebDriverWait(self.driver, 5).until(
                        EC.presence_of_element_located((By.CLASS_NAME, "alert-popup-overlay"))
                    )
                    self.logger.info("检测到验证弹窗")
                except:
                    self.logger.info("未检测到验证弹窗")
                    return

                # 等待弹窗完全加载
                sleep(random.uniform(1, 2))

                # 尝试不同的选择器找到确认按钮
                confirm_button = None
                for i, selector in enumerate(popup_selectors):
                    try:
                        self.logger.debug(f"尝试选择器 {i+1}: {selector}")
                        confirm_button = WebDriverWait(self.driver, 3).until(
                            EC.element_to_be_clickable((By.XPATH, selector))
                        )
                        self.logger.info(f"使用选择器 {i+1} 成功找到确认按钮")
                        break
                    except:
                        continue

                if not confirm_button:
                    self.logger.warning("未找到可点击的确认按钮，尝试通过关闭按钮关闭弹窗")
                    try:
                        close_button = self.driver.find_element(By.CLASS_NAME, "alert-popup-close")
                        close_button.click()
                        self.logger.info("通过关闭按钮关闭了验证弹窗")
                        return
                    except:
                        self.logger.warning("也未找到关闭按钮")
                        if attempt < max_retries - 1:
                            continue
                        else:
                            self.logger.error("无法处理验证弹窗，继续执行...")
                            return

                # 模拟人类行为：先滚动到按钮位置
                self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", confirm_button)
                sleep(random.uniform(0.5, 1))

                # 模拟鼠标悬停
                ActionChains(self.driver).move_to_element(confirm_button).perform()
                sleep(random.uniform(0.3, 0.8))

                # 点击确认按钮
                confirm_button.click()
                self.logger.info("成功点击验证弹窗的确认按钮")

                # 等待弹窗消失
                try:
                    WebDriverWait(self.driver, 5).until(
                        EC.invisibility_of_element_located((By.CLASS_NAME, "alert-popup-overlay"))
                    )
                    self.logger.info("验证弹窗已消失")
                except:
                    self.logger.warning("验证弹窗可能未完全消失")

                # 额外等待，确保页面稳定
                sleep(random.uniform(1, 2))
                return

            except Exception as e:
                self.logger.warning(f"处理验证弹窗失败 (尝试 {attempt + 1}/{max_retries}): {e}")
                if attempt < max_retries - 1:
                    sleep(random.uniform(2, 4))

        self.logger.error(f"处理验证弹窗失败，已重试 {max_retries} 次，继续执行...")

        # 尝试最后的兜底方案：按ESC键关闭弹窗
        try:
            from selenium.webdriver.common.keys import Keys
            self.driver.find_element(By.TAG_NAME, "body").send_keys(Keys.ESCAPE)
            self.logger.info("尝试使用ESC键关闭弹窗")
            sleep(1)
        except:
            pass

    def verify_login_status(self) -> None:
        """
        验证登录状态（可选）
        """
        try:
            # 等待一段时间让页面稳定
            sleep(2)

            try:
                # 如果找到用户相关元素，说明可能已登录
                user_elements = self.driver.find_elements(By.XPATH, "//*[contains(@class, 'user') or contains(@class, 'login') or contains(@class, 'avatar')]")
                if user_elements:
                    self.logger.info("检测到可能的登录状态")
                else:
                    self.logger.warning("未检测到明显的登录状态，请确认是否已正确登录")
            except:
                self.logger.info("无法检测登录状态，请手动确认")

        except Exception as e:
            self.logger.debug(f"验证登录状态时出错: {e}")
            # 不抛出异常，因为这只是可选的验证

    def _wait_for_enter(self, prompt: str) -> None:
        """
        交互环境下等待用户确认；非交互环境自动继续。
        """
        if sys.stdin is None or not sys.stdin.isatty():
            self.logger.info("检测到非交互环境，自动继续执行")
            return
        try:
            input(prompt)
        except EOFError:
            self.logger.info("读取标准输入失败，自动继续执行")

    def select_location(self) -> None:
        """
        选择地区（增加反反爬虫措施）
        """
        self.location_filter_applied = False
        try:
            # 等待页面完全加载
            self.wait_for_page_load()

            # 选择省份
            self.select_province_with_retry()

            # 如果有城市选择，则选择城市
            if self.city:
                self.select_city_with_retry()
                self.logger.info(f"地区选择完成: {self.province}-{self.city}")
            else:
                self.logger.info(f"地区选择完成: {self.province}")
            self.location_filter_applied = True

        except Exception as e:
            self.logger.warning(f"选择地区失败，降级为列表后置过滤模式: {e}")

    def wait_for_page_load(self) -> None:
        """
        等待页面完全加载
        """
        self.logger.info("等待页面完全加载...")
        # 随机等待时间，模拟人类行为
        sleep(random.uniform(2, 4))

        # 等待关键元素加载
        try:
            WebDriverWait(self.driver, 15).until(
                EC.presence_of_element_located((By.CLASS_NAME, "province"))
            )
            self.logger.info("页面加载完成")
        except Exception as e:
            self.logger.warning(f"等待页面加载超时: {e}")

    def select_province_with_retry(self) -> None:
        """
        选择省份（带重试机制）
        """
        max_retries = 3
        province_name = self.PROVINCE_LABELS[self.province]
        province_xpath = (
            f"//dl[contains(@class,'province')]//dd"
            f"//a[normalize-space()='{province_name}']"
        )
        for attempt in range(max_retries):
            try:
                self.logger.info(f"尝试选择省份: {self.province} (第{attempt + 1}次)")

                if attempt > 0:
                    self.driver.refresh()
                    self.wait_for_page_load()

                province_element = WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located((By.CLASS_NAME, "province"))
                )
                self._move_and_click(province_element)

                self.wait_for_dropdown_appear(province_xpath)

                province_option = WebDriverWait(self.driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, province_xpath))
                )
                self._move_and_click(province_option)

                if self.verify_province_selection():
                    self.logger.info(f"省份选择成功: {self.province}")
                    return
                else:
                    self.logger.warning(f"省份选择未生效，准备重试 (尝试 {attempt + 1}/{max_retries})")

            except Exception as e:
                self.logger.warning(f"选择省份失败 (尝试 {attempt + 1}/{max_retries}): {e}")

            # 重试前等待
            if attempt < max_retries - 1:
                sleep(random.uniform(2, 4))

        raise Exception(f"选择省份失败，已重试 {max_retries} 次")

    def select_city_with_retry(self) -> None:
        """
        选择城市（带重试机制）
        """
        max_retries = 3
        city_name = self.CITY_LABELS[self.city]
        city_xpath = (
            f"//dl[contains(@class,'city')]//dd"
            f"//a[normalize-space()='{city_name}']"
        )
        for attempt in range(max_retries):
            try:
                self.logger.info(f"尝试选择城市: {self.city} (第{attempt + 1}次)")

                if attempt > 0:
                    self.driver.refresh()
                    self.wait_for_page_load()

                city_element = WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located((By.CLASS_NAME, "city"))
                )
                self._move_and_click(city_element)

                self.wait_for_dropdown_appear(city_xpath)

                city_option = WebDriverWait(self.driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, city_xpath))
                )
                self._move_and_click(city_option)

                if self.verify_city_selection():
                    self.logger.info(f"城市选择成功: {self.city}")
                    return
                else:
                    self.logger.warning(f"城市选择未生效，准备重试 (尝试 {attempt + 1}/{max_retries})")

            except Exception as e:
                self.logger.warning(f"选择城市失败 (尝试 {attempt + 1}/{max_retries}): {e}")

            # 重试前等待
            if attempt < max_retries - 1:
                sleep(random.uniform(2, 4))

        raise Exception(f"选择城市失败，已重试 {max_retries} 次")

    def _move_and_click(self, element) -> None:
        try:
            self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", element)
        except Exception:
            pass
        sleep(random.uniform(0.5, 1.2))
        try:
            ActionChains(self.driver).move_to_element(element).perform()
        except Exception:
            pass
        sleep(random.uniform(0.3, 0.8))
        try:
            element.click()
        except Exception:
            self.driver.execute_script("arguments[0].click();", element)

    def wait_for_dropdown_appear(self, option_xpath: str = None) -> None:
        """
        等待下拉菜单出现
        """
        if option_xpath:
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.XPATH, option_xpath))
            )
        else:
            sleep(random.uniform(1, 2))

    def verify_province_selection(self) -> bool:
        """
        验证省份选择是否生效
        """
        try:
            # 等待页面更新
            sleep(random.uniform(2, 4))

            # 检查URL是否发生变化
            current_url = self.driver.current_url
            if self.province.lower() in current_url.lower():
                return True

            # 检查页面元素变化
            try:
                province_element = self.driver.find_element(By.XPATH, "//dl[contains(@class,'province')]/dt/em")
                if self.PROVINCE_LABELS[self.province] in province_element.text:
                    return True
            except:
                pass

            # 检查是否有新的列表项加载
            try:
                list_elements = self._find_listing_elements()
                if len(list_elements) > 0:
                    return True
            except:
                pass

            return False

        except Exception as e:
            self.logger.debug(f"验证省份选择时出错: {e}")
            return False

    def verify_city_selection(self) -> bool:
        """
        验证城市选择是否生效
        """
        try:
            # 等待页面更新
            sleep(random.uniform(2, 4))

            # 检查URL是否发生变化
            current_url = self.driver.current_url
            if self.city.lower() in current_url.lower():
                return True

            # 检查页面元素变化
            try:
                city_element = self.driver.find_element(By.XPATH, "//dl[contains(@class,'city')]/dt/em")
                if self.CITY_LABELS[self.city] in city_element.text:
                    return True
            except:
                pass

            # 检查是否有新的列表项加载
            try:
                list_elements = self._find_listing_elements()
                if len(list_elements) > 0:
                    return True
            except:
                pass

            return False

        except Exception as e:
            self.logger.debug(f"验证城市选择时出错: {e}")
            return False

    def crawl_auction_data(self) -> None:
        """
        爬取拍卖数据
        """
        if not self.location_filter_applied and self.start_page > 1:
            self.logger.warning(
                f"地区筛选未生效且起始页为第 {self.start_page} 页，"
                "为避免回退到第 1 页白跑，本分片直接退出等待补跑"
            )
            return

        page_no = int(self.driver.find_element(By.CLASS_NAME, "ui-pager-current").text)
        page_no = self.transfer_to_start_page(page_no, self.start_page)

        consecutive_failures = 0  # 连续失败次数
        max_consecutive_failures = 3  # 最大连续失败次数

        # 存档恢复模式的日志记录
        if self.resume_from_archive:
            if self.last_crawled_asset_name:
                self.logger.info(f"存档恢复模式启动，正在寻找资产名称: {self.last_crawled_asset_name}")
            else:
                self.logger.info("存档恢复模式启动，但未找到有效的存档数据，将从头开始爬取")

        while page_no <= self.max_pages and not self.should_stop:
            try:
                self.logger.info(f"正在爬取第 {page_no} 页")

                # 随机等待页面加载，模拟人类行为
                sleep(random.uniform(3, 6))

                # 获取列表项
                list_elements = self._find_listing_elements()
                if not list_elements:
                    self.logger.info("没有找到更多数据，爬取结束")
                    break

                self.logger.info(f"本页找到 {len(list_elements)} 个拍卖项")

                # 处理每个拍卖项
                success_count = 0
                for index, element in enumerate(list_elements):
                    try:
                        # 检查是否需要停止爬取
                        if self.should_stop:
                            self.logger.info("检测到停止信号，结束爬取")
                            break
                        if self.max_items and self.processed_items >= self.max_items:
                            self.logger.info(f"达到最大处理数量限制: {self.max_items}")
                            self.should_stop = True
                            break

                        # 添加随机延时，避免操作过快
                        if index > 0:
                            sleep(random.uniform(1, 3))

                        self.logger.info(f"正在处理第 {index + 1} 个拍卖项")
                        self.process_auction_item(element)
                        success_count += 1

                        # 每处理几个项目后稍作休息
                        if (index + 1) % 5 == 0:
                            self.logger.info(f"已处理 {index + 1} 个项目，休息片刻...")
                            sleep(random.uniform(5, 10))

                    except Exception as e:
                        self.logger.error(f"处理拍卖项 {index + 1} 时出错: {e}")
                        continue

                self.logger.info(f"第 {page_no} 页处理完成，成功处理 {success_count}/{len(list_elements)} 个拍卖项")

                # 如果因为时间截止而停止，退出循环
                if self.should_stop:
                    self.logger.info("达到时间截止条件，停止爬取")
                    break

                # 重置连续失败计数

                # 翻页
                target_page = page_no + 1
                new_page_no = self.transfer_to_start_page(page_no, target_page)

                if new_page_no != target_page:
                    self.logger.warning(f"翻页失败，目标页数: {target_page}，实际页数: {new_page_no}")
                    self.save_data()
                    return

                page_no = new_page_no

            except Exception as e:
                self.logger.error(f"爬取第 {page_no} 页时出错: {e}")
                self.save_data()
                return

        # 爬取结束，保存数据
        self.logger.info("数据爬取完成，正在保存数据...")
        self.save_data()

    def process_auction_item(self, element) -> None:
        """
        处理单个拍卖项

        Args:
            element: 拍卖项元素
        """
        try:
            # 获取拍卖状态
            status_element = element.find_element(By.XPATH, ".//a/div[3]/div[1]")
            item_status = status_element.text

            # 获取基本信息
            link = element.find_element(By.XPATH, ".//a").get_property("href")
            item_name = element.find_element(By.XPATH, ".//a/div[2]/div[1]").text
            image = element.find_element(By.XPATH, ".//a/div[1]/div/img").get_attribute('src')
            current_value = element.find_element(By.XPATH, ".//a/div[2]/div[2]/div[2]/em/b").text
            esti_value = element.find_element(By.XPATH, ".//a/div[2]/div[3]/div[1]/em").text

            # 如果启用了存档恢复模式，检查是否应该开始爬取
            if self.resume_from_archive and not self.should_start_crawling:
                if item_name == self.last_crawled_asset_name:
                    self.logger.info(f"找到存档中的最后一条记录: {item_name}，跳过该记录，从下一条开始爬取")
                    self.should_start_crawling = True
                    return  # 跳过这一条记录
                else:
                    self.logger.info(f"跳过记录: {item_name} (正在寻找: {self.last_crawled_asset_name})")
                    return  # 继续跳过，直到找到目标记录

            # 跳过车位、车库拍卖项
            if any(keyword in item_name for keyword in ['车位', '车库', '地下室']):
                self.logger.info(f"跳过车位、车库、地下室拍卖项: {item_name}")
                return

            # 获取详细信息
            detail_info = self.get_auction_detail(link)
            if not detail_info:
                return
            current_asset_name = detail_info.get('标题', '') or detail_info.get('资产名称', '')

            province_label = self.PROVINCE_LABELS.get(self.province, "")
            if province_label:
                province_text = detail_info.get("完整地址", "") + " " + current_asset_name
                if province_label not in province_text:
                    self.logger.info(f"跳过非目标省份标的: {current_asset_name}")
                    return

            if self.city:
                city_label = self.CITY_LABELS.get(self.city, "")
                location_text = detail_info.get("完整地址", "") + " " + current_asset_name
                if city_label and city_label not in location_text:
                    self.logger.info(f"跳过非目标城市标的: {current_asset_name}")
                    return

            # 检查结束时间是否早于截止时间
            end_time = detail_info.get('结束时间', '')
            if self._is_end_time_before_cutoff(end_time):
                self.logger.info(f"拍卖项 '{current_asset_name}' 的结束时间早于截止时间，设置停止标志")
                self.should_stop = True
                # 仍然保存当前这一条数据，然后停止

            # 构建数据项
            detail_info["当前价"] = current_value
            detail_info["评估价"] = esti_value
            detail_info["图片"] = image
            detail_info["竞价状态"] = item_status
            detail_info["结束时间"] = detail_info.get("结束时间", "")

            data_item = {column: detail_info.get(column, "") for column in self.OUTPUT_COLUMNS}

            self.add_data(data_item)
            self.processed_items += 1
            self.save_data()
            self.logger.info(f"成功处理拍卖项: {current_asset_name}")

            # 如果设置了停止标志，提前返回
            if self.should_stop:
                return

        except Exception as e:
            self.logger.error(f"处理拍卖项时出错: {e}")

    def get_auction_detail(self, url: str) -> Optional[Dict[str, Any]]:
        """
        获取拍卖详情

        Args:
            url: 拍卖详情页URL

        Returns:
            Optional[Dict[str, Any]]: 详情信息
        """
        # 记录主窗口句柄
        main_window = self.driver.current_window_handle
        last_error = None

        for attempt in range(1, 3):
            detail_window = None
            try:
                # 打开新窗口
                self.driver.execute_script(f"window.open('{url}', '_blank');")

                # 切换到新窗口
                all_windows = self.driver.window_handles
                for window in all_windows:
                    if window != main_window:
                        detail_window = window
                if not detail_window:
                    raise RuntimeError("未找到详情页窗口")
                self.driver.switch_to.window(detail_window)

                # 等待页面加载
                WebDriverWait(self.driver, 12).until(EC.presence_of_element_located((By.TAG_NAME, "body")))

                # 处理验证弹窗
                self.handle_verification_popup()

                # 获取详细信息
                detail_info = self.extract_detail_info(url)
                folder_path = self._build_asset_folder(detail_info)
                self.current_asset_folder = folder_path
                detail_info["资源目录"] = folder_path
                detail_info["附件索引"] = self._collect_attachment_index()
                detail_info.update(self._capture_detail_snapshots(folder_path))

                if self.crawl_mode == "fast":
                    detail_info["解析状态"] = "待解析"
                    return detail_info

                # 下载附件和图片
                self.download_attachments(detail_info.get('标题', '') or detail_info.get('资产名称', ''))

                # 获取标的物调查表
                self.extract_property_survey_table(detail_info.get('标题', '') or detail_info.get('资产名称', ''))

                # 获取竞买公告和竞买须知
                notice_info = self.extract_notice_info(detail_info.get('标题', '') or detail_info.get('资产名称', ''))

                # 获取竞价记录
                self.extract_bidding_info(detail_info.get('标题', '') or detail_info.get('资产名称', ''))

                # 获取优先购买权人
                self.extract_priority_purchaser(detail_info.get('标题', '') or detail_info.get('资产名称', ''))

                detail_info.update(notice_info or {})
                detail_info.update(self.enrich_detail_from_assets(folder_path, detail_info))
                detail_info["解析状态"] = "已解析"

                return detail_info

            except Exception as e:
                last_error = e
                self.logger.error(f"获取拍卖详情失败(第 {attempt} 次): {e}")
                if attempt < 2:
                    sleep(random.uniform(1, 2))
            finally:
                self.current_asset_folder = ""
                try:
                    if detail_window and detail_window in self.driver.window_handles:
                        self.driver.switch_to.window(detail_window)
                        self.driver.close()
                except Exception:
                    pass
                try:
                    if main_window in self.driver.window_handles:
                        self.driver.switch_to.window(main_window)
                except Exception:
                    pass

        self.logger.error(f"获取拍卖详情失败: {last_error}")
        return None

    def backfill_from_excel(
        self,
        input_file: str,
        output_file: str = "",
        limit: Optional[int] = None,
        only_pending: bool = True,
        start_index: int = 0,
        offline_only: bool = False,
    ) -> str:
        df = pd.read_excel(input_file).fillna("").astype(object)

        for column in self.OUTPUT_COLUMNS + self.EXTRA_PARSE_COLUMNS:
            if column not in df.columns:
                df[column] = ""

        if only_pending:
            pending_mask = df["解析状态"].astype(str).str.strip().ne("已解析")
            work_df = df[pending_mask]
        else:
            work_df = df

        if start_index:
            work_df = work_df.iloc[start_index:]
        if limit:
            work_df = work_df.head(limit)

        if work_df.empty:
            output_path = output_file or input_file
            self.data_storage.write_excel_atomic(df, output_path)
            return output_path

        if not offline_only:
            self.setup_driver()
        try:
            if not offline_only:
                self.wait_for_manual_page_open()
                self.wait_for_manual_login()

            for row_index in work_df.index:
                row = df.loc[row_index].to_dict()
                link = clean_text(str(row.get("链接", "")))
                if not link:
                    continue
                try:
                    updated_row = self.refresh_row_from_assets(row) if offline_only else self.backfill_detail_row(row)
                    for key, value in updated_row.items():
                        df.at[row_index, key] = value
                    self.logger.info(f"回填完成: {link}")
                except Exception as e:
                    self.logger.error(f"回填失败 {link}: {e}")
                    df.at[row_index, "解析状态"] = f"回填失败:{clean_text(str(e))[:80]}"

                output_path = output_file or input_file
                self.data_storage.write_excel_atomic(df, output_path)

        finally:
            if not offline_only and hasattr(self, "driver") and self.driver:
                try:
                    self.driver.quit()
                except Exception:
                    pass

        return output_file or input_file

    def backfill_detail_row(self, row: Dict[str, Any]) -> Dict[str, Any]:
        url = clean_text(str(row.get("链接", "")))
        if not url:
            raise ValueError("缺少链接")

        main_window = self.driver.current_window_handle
        detail_window = None
        last_error = None

        for attempt in range(1, 3):
            try:
                self.driver.execute_script(f"window.open('{url}', '_blank');")
                for window in self.driver.window_handles:
                    if window != main_window:
                        detail_window = window
                if not detail_window:
                    raise RuntimeError("未找到详情页窗口")

                self.driver.switch_to.window(detail_window)
                WebDriverWait(self.driver, 12).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
                self.handle_verification_popup()

                detail_info = self.extract_detail_info(url)
                merged = dict(row)
                merged.update({k: v for k, v in detail_info.items() if clean_text(str(v))})

                folder_path = clean_text(str(row.get("资源目录", ""))) or self._build_asset_folder(merged)
                self.current_asset_folder = folder_path
                merged["资源目录"] = folder_path
                os.makedirs(folder_path, exist_ok=True)

                attachment_index = self._collect_attachment_index()
                if attachment_index and attachment_index != "[]":
                    merged["附件索引"] = attachment_index
                merged.update(self._capture_detail_snapshots(folder_path))

                asset_name = merged.get("标题", "") or merged.get("资产名称", "")
                self.download_attachments(asset_name)
                self.extract_property_survey_table(asset_name)
                notice_info = self.extract_notice_info(asset_name)
                merged.update(notice_info or {})
                merged.update(self.enrich_detail_from_assets(folder_path, merged))
                merged["回填时间"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                merged["解析状态"] = "已解析"

                result = {}
                for key in list(dict.fromkeys(self.OUTPUT_COLUMNS + self.EXTRA_PARSE_COLUMNS + list(row.keys()) + list(merged.keys()))):
                    result[key] = merged.get(key, row.get(key, ""))
                return result

            except Exception as e:
                last_error = e
                self.logger.error(f"回填详情失败(第 {attempt} 次): {e}")
                if attempt < 2:
                    sleep(random.uniform(1, 2))
            finally:
                self.current_asset_folder = ""
                try:
                    if detail_window and detail_window in self.driver.window_handles:
                        self.driver.switch_to.window(detail_window)
                        self.driver.close()
                except Exception:
                    pass
                try:
                    if main_window in self.driver.window_handles:
                        self.driver.switch_to.window(main_window)
                except Exception:
                    pass
                detail_window = None

        raise RuntimeError(last_error or "回填详情失败")

    def refresh_row_from_assets(self, row: Dict[str, Any]) -> Dict[str, Any]:
        merged = dict(row)
        folder_path = clean_text(str(merged.get("资源目录", "")))
        if not folder_path or not os.path.exists(folder_path):
            raise ValueError("资源目录不存在")

        if not clean_text(str(merged.get("竞买公告", ""))):
            notice_file = os.path.join(folder_path, "竞买公告和竞买须知.xlsx")
            if os.path.exists(notice_file):
                try:
                    notice_df = pd.read_excel(notice_file).fillna("")
                    if not notice_df.empty:
                        for field in ["竞买公告", "竞买须知", "拍卖公告"]:
                            if field in notice_df.columns and not clean_text(str(merged.get(field, ""))):
                                merged[field] = clean_text(str(notice_df.iloc[0].get(field, "")))
                except Exception:
                    pass

        detail_text = clean_text(str(merged.get("标的物详情描述", ""))) or clean_text(str(merged.get("标的物介绍原文", "")))
        merged["标的物介绍原文"] = clean_text(str(merged.get("标的物介绍原文", ""))) or detail_text

        merged.update(self.enrich_detail_from_assets(folder_path, merged))
        merged["回填时间"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        if clean_text(str(merged.get("解析状态", ""))) != "已解析":
            merged["解析状态"] = "已解析"
        return merged

    def extract_detail_info(self, url: str) -> Dict[str, Any]:
        """
        提取详情信息

        Returns:
            Dict[str, Any]: 详情信息
        """
        try:
            # 获取标题
            name = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.CLASS_NAME, "pm-name"))
            ).text
            title = clean_text(name)
            asset_id_match = re.search(r"(\d+)", url)
            asset_id = asset_id_match.group(1) if asset_id_match else ""

            # 获取成交价格
            try:
                final_price = WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, "//*[@id='pageContainer']/div[2]/div[1]/div[2]/div[3]/div[3]/div[1]/div/div[2]"))
                ).text
            except:
                final_price = ''

            # 获取流拍信息
            bidding_result = self.driver.find_element(By.XPATH, "//*[@id='pageContainer']/div[2]/div[1]/div[2]/div[3]/div[1]/div[1]/div").text
            if '流拍' in bidding_result:
                if_unsold = '是'
                unsold_reason = '本标的物已流拍'
            else:
                if_unsold = '否'
                unsold_reason = ''

            # 获取其他信息
            info_html = self.driver.find_element(By.XPATH, "//*[@id='pageContainer']/div[2]/div[1]/div[2]/div[3]/div[1]/div[2]/div").text

            # 使用正则表达式提取信息
            end_time_match = re.search(r'\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}', info_html)
            end_time = end_time_match.group(0) if end_time_match else ''

            watched_match = re.search(r'(\d+)人围观', info_html)
            watched_times = watched_match.group(1) if watched_match else ''

            signin_match = re.search(r'(\d+)人报名', info_html)
            signin_times = signin_match.group(1) if signin_match else ''

            attention_match = re.search(r'(\d+)人关注', info_html)
            attention_times = attention_match.group(1) if attention_match else ''

            bidding_info = self.driver.find_element(By.XPATH, "//*[@id='pageContainer']/div[2]/div[1]/div[2]/div[4]/div[2]/div/div[1]/div/ul").text #竞价信息
            bidding_info = bidding_info.replace(',', '')
            make_up_match = re.search(r'加价幅度：\n￥\n(\d+)\n', bidding_info)
            make_up = make_up_match.group(1) if make_up_match else ''
            delay_cycle_match = re.search(r'延时周期：\n(\d+)分钟/次\n', bidding_info)
            delay_cycle = delay_cycle_match.group(1) if delay_cycle_match else ''
            # 判断是否为变卖
            if '变卖' in name:
                sell_off_price_match = re.search(r'变卖价：\n￥\n(\d+)\n', bidding_info)
                sell_off_price = sell_off_price_match.group(1) if sell_off_price_match else ''
                sell_off_cycle_match = re.search(r'变卖周期：\n(\d+)天\n', bidding_info)
                sell_off_cycle = sell_off_cycle_match.group(1) if sell_off_cycle_match else ''

                return {
                    '标的物ID': asset_id, '链接': url, '标题': title, '资产名称': title, '结束时间': end_time, '是否流拍': if_unsold, '流拍原因': unsold_reason,
                    '围观人数': watched_times, '报名人数': signin_times, '关注提醒人数': attention_times,
                    "成交价格": final_price, '起拍价格': '', '变卖价格': sell_off_price,
                    '加价幅度': make_up, '保证金': '', '竞价周期': '', '变卖周期': sell_off_cycle, '延时周期': delay_cycle,
                    '当前价': final_price, '起拍价': '', '保证金': '', '评估价': '', '加价幅度': make_up, '竞价周期': '', '延时周期': delay_cycle,
                }
            else:
                start_price_match = re.search(r'起拍价：\n￥\n(\d+)\n', bidding_info)
                start_price = start_price_match.group(1) if start_price_match else ''
                margin_match = re.search(r'保证金：\n￥\n(\d+)\n', bidding_info)
                margin = margin_match.group(1) if margin_match else ''
                bidding_cycle_match = re.search(r'竞价周期：\n(\d+)天\n', bidding_info)
                bidding_cycle = bidding_cycle_match.group(1) if bidding_cycle_match else ''

                detail = {
                    '标的物ID': asset_id, '链接': url, '标题': title, '资产名称': title, '结束时间': end_time, '是否流拍': if_unsold, '流拍原因': unsold_reason,
                    '围观人数': watched_times, '报名人数': signin_times, '关注提醒人数': attention_times,
                    "成交价格": final_price, '起拍价格': start_price, '变卖价格': '',
                    '加价幅度': make_up, '保证金': margin, '竞价周期': bidding_cycle, '变卖周期': '', '延时周期': delay_cycle
                }
                detail["起拍价"] = start_price
                detail["保证金"] = margin
                detail["竞价周期"] = bidding_cycle
                detail["延时周期"] = delay_cycle
                detail["当前价"] = final_price
                detail["评估价"] = ""
                detail.update(self.extract_basic_detail_blocks())
                return detail

        except Exception as e:
            self.logger.error(f"提取详情信息失败: {e}")
            return {}

    def extract_basic_detail_blocks(self) -> Dict[str, Any]:
        result = {
            "完整地址": "",
            "处置法院": "",
            "城市": "",
            "标的物介绍原文": "",
            "标的物详情描述": "",
            "竞买公告": "",
            "竞买须知": "",
            "拍卖公告": "",
            "图片链接": "[]",
            "图片数量": 0,
            "大家都在问_QA": "",
            "详情页截图路径": "",
            "正文截图路径": "",
            "附件索引": "[]",
            "资源目录": "",
            "解析状态": "待解析" if self.crawl_mode == "fast" else "",
        }

        try:
            address_element = self.driver.find_element(By.XPATH, "//*[contains(@class,'pm-location')]//em")
            result["完整地址"] = clean_text(address_element.text)
            city_match = re.search(r"([^\s]+市)", result["完整地址"])
            if city_match:
                result["城市"] = city_match.group(1)
        except Exception:
            pass

        try:
            court_element = self.driver.find_element(
                By.XPATH,
                "//*[contains(@class,'pm-contact-container')]//dl[dt[contains(.,'处置机构')]]/dd"
            )
            result["处置法院"] = clean_text(court_element.text)
        except Exception:
            pass

        try:
            intro_element = self.driver.find_element(By.XPATH, "//*[@id='pmMainFloor']/ul/li[1]")
            result["标的物介绍原文"] = clean_text(intro_element.text)
            result["标的物详情描述"] = result["标的物介绍原文"]
        except Exception:
            pass

        try:
            image_elements = self.driver.find_elements(By.XPATH, "//img")
            image_urls = filter_image_urls([img.get_attribute("src") or img.get_attribute("data-lazy-img") for img in image_elements])
            result["图片链接"] = dump_json(image_urls)
            result["图片数量"] = len(image_urls)
        except Exception:
            pass

        return result

    def enrich_detail_from_assets(self, folder_path: str, detail_info: Dict[str, Any]) -> Dict[str, Any]:
        enriched: Dict[str, Any] = {}
        intro_text = clean_text(
            detail_info.get("标的物介绍原文", "")
            or detail_info.get("拍卖公告", "")
            or detail_info.get("竞买公告", "")
        )
        sections = parse_intro_sections(intro_text)
        enriched["拍品名称"] = sections.get("拍品名称", detail_info.get("标题", ""))
        enriched["拍品所有人"] = sections.get("拍品所有人", "")
        enriched["拍品现状"] = sections.get("拍品现状", "")
        enriched["租赁情况"] = sections.get("租赁情况", "")
        enriched["钥匙/占用情况"] = sections.get("钥匙/占用情况", sections.get("租赁情况", ""))
        enriched["户籍/工商注册"] = sections.get("户籍/工商注册", "")
        enriched["权利限制状况及抵押状况"] = sections.get("权利限制状况及抵押状况", "")
        enriched["成交后提供的文件"] = sections.get("成交后提供的文件", "")
        enriched["拍品介绍"] = sections.get("拍品介绍", "")
        enriched["房屋权属状况"] = sections.get("房屋权属状况", "")
        enriched["土地权属状况"] = sections.get("土地权属状况", "")

        enriched["标的物名称"] = enriched["拍品名称"] or detail_info.get("标题", "")
        enriched["权利来源"] = sections.get("权利来源", "")
        enriched["权证情况"] = sections.get("权证情况", "")
        enriched["被执行人"] = enriched["拍品所有人"]
        enriched["钥匙"] = enriched["钥匙/占用情况"]
        enriched["户籍注册"] = enriched["户籍/工商注册"]
        enriched["欠费情况"] = sections.get("欠费情况", "")
        enriched["提供文件"] = enriched["成交后提供的文件"]

        merged_text = clean_text(
            "\n".join(
                [
                    detail_info.get("标的物介绍原文", ""),
                    detail_info.get("拍卖公告", ""),
                    detail_info.get("竞买公告", ""),
                    detail_info.get("竞买须知", ""),
                ]
            )
        )
        attachment_texts = []
        for attachment_path in list_attachment_files(folder_path):
            attachment_text = extract_attachment_text(attachment_path)
            if attachment_text:
                attachment_texts.append(attachment_text)
                merged_text += "\n" + attachment_text
        enriched["附件文本"] = clean_text("\n".join(attachment_texts))

        survey_fields = self.extract_fields_from_survey_xlsx(folder_path)
        labeled_fields = extract_labeled_fields(merged_text)
        pdf_fields = extract_pdf_fields(merged_text)
        rights_fields = extract_rights_status_text(merged_text)

        field_order = [
            "标的物名称", "权利来源", "权证情况", "被执行人", "钥匙", "户籍注册", "欠费情况", "提供文件",
            "建筑面积", "房屋类型", "房屋用途", "总层数", "核准日期", "所有权来源", "土地用途", "土地性质", "使用期限", "所在层",
        ]
        for field in field_order:
            current = clean_text(str(enriched.get(field, "")))
            candidate = survey_fields.get(field) or labeled_fields.get(field) or pdf_fields.get(field, "")
            if not current and candidate:
                enriched[field] = normalize_field_value(field, candidate)
            elif current:
                enriched[field] = normalize_field_value(field, current)

        if not enriched.get("标的物名称"):
            enriched["标的物名称"] = detail_info.get("标题", "")

        enriched.update(survey_fields)
        for key, value in rights_fields.items():
            if value and not enriched.get(key):
                enriched[key] = value
        enriched = postprocess_structured_fields(enriched)
        enriched.update(extract_region_fields(detail_info.get("完整地址", "")))

        return enriched

    def extract_fields_from_survey_xlsx(self, folder_path: str) -> Dict[str, Any]:
        result: Dict[str, Any] = {}
        if not folder_path or not os.path.exists(folder_path):
            return result

        xlsx_files = [
            os.path.join(folder_path, name)
            for name in os.listdir(folder_path)
            if name.endswith(".xlsx") and "调查情况表" in name
        ]
        if not xlsx_files:
            return result

        try:
            df = pd.read_excel(xlsx_files[0], header=None).fillna("")
            rows = [[str(v) for v in row] for row in df.values.tolist()]
            result.update(parse_survey_table_rows(rows))

            row_texts = []
            for row in rows:
                cells = [clean_text(str(v)) for v in row if clean_text(str(v))]
                if cells:
                    row_texts.append(" ".join(cells))

            parsed = extract_labeled_fields("\n".join(row_texts))
            for key, value in parsed.items():
                if result.get(key):
                    continue
                normalized = normalize_field_value(key, value)
                if normalized:
                    result[key] = normalized
        except Exception:
            return result

        return result
    def extract_property_survey_table(self, asset_name: str) -> None:
        """
        提取标的物调查表

        Args:
            asset_name: 资产名称
        """
        if not asset_name:
            return

        try:
            # 创建文件夹
            folder_path = self._resolve_asset_folder(asset_name)

            # 提取标的物调查表
            try:
                table_content = WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, f"//*[@id='pmMainFloor']/ul/li[1]/div[2]/div"))
                )
                table_html = table_content.get_attribute('outerHTML')
                soup = BeautifulSoup(table_html, 'lxml')
                tables = soup.find_all('table')
                if tables:
                    table = tables[0]
                    df = pd.read_html(str(table))[0]
                    file_path = os.path.join(folder_path, "拍卖标的物调查情况表（房产）.xlsx")
                    df.to_excel(file_path, index=False)
                    self.logger.info(f"标的物调查表已保存到: {file_path}")
                else:
                    self.logger.info("未找到表格内容")
            except Exception as e:
                self.logger.info("未找到表格内容")

        except Exception as e:
            self.logger.error(f"提取标的物调查表失败: {e}")
    def download_attachments(self, asset_name: str) -> str:
        """
        下载附件和图片

        Args:
            asset_name: 资产名称
        """
        if not asset_name:
            return ""

        try:
            # 创建文件夹
            folder_path = self._resolve_asset_folder(asset_name)

            # # 检查是否有配资服务
            # try:
            #     self.driver.find_element(By.CLASS_NAME, "pm-pzfw")
            #     third_div = '6'
            # except:
            #     third_div = '5'

            # 下载PDF文件
            try:
                file_list = self.driver.find_elements(By.XPATH, f"//*[@id='pmMainFloor']/ul/li[1]/div[1]/div/div/div[1]/ul/li")
                for file in file_list:
                    file_url = file.find_element(By.XPATH, ".//*[@id='openAttachmentTag']").get_property("href")
                    file_name = file.find_element(By.XPATH, ".//*[@id='openAttachmentTag']").text
                    file_path = os.path.join(folder_path, file_name)
                    self.data_storage.download_file(file_url, file_path)
            except:
                pass

            # 下载图片
            try:
                img_list = self.driver.find_elements(By.XPATH, f"//*[@id='pmMainFloor']/ul/li[1]/div[2]/a")
                for i, img in enumerate(img_list):
                    img_url = img.get_attribute('href')
                    img_path = os.path.join(folder_path, f"{i}.jpg")
                    self.data_storage.download_file(img_url, img_path)
            except:
                pass
            return folder_path

        except Exception as e:
            self.logger.error(f"下载附件失败: {e}")
            return ""

    def extract_notice_info(self, asset_name: str = None):
        """
        提取竞买公告和竞买须知

        Args:
            asset_name: 资产名称，用于创建文件夹

        """
        if not asset_name:
            return {}

        try:
            find_floors = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//*[@id='pmMainFloor']/ul"))
            )
            note_content = ""
            instruction_content = ""

            if not note_content:
                try:
                    find_notice = find_floors.find_element(By.XPATH, "./li[2]")
                    try:
                        notice_content = find_notice.find_element(By.XPATH, "./div[2]")
                        html_content = BeautifulSoup(notice_content.get_attribute('outerHTML'), 'lxml')
                    except Exception:
                        html_content = BeautifulSoup(find_notice.get_attribute('outerHTML'), 'lxml')
                    note_content = clean_text(html_content.get_text(separator='\n', strip=True))
                except Exception:
                    pass

            if not instruction_content:
                try:
                    find_rule = find_floors.find_element(By.XPATH, "./li[3]")
                    try:
                        rule_content = find_rule.find_element(By.XPATH, "./div[2] | ./div")
                        html_content = BeautifulSoup(rule_content.get_attribute('outerHTML'), 'lxml')
                    except Exception:
                        html_content = BeautifulSoup(find_rule.get_attribute('outerHTML'), 'lxml')
                    instruction_content = clean_text(html_content.get_text(separator='\n', strip=True))
                except Exception:
                    pass

            if not note_content or not instruction_content:
                floor_items = find_floors.find_elements(By.XPATH, "./li")
                for item in floor_items:
                    item_text = clean_text(item.text)
                    if not item_text:
                        continue
                    if (not note_content) and ("拍卖公告" in item_text or ("竞买公告" in item_text and "竞拍流程" not in item_text)):
                        note_content = item_text
                    if (not instruction_content) and ("拍卖须知" in item_text or "竞买须知" in item_text):
                        instruction_content = item_text

            result = {
                "竞买公告": note_content,
                "竞买须知": instruction_content,
                "拍卖公告": note_content,
            }

            # 如果有资产名称，保存到Excel文件
            if asset_name:
                folder_path = self._resolve_asset_folder(asset_name)
                file_path = os.path.join(folder_path, "竞买公告和竞买须知.xlsx")
                pd.DataFrame([result]).to_excel(file_path, index=False)
                self.logger.info(f"竞买公告和竞买须知已保存到: {file_path}")
            return result
        except Exception as e:
            self.logger.error(f"提取竞买公告和竞买须知失败: {e}")
            return {}

    def extract_bidding_info(self, asset_name: str = None):
        """
        提取竞价记录信息

        Args:
            asset_name: 资产名称，用于创建文件夹
        """
        if not asset_name:
            return

        try:
            # 查找竞价记录区域
            bidding_record = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//*[contains(@class, 'floor') and contains(@class, 'floor-bid')]"))
            )

            bidding_records = []

            while True:
                try:
                    # 获取当前页面的竞价记录
                    table_content = bidding_record.get_attribute('outerHTML')
                    soup = BeautifulSoup(table_content, 'lxml')
                    bidding_list = soup.find('tbody').find_all('tr')

                    for row in bidding_list:
                        columns = row.find_all('td')
                        if len(columns) >= 4:
                            status = columns[0].get_text(strip=True)
                            bidding_code = columns[1].get_text(strip=True)
                            price = columns[2].get_text(strip=True)
                            bidding_time = columns[3].get_text(strip=True)

                            bidding_records.append({
                                "状态": status,
                                "价格": price,
                                "竞拍人": bidding_code,
                                "时间": bidding_time
                            })

                    # 尝试查找下一页按钮
                    try:
                        bidding_pager = self.driver.find_element(By.CLASS_NAME, "index_ui_pager__x0-LU")

                        next_button = bidding_pager.find_element(By.CLASS_NAME, "index_ui_pager_next__Rqo9l ")

                        # 检查下一页按钮是否可用
                        if "index_disabled__bPJgO" in next_button.get_attribute('class'):
                            self.logger.info("已到达竞价记录最后一页")
                            break

                    except Exception as e:
                        self.logger.info("无下一页信息或已到达最后一页")
                        break

                    # 点击下一页
                    try:
                        self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", next_button)
                        WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable(next_button)).click()
                        sleep(random.uniform(2, 4))  # 随机等待时间
                        self.logger.info("正在查找下一页出价信息...")
                    except Exception as e:
                        self.logger.warning(f"翻页失败: {e}")
                        break

                except Exception as e:
                    self.logger.error(f"处理竞价记录页面时出错: {e}")
                    break

            # 保存竞价记录到Excel文件
            if bidding_records:
                folder_path = self._resolve_asset_folder(asset_name)
                file_path = os.path.join(folder_path, "出价记录.xlsx")
                bidding_df = pd.DataFrame(bidding_records)
                bidding_df.to_excel(file_path, index=False)
                self.logger.info(f"竞价记录已保存到: {file_path}，共 {len(bidding_records)} 条记录")
            else:
                self.logger.info("未找到竞价记录")

        except Exception as e:
            self.logger.error(f"提取竞价记录失败: {e}")

    def extract_priority_purchaser(self, asset_name: str = None):
        """
        提取优先购买权人

        Args:
            asset_name: 资产名称，用于创建文件夹
        """
        if not asset_name:
            return

        try:
            # 查找优先购买权人区域
            priority_purchase = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.CLASS_NAME, "purchaserList"))
            )

            # 获取优先购买权人的HTML内容
            people_content = priority_purchase.get_attribute('outerHTML')
            people_content = BeautifulSoup(people_content, 'lxml')
            purchasers = people_content.find_all('table')

            if purchasers:
                # 获取第一个表格（通常只有一个优先购买权人表格）
                purchaser = purchasers[0]

                # 使用pandas读取HTML表格
                df = pd.read_html(str(purchaser))[0]

                # 保存到Excel文件
                folder_path = self._resolve_asset_folder(asset_name)
                file_path = os.path.join(folder_path, "优先购买权人.xlsx")
                df.to_excel(file_path, index=False)

                self.logger.info(f"优先购买权人信息已保存到: {file_path}")
                self.logger.info(f"共找到 {len(df)} 条优先购买权人记录")
            else:
                self.logger.info("未找到优先购买权人表格")

        except Exception as e:
            self.logger.info("未找到优先购买权人信息")
            self.logger.debug(f"提取优先购买权人失败: {e}")

    def transfer_to_start_page(self, current_page: int, target_page: int) -> int:
        """
        跳转到指定页面

        Args:
            current_page: 当前页码
            target_page: 目标页码

        Returns:
            int: 实际到达的页码
        """
        consecutive_failures = 0
        while consecutive_failures < 3 and current_page < target_page:
            try:
                current_page = int(self.driver.find_element(By.CLASS_NAME, "ui-pager-current").text)
                initial_signature = self.get_page_content_signature()

                a_elements = self.driver.find_elements(By.XPATH, '//div[@class="ui-pager"]/a')
                next_button = self.driver.find_element(By.CLASS_NAME, "ui-pager-next")
                fast_button = a_elements[-2] if len(a_elements) >= 2 else None

                use_fast_jump = current_page < target_page - 3 and fast_button is not None
                expected_delta = 6 if (use_fast_jump and current_page == 1) else (3 if use_fast_jump else 1)
                click_target = fast_button if use_fast_jump else next_button

                self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", click_target)
                sleep(random.uniform(1, 2))

                ActionChains(self.driver).move_to_element(click_target).perform()
                sleep(random.uniform(0.5, 1))

                try:
                    click_target.click()
                except Exception:
                    self.driver.execute_script("arguments[0].click();", click_target)

                transferred_page = current_page
                try:
                    WebDriverWait(self.driver, 10).until(
                        lambda d: int(d.find_element(By.CLASS_NAME, "ui-pager-current").text) != current_page
                    )
                    transferred_page = int(self.driver.find_element(By.CLASS_NAME, "ui-pager-current").text)
                except Exception:
                    if self.wait_for_page_change(initial_signature, max_wait=8):
                        transferred_page = int(self.driver.find_element(By.CLASS_NAME, "ui-pager-current").text)

                if transferred_page >= current_page + 1:
                    current_page = transferred_page
                    self.logger.info(f"成功跳转到第 {current_page} 页")
                    consecutive_failures = 0
                else:
                    self.logger.warning(
                        f"跳转页面失败，当前页码: {current_page}，目标页码: {target_page}，"
                        f"预期至少前进 {expected_delta} 页，实际到达: {transferred_page}"
                    )
                    consecutive_failures += 1
                    sleep(random.uniform(2, 4))
            except Exception as e:
                self.logger.error(f"跳转页面失败: {e}")
                consecutive_failures += 1
                try:
                    self.driver.refresh()
                    WebDriverWait(self.driver, 15).until(
                        EC.presence_of_element_located((By.CLASS_NAME, "ui-pager-current"))
                    )
                    sleep(random.uniform(2, 4))
                    current_page = int(self.driver.find_element(By.CLASS_NAME, "ui-pager-current").text)
                except Exception as refresh_error:
                    self.logger.warning(f"翻页失败后刷新列表页也失败: {refresh_error}")
                    break

        return current_page

    def get_page_content_signature(self) -> str:
        """
        获取页面内容签名，用于检测页面是否发生变化

        Returns:
            str: 页面内容的签名
        """
        try:
            # 获取列表项的数量和部分内容作为签名
            list_elements = self._find_listing_elements()
            if not list_elements:
                return ""

            # 取前几个元素的文本作为签名
            signature_parts = []
            for i, element in enumerate(list_elements[:3]):  # 只取前3个元素
                try:
                    text = element.text[:50]  # 只取前50个字符
                    signature_parts.append(text)
                except:
                    signature_parts.append(f"element_{i}")

            return "|".join(signature_parts)

        except Exception as e:
            self.logger.debug(f"获取页面签名失败: {e}")
            return ""

    def wait_for_page_change(self, initial_signature: str, max_wait: int = 15) -> bool:
        """
        等待页面内容发生变化

        Args:
            initial_signature: 初始页面签名
            max_wait: 最大等待时间（秒）

        Returns:
            bool: 页面是否发生变化
        """
        start_time = time.time()

        while time.time() - start_time < max_wait:
            try:
                # 等待一段时间再检查
                sleep(random.uniform(1, 2))

                # 获取当前页面签名
                current_signature = self.get_page_content_signature()

                # 检查页面是否发生变化
                if current_signature and current_signature != initial_signature:
                    self.logger.debug("检测到页面内容发生变化")
                    return True

            except Exception as e:
                self.logger.debug(f"等待页面变化时出错: {e}")

        self.logger.warning(f"等待页面变化超时 ({max_wait} 秒)")
        return False
