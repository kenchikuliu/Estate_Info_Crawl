# -*- coding: utf-8 -*-
"""
主程序入口
"""
import argparse
import sys
from pathlib import Path
from typing import List
import pandas as pd
from config import Config
from utils.data_storage import DataStorage
from utils.jd_detail_parser import (
    POI_COLUMN_SPECS,
    clean_text,
    enrich_poi_distances,
    extract_coordinates,
    extract_region_fields,
    geocode_address,
)
from scripts import ali_sf_index_crawl

def run_jd_auction_spider(start_page: int = 1, max_pages: int = None, province: str = None, city: str = None, cutoff_time: str = None, resume_from_archive: bool = False, max_items: int = None, output_suffix: str = "", crawl_mode: str = "fast", backfill_file: str = "", backfill_output: str = "", backfill_limit: int = None, backfill_all: bool = False, backfill_start: int = 0, backfill_offline_only: bool = False) -> None:
    """
    运行京东法拍房爬虫

    Args:
        start_page: 开始页码
        max_pages: 最大爬取页数
        province: 要爬取的省份
        city: 要爬取的城市
        cutoff_time: 截止时间，格式为"YYYY年MM月DD日 HH:MM:SS"
        resume_from_archive: 是否从存档恢复爬取
    """
    print("=" * 50)
    print("京东法拍房爬虫")
    print("=" * 50)
    print("使用说明：")
    print("1. 程序会自动打开京东法拍网站")
    print("2. 如未登录，请在浏览器中手动完成登录")
    print("3. 登录完成后按回车键继续")
    if resume_from_archive:
        print("4. 存档恢复模式已启用，将从上次爬取停止的位置继续")
    print("=" * 50)

    from spiders.jd_auction_spider import JDAuctionSpider

    spider = JDAuctionSpider(
        start_page=start_page,
        max_pages=max_pages,
        province=province,
        city=city,
        cutoff_time=cutoff_time,
        resume_from_archive=resume_from_archive,
        max_items=max_items,
        output_suffix=output_suffix,
        crawl_mode=crawl_mode,
    )
    if backfill_file:
        output_path = spider.backfill_from_excel(
            input_file=backfill_file,
            output_file=backfill_output,
            limit=backfill_limit,
            only_pending=not backfill_all,
            start_index=backfill_start,
            offline_only=backfill_offline_only,
        )
        print(f"回填完成，输出文件: {output_path}")
    else:
        spider.start()

def run_lianjia_spider(districts: List[str] = None, max_pages: int = None) -> None:
    """
    运行链家二手房爬虫

    Args:
        districts: 要爬取的区域列表
        max_pages: 每个区域最大爬取页数
    """
    print("=" * 50)
    print("链家二手房爬虫")
    print("=" * 50)
    print("使用说明：")
    print("1. 程序会自动打开浏览器")
    print("2. 请在浏览器中完成链家网站的登录")
    print("3. 登录完成后按回车键继续")
    print("=" * 50)

    from spiders.lianjia_spider import LianjiaSpider

    spider = LianjiaSpider(districts=districts, max_pages=max_pages)
    spider.start()


def run_ali_spider(
    province_key: str = "gd",
    cities: str = "all",
    categories: str = "all",
    output_dir: str = "output",
    checkpoint: str = "",
    merge_output: str = "",
    stats_output: str = "",
    workers: int = 8,
    page_chunk_size: int = 100,
    save_every_pages: int = 1000,
    sort: str = "1",
    status_orders: str = "",
    with_public_bid_detail: bool = False,
    public_bid_detail_workers: int = 8,
    areas_json: str = "",
    legacy_gb2260: str = "",
    max_pages_per_partition: int = 0,
) -> None:
    print("=" * 50)
    print("阿里法拍房爬虫")
    print("=" * 50)
    print("使用说明：")
    print("1. 当前版本通过阿里 H5 列表接口抓取索引层数据")
    print("2. 如需更完整字段，可在抓取后继续做详情回填")
    print("=" * 50)

    ali_args = [
        f"--province-key={province_key}",
        f"--cities={cities}",
        f"--categories={categories}",
        f"--output-dir={output_dir}",
        f"--workers={workers}",
        f"--page-chunk-size={page_chunk_size}",
        f"--save-every-pages={save_every_pages}",
        f"--sort={sort}",
        f"--public-bid-detail-workers={public_bid_detail_workers}",
        f"--max-pages-per-partition={max_pages_per_partition}",
    ]
    if checkpoint:
        ali_args.append(f"--checkpoint={checkpoint}")
    if merge_output:
        ali_args.append(f"--merge-output={merge_output}")
    if stats_output:
        ali_args.append(f"--stats-output={stats_output}")
    if status_orders:
        ali_args.append(f"--status-orders={status_orders}")
    if areas_json:
        ali_args.append(f"--areas-json={areas_json}")
    if legacy_gb2260:
        ali_args.append(f"--legacy-gb2260={legacy_gb2260}")
    if with_public_bid_detail:
        ali_args.append("--with-public-bid-detail")

    ali_sf_index_crawl.main(ali_args)

def show_available_districts() -> None:
    """
    显示可用的区域
    """
    print("可用的深圳区域：")
    for district, sub_districts in Config.SHENZHEN_DISTRICTS.items():
        print(f"  {district}: {', '.join(sub_districts[:5])}{'...' if len(sub_districts) > 5 else ''}")

def show_available_provinces_cities() -> None:
    """
    显示可用的省份和城市
    """
    print("京东法拍房可用的省份和城市：")
    for province, cities in Config.JD_AUCTION_CONFIG["province_city_mapping"].items():
        if cities:
            print(f"  {province}: {', '.join(cities)}")
        else:
            print(f"  {province}: 仅支持省份级别")


def merge_backfill_excels(input_files: List[str], output_file: str) -> None:
    if not input_files:
        raise ValueError("缺少待合并文件")

    frames = []
    for path in input_files:
        df = pd.read_excel(path).fillna("").astype(object)
        df["_source_file"] = path
        frames.append(df)

    merged = pd.concat(frames, ignore_index=True)
    all_columns = [col for col in merged.columns if col != "_source_file"]

    def reduce_group(group: pd.DataFrame) -> pd.Series:
        result = {}
        status_series = group["解析状态"].astype(str) if "解析状态" in group.columns else pd.Series([], dtype=str)
        parsed_group = group[status_series == "已解析"]
        preferred = parsed_group if not parsed_group.empty else group

        for column in all_columns:
            values = preferred[column].astype(str) if column in preferred.columns else pd.Series([], dtype=str)
            chosen = ""
            for value in values:
                if value.strip():
                    chosen = value
                    break
            if not chosen and column in group.columns:
                for value in group[column].astype(str):
                    if value.strip():
                        chosen = value
                        break
            result[column] = chosen

        if "解析状态" in result:
            if (status_series == "已解析").any():
                result["解析状态"] = "已解析"
            elif status_series.astype(str).str.startswith("回填失败").any():
                failure = group.loc[status_series.astype(str).str.startswith("回填失败"), "解析状态"].astype(str).iloc[0]
                result["解析状态"] = failure

        return pd.Series(result)

    key_column = "链接" if "链接" in merged.columns else merged.columns[0]
    final_df = merged.groupby(key_column, dropna=False, sort=False).apply(reduce_group).reset_index(drop=True)
    DataStorage.write_excel_atomic(final_df, output_file)
    print(f"合并完成，输出文件: {output_file}")


def enrich_excel_with_location_poi(input_file: str, output_file: str = "", limit: int = None, start: int = 0) -> None:
    df = pd.read_excel(input_file).fillna("").astype(object)
    poi_columns = ["经度", "纬度"] + list(POI_COLUMN_SPECS.keys())
    for column in ["完整地址", "格式化地址", "省", "市", "区县", "坐标"] + poi_columns:
        if column not in df.columns:
            df[column] = ""

    work_df = df.iloc[start:]
    if limit:
        work_df = work_df.head(limit)

    for row_index in work_df.index:
        row = df.loc[row_index]
        address = clean_text(str(row.get("完整地址", ""))) or clean_text(str(row.get("格式化地址", "")))
        if not address:
            continue

        region_fields = extract_region_fields(address)
        for key in ["省", "市", "区县"]:
            if clean_text(str(row.get(key, ""))):
                continue
            if region_fields.get(key):
                df.at[row_index, key] = region_fields[key]
        if not clean_text(str(row.get("格式化地址", ""))) and region_fields.get("格式化地址"):
            df.at[row_index, "格式化地址"] = region_fields["格式化地址"]

        lon, lat = extract_coordinates(str(row.get("坐标", "")))
        existing_lon = clean_text(str(row.get("经度", "")))
        existing_lat = clean_text(str(row.get("纬度", "")))
        if existing_lon and existing_lat:
            try:
                lon = float(existing_lon)
                lat = float(existing_lat)
            except Exception:
                lon, lat = None, None

        if lon is None or lat is None:
            geo = geocode_address(address)
            if geo:
                for key, value in geo.items():
                    if value:
                        df.at[row_index, key] = value
                try:
                    lon = float(geo.get("经度", "") or 0)
                    lat = float(geo.get("纬度", "") or 0)
                except Exception:
                    lon, lat = None, None

        if lon is None or lat is None:
            continue

        df.at[row_index, "经度"] = str(lon)
        df.at[row_index, "纬度"] = str(lat)

        poi_results = enrich_poi_distances(lat, lon)
        for column in POI_COLUMN_SPECS:
            if clean_text(str(row.get(column, ""))):
                continue
            df.at[row_index, column] = poi_results.get(column, "")

    final_output = output_file or input_file
    DataStorage.write_excel_atomic(df, final_output)
    print(f"地址/POI 回填完成，输出文件: {final_output}")

def main():
    """
    主函数
    """
    parser = argparse.ArgumentParser(description="房产信息爬虫")
    parser.add_argument("--spider", choices=["jd", "lianjia", "ali", "both"], default=None,
                       help="选择要运行的爬虫: jd(京东法拍房), lianjia(链家二手房), ali(阿里法拍房), both(两个都运行)")

    # 京东法拍房参数
    parser.add_argument("--jd-start-page", type=int, default=None,
                       help="京东法拍房开始页码 (默认: 1)")
    parser.add_argument("--jd-max-pages", type=int, default=None,
                       help="京东法拍房最大爬取页数")
    parser.add_argument("--jd-province", type=str, default=None,
                       help="京东法拍房要爬取的省份")
    parser.add_argument("--jd-city", type=str, default=None,
                       help="京东法拍房要爬取的城市")
    parser.add_argument("--jd-cutoff-time", type=str, default=None,
                       help="京东法拍房截止时间，格式为'YYYY年MM月DD日 HH:MM:SS'，当拍卖结束时间早于此时间时停止爬取")
    parser.add_argument("--jd-resume-from-archive", action="store_true",
                       help="京东法拍房从存档恢复爬取，自动找到最后一条记录并从下一条开始")
    parser.add_argument("--jd-max-items", type=int, default=None,
                       help="京东法拍房本次最多处理多少个标的，便于快速验证")
    parser.add_argument("--jd-output-suffix", type=str, default="",
                       help="京东法拍房输出文件后缀，便于并行分片运行")
    parser.add_argument("--jd-crawl-mode", choices=["fast", "full"], default="fast",
                       help="京东法拍房抓取模式: fast(基础字段+截图+索引), full(完整下载解析)")
    parser.add_argument("--jd-backfill-file", type=str, default="",
                       help="对已抓取Excel执行二阶段回填解析")
    parser.add_argument("--jd-backfill-output", type=str, default="",
                       help="二阶段回填后的输出文件路径，默认覆盖输入文件")
    parser.add_argument("--jd-backfill-limit", type=int, default=None,
                       help="二阶段回填最多处理多少条")
    parser.add_argument("--jd-backfill-start", type=int, default=0,
                       help="二阶段回填从第几条待处理记录开始，便于并行分片")
    parser.add_argument("--jd-backfill-all", action="store_true",
                       help="二阶段回填时不只处理待解析行，而是处理全部行")
    parser.add_argument("--jd-backfill-offline-only", action="store_true",
                       help="仅根据已有资源目录/附件/公告文件离线重解析，不打开详情页")

    # 链家二手房参数
    parser.add_argument("--lianjia-districts", nargs="+", default=None,
                       help="链家二手房要爬取的区域列表")
    parser.add_argument("--lianjia-max-pages", type=int, default=None,
                       help="链家二手房每个区域最大爬取页数")

    # 其他参数
    parser.add_argument("--show-districts", action="store_true",
                       help="显示可用的深圳区域")
    parser.add_argument("--show-provinces", action="store_true",
                       help="显示京东法拍房可用的省份和城市")
    parser.add_argument("--merge-excels", nargs="+", default=None,
                       help="合并多个回填Excel文件")
    parser.add_argument("--merge-output", type=str, default="",
                       help="合并后的输出文件路径")
    parser.add_argument("--enrich-location-file", type=str, default="",
                       help="对 Excel 执行地址解析与周边 POI 距离回填")
    parser.add_argument("--enrich-location-output", type=str, default="",
                       help="地址/POI 回填输出文件路径，默认覆盖输入文件")
    parser.add_argument("--enrich-location-limit", type=int, default=None,
                       help="地址/POI 回填最多处理多少条")
    parser.add_argument("--enrich-location-start", type=int, default=0,
                       help="地址/POI 回填从第几条开始")

    # 阿里法拍房参数
    parser.add_argument("--ali-province-key", type=str, default="gd",
                       help="阿里法拍房省份预设 key，当前默认 gd")
    parser.add_argument("--ali-cities", type=str, default="all",
                       help="阿里法拍房城市列表，逗号分隔或 all")
    parser.add_argument("--ali-categories", type=str, default="all",
                       help="阿里法拍房类目列表，逗号分隔或 all")
    parser.add_argument("--ali-output-dir", type=str, default="output",
                       help="阿里法拍房输出目录")
    parser.add_argument("--ali-checkpoint", type=str, default="",
                       help="阿里法拍房 checkpoint 路径")
    parser.add_argument("--ali-merge-output", type=str, default="",
                       help="阿里法拍房合并后的 Excel 输出路径")
    parser.add_argument("--ali-stats-output", type=str, default="",
                       help="阿里法拍房统计 Excel 输出路径")
    parser.add_argument("--ali-workers", type=int, default=8,
                       help="阿里法拍房并发抓取 worker 数")
    parser.add_argument("--ali-page-chunk-size", type=int, default=100,
                       help="阿里法拍房分页分块大小")
    parser.add_argument("--ali-save-every-pages", type=int, default=1000,
                       help="阿里法拍房每处理多少页落盘一次")
    parser.add_argument("--ali-sort", type=str, default="1",
                       help="阿里法拍房排序参数")
    parser.add_argument("--ali-status-orders", type=str, default="",
                       help="阿里法拍房状态筛选，逗号分隔")
    parser.add_argument("--ali-with-public-bid-detail", action="store_true",
                       help="阿里法拍房是否拉取公开竞价详情")
    parser.add_argument("--ali-public-bid-detail-workers", type=int, default=8,
                       help="阿里法拍房公开竞价详情并发数")
    parser.add_argument("--ali-areas-json", type=str, default="",
                       help="阿里法拍房 areas.json 路径")
    parser.add_argument("--ali-legacy-gb2260", type=str, default="",
                       help="阿里法拍房 legacy gb2260 路径")
    parser.add_argument("--ali-max-pages-per-partition", type=int, default=0,
                       help="阿里法拍房每个分区最大页数，便于 smoke test")

    args = parser.parse_args()

    # 显示可用区域
    if args.show_districts:
        show_available_districts()
        return

    # 显示可用省份和城市
    if args.show_provinces:
        show_available_provinces_cities()
        return

    if args.merge_excels:
        if not args.merge_output:
            raise ValueError("使用 --merge-excels 时必须提供 --merge-output")
        merge_backfill_excels(args.merge_excels, args.merge_output)
        return

    if args.enrich_location_file:
        enrich_excel_with_location_poi(
            input_file=args.enrich_location_file,
            output_file=args.enrich_location_output,
            limit=args.enrich_location_limit,
            start=args.enrich_location_start,
        )
        return

    if not args.spider:
        parser.error("请使用 --spider 选择要运行的爬虫，或使用 --show-districts / --show-provinces 查看可用配置")

    if args.spider == "ali":
        run_ali_spider(
            province_key=args.ali_province_key,
            cities=args.ali_cities,
            categories=args.ali_categories,
            output_dir=args.ali_output_dir,
            checkpoint=args.ali_checkpoint,
            merge_output=args.ali_merge_output,
            stats_output=args.ali_stats_output,
            workers=args.ali_workers,
            page_chunk_size=args.ali_page_chunk_size,
            save_every_pages=args.ali_save_every_pages,
            sort=args.ali_sort,
            status_orders=args.ali_status_orders,
            with_public_bid_detail=args.ali_with_public_bid_detail,
            public_bid_detail_workers=args.ali_public_bid_detail_workers,
            areas_json=args.ali_areas_json,
            legacy_gb2260=args.ali_legacy_gb2260,
            max_pages_per_partition=args.ali_max_pages_per_partition,
        )
        return

    try:
        jd_province = args.jd_province or Config.JD_AUCTION_CONFIG["default_province"]
        jd_city = args.jd_city
        if args.spider in ["jd", "both"] and jd_city is None:
            default_city = Config.JD_AUCTION_CONFIG.get("default_city")
            allowed_cities = Config.JD_AUCTION_CONFIG["province_city_mapping"].get(jd_province, [])
            if default_city in allowed_cities:
                jd_city = default_city

        if args.spider in ["jd", "both"]:
            run_jd_auction_spider(
                start_page=args.jd_start_page,
                max_pages=args.jd_max_pages,
                province=jd_province,
                city=jd_city,
                cutoff_time=args.jd_cutoff_time,
                resume_from_archive=args.jd_resume_from_archive,
                max_items=args.jd_max_items,
                output_suffix=args.jd_output_suffix,
                crawl_mode=args.jd_crawl_mode,
                backfill_file=args.jd_backfill_file,
                backfill_output=args.jd_backfill_output,
                backfill_limit=args.jd_backfill_limit,
                backfill_all=args.jd_backfill_all,
                backfill_start=args.jd_backfill_start,
                backfill_offline_only=args.jd_backfill_offline_only,
            )

        if args.spider in ["lianjia", "both"]:
            run_lianjia_spider(
                districts=args.lianjia_districts,
                max_pages=args.lianjia_max_pages
            )

    except KeyboardInterrupt:
        print("\n程序被用户中断")
        sys.exit(0)
    except Exception as e:
        print(f"程序运行出错: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
