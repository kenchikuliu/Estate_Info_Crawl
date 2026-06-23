# -*- coding: utf-8 -*-
"""Run the JD/Ali region pipeline into a D: drive output root.

The pipeline mirrors the Beijing/Guangdong handoff shape:
index -> detail backfill -> optional Ali schema alignment -> merge/dedupe ->
surrounding/intro organization.
"""
from __future__ import annotations

import argparse
import json
import subprocess
import sys
from datetime import datetime
from pathlib import Path
from typing import Dict, Iterable, List


PROJECT_ROOT = Path(__file__).resolve().parents[1]

REGION_LABELS = {
    "sh": "上海",
    "zj": "浙江",
    "hb": "湖北",
}

DEFAULT_AREAS_JSON = Path(
    r"C:\Users\Administrator\Downloads\_repo_inspect\Administrative-divisions-of-China\dist\areas.json"
)
DEFAULT_LEGACY_GB2260 = Path(
    r"C:\Users\Administrator\Downloads\_repo_inspect\auction-mcp\gb2260_200712.json"
)


def run_command(command: List[str], log_path: Path, dry_run: bool = False) -> None:
    log_path.parent.mkdir(parents=True, exist_ok=True)
    printable = " ".join(command)
    print(f"RUN {printable}", flush=True)
    if dry_run:
        return

    with log_path.open("a", encoding="utf-8") as log_file:
        log_file.write(f"\n\n[{datetime.now():%Y-%m-%d %H:%M:%S}] RUN {printable}\n")
        log_file.flush()
        result = subprocess.run(
            command,
            cwd=PROJECT_ROOT,
            stdout=log_file,
            stderr=subprocess.STDOUT,
            text=True,
            check=False,
        )
        log_file.write(f"\n[{datetime.now():%Y-%m-%d %H:%M:%S}] EXIT {result.returncode}\n")
    if result.returncode != 0:
        raise RuntimeError(f"Command failed with exit {result.returncode}: {printable}. See {log_path}")


def py_command(script: str, *args: str) -> List[str]:
    return [sys.executable, str(PROJECT_ROOT / script), *map(str, args)]


def paths_for_region(root: Path, key: str) -> Dict[str, Path]:
    label = REGION_LABELS[key]
    region_root = root / key
    index_root = region_root / "index"
    detail_root = region_root / "detail"
    final_root = region_root / "final"
    logs_root = region_root / "logs"
    return {
        "region_root": region_root,
        "index_root": index_root,
        "detail_root": detail_root,
        "final_root": final_root,
        "logs_root": logs_root,
        "jd_index_dir": index_root / "jd_parts",
        "jd_index": index_root / f"京东法拍房_{label}_全量索引_API合并.xlsx",
        "jd_index_checkpoint": index_root / f"京东法拍房_{label}_API城市索引.checkpoint.json",
        "jd_index_stats": index_root / f"京东法拍房_{label}_API城市索引_统计.xlsx",
        "jd_detail": detail_root / f"京东法拍房_{label}_详情回填_API.xlsx",
        "jd_detail_jsonl": detail_root / f"京东法拍房_{label}_详情回填_API.jsonl",
        "jd_detail_checkpoint": detail_root / f"京东法拍房_{label}_详情回填_API.checkpoint.json",
        "ali_index_dir": index_root / "ali_parts",
        "ali_index": index_root / f"阿里法拍房_{label}_全量索引.xlsx",
        "ali_index_checkpoint": index_root / f"阿里法拍房_{label}_索引.checkpoint.json",
        "ali_index_stats": index_root / f"阿里法拍房_{label}_索引统计.xlsx",
        "ali_detail": detail_root / f"阿里法拍房_{label}_详情回填_API.xlsx",
        "ali_detail_jsonl": detail_root / f"阿里法拍房_{label}_详情回填_API.jsonl",
        "ali_detail_checkpoint": detail_root / f"阿里法拍房_{label}_详情回填_API.checkpoint.json",
        "ali_aligned": detail_root / f"阿里法拍房_{label}_详情回填_API_京东字段对齐.xlsx",
        "ali_aligned_report": detail_root / f"阿里法拍房_{label}_详情回填_API_京东字段对齐.schema_verify.json",
        "merged": final_root / f"法拍房_{label}_京东阿里合并去重.xlsx",
        "merge_report": final_root / f"法拍房_{label}_京东阿里合并去重.report.json",
        "organized": final_root / f"法拍房_{label}_京东阿里合并去重_周边标的整理版.xlsx",
        "organized_report": final_root / f"法拍房_{label}_京东阿里合并去重_周边标的整理版.report.json",
    }


def write_manifest(root: Path, regions: Iterable[str]) -> None:
    manifest = {
        "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "root": str(root),
        "regions": {key: {name: str(path) for name, path in paths_for_region(root, key).items()} for key in regions},
    }
    root.mkdir(parents=True, exist_ok=True)
    (root / "pipeline_manifest.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")


def run_region(args: argparse.Namespace, key: str) -> None:
    label = REGION_LABELS[key]
    paths = paths_for_region(args.root, key)
    for directory_name in ("index_root", "detail_root", "final_root", "logs_root"):
        paths[directory_name].mkdir(parents=True, exist_ok=True)

    if args.stage in {"all", "index"}:
        run_command(
            py_command(
                "scripts/jd_api_index_crawl.py",
                "--province-key",
                key,
                "--output-dir",
                str(paths["jd_index_dir"]),
                "--checkpoint",
                str(paths["jd_index_checkpoint"]),
                "--merge-output",
                str(paths["jd_index"]),
                "--stats-output",
                str(paths["jd_index_stats"]),
                "--workers",
                str(args.jd_index_workers),
                "--page-size",
                str(args.jd_page_size),
                "--max-bucket-size",
                str(args.jd_max_bucket_size),
                "--start-year",
                str(args.start_year),
                "--end-year",
                str(args.end_year),
                *(["--with-config"] if args.jd_with_config else []),
            ),
            paths["logs_root"] / "jd_index.log",
            args.dry_run,
        )
        run_command(
            py_command(
                "scripts/ali_sf_index_crawl.py",
                "--province-key",
                key,
                "--cities",
                args.ali_cities,
                "--categories",
                args.ali_categories,
                "--output-dir",
                str(paths["ali_index_dir"]),
                "--checkpoint",
                str(paths["ali_index_checkpoint"]),
                "--merge-output",
                str(paths["ali_index"]),
                "--stats-output",
                str(paths["ali_index_stats"]),
                "--workers",
                str(args.ali_index_workers),
                "--page-chunk-size",
                str(args.ali_page_chunk_size),
                "--save-every-pages",
                str(args.ali_save_every_pages),
                "--max-pages-per-partition",
                str(args.ali_max_pages_per_partition),
                "--areas-json",
                str(args.areas_json),
                "--legacy-gb2260",
                str(args.legacy_gb2260),
            ),
            paths["logs_root"] / "ali_index.log",
            args.dry_run,
        )

    if args.stage in {"all", "detail"}:
        run_command(
            py_command(
                "scripts/jd_api_detail_backfill.py",
                "--input",
                str(paths["jd_index"]),
                "--output",
                str(paths["jd_detail"]),
                "--jsonl-output",
                str(paths["jd_detail_jsonl"]),
                "--checkpoint",
                str(paths["jd_detail_checkpoint"]),
                "--workers",
                str(args.jd_detail_workers),
                "--save-every",
                str(args.detail_save_every),
                "--raw-text-limit",
                str(args.raw_text_limit),
                *(["--limit", str(args.jd_detail_limit)] if args.jd_detail_limit else []),
                "--include-unfinished-index",
            ),
            paths["logs_root"] / "jd_detail.log",
            args.dry_run,
        )
        run_command(
            py_command(
                "scripts/ali_sf_api_detail_backfill.py",
                "--input",
                str(paths["ali_index"]),
                "--output",
                str(paths["ali_detail"]),
                "--jsonl-output",
                str(paths["ali_detail_jsonl"]),
                "--checkpoint",
                str(paths["ali_detail_checkpoint"]),
                "--workers",
                str(args.ali_detail_workers),
                "--save-every",
                str(args.detail_save_every),
                "--raw-text-limit",
                str(args.raw_text_limit),
                *(["--limit", str(args.ali_detail_limit)] if args.ali_detail_limit else []),
                "--include-unfinished-index",
            ),
            paths["logs_root"] / "ali_detail.log",
            args.dry_run,
        )
        run_command(
            py_command(
                "scripts/ali_sf_align_to_jd_schema.py",
                "--ali-input",
                str(paths["ali_detail"]),
                "--ali-index",
                str(paths["ali_index"]),
                "--ali-detail-jsonl",
                str(paths["ali_detail_jsonl"]),
                "--jd-schema",
                str(paths["jd_detail"]),
                "--output",
                str(paths["ali_aligned"]),
                "--report",
                str(paths["ali_aligned_report"]),
            ),
            paths["logs_root"] / "ali_align.log",
            args.dry_run,
        )

    if args.stage in {"all", "final"}:
        run_command(
            py_command(
                "scripts/merge_jd_ali_region_tables.py",
                "--jd",
                str(paths["jd_detail"]),
                "--ali",
                str(paths["ali_aligned"]),
                "--output",
                str(paths["merged"]),
                "--report",
                str(paths["merge_report"]),
            ),
            paths["logs_root"] / "merge.log",
            args.dry_run,
        )
        run_command(
            py_command(
                "scripts/enrich_surrounding_intro_fields.py",
                "--input",
                str(paths["merged"]),
                "--output",
                str(paths["organized"]),
                "--report",
                str(paths["organized_report"]),
                "--tmpdir",
                str(paths["final_root"] / "tmp"),
            ),
            paths["logs_root"] / "organize.log",
            args.dry_run,
        )
    print(f"region_done {label} root={paths['region_root']}", flush=True)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=Path(r"D:\temp\estate_sh_zj_hb_20260623"))
    parser.add_argument("--regions", default="sh,zj,hb")
    parser.add_argument("--stage", choices=["all", "index", "detail", "final"], default="all")
    parser.add_argument("--dry-run", action="store_true")
    parser.add_argument("--start-year", type=int, default=2000)
    parser.add_argument("--end-year", type=int, default=2035)
    parser.add_argument("--jd-index-workers", type=int, default=4)
    parser.add_argument("--jd-page-size", type=int, default=1000)
    parser.add_argument("--jd-max-bucket-size", type=int, default=9000)
    parser.add_argument("--jd-with-config", action="store_true")
    parser.add_argument("--ali-cities", default="all")
    parser.add_argument("--ali-categories", default="all")
    parser.add_argument("--ali-index-workers", type=int, default=8)
    parser.add_argument("--ali-page-chunk-size", type=int, default=100)
    parser.add_argument("--ali-save-every-pages", type=int, default=1000)
    parser.add_argument("--ali-max-pages-per-partition", type=int, default=0)
    parser.add_argument("--areas-json", type=Path, default=DEFAULT_AREAS_JSON)
    parser.add_argument("--legacy-gb2260", type=Path, default=DEFAULT_LEGACY_GB2260)
    parser.add_argument("--jd-detail-workers", type=int, default=12)
    parser.add_argument("--ali-detail-workers", type=int, default=16)
    parser.add_argument("--jd-detail-limit", type=int, default=0)
    parser.add_argument("--ali-detail-limit", type=int, default=0)
    parser.add_argument("--detail-save-every", type=int, default=500)
    parser.add_argument("--raw-text-limit", type=int, default=2000)
    args = parser.parse_args()
    args.regions = [part.strip().lower() for part in args.regions.split(",") if part.strip()]
    unknown = [key for key in args.regions if key not in REGION_LABELS]
    if unknown:
        raise ValueError(f"unknown regions: {', '.join(unknown)}")
    return args


def main() -> None:
    args = parse_args()
    args.root.mkdir(parents=True, exist_ok=True)
    write_manifest(args.root, args.regions)
    for key in args.regions:
        run_region(args, key)


if __name__ == "__main__":
    main()
