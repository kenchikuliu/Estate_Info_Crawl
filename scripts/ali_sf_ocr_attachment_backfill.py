# -*- coding: utf-8 -*-
"""
OCR cached Ali attachment files that normal text extraction could not parse.

This is intended for scanned PDFs and image-only DOCX attachments. It is local
only: no network download is performed.
"""
from __future__ import annotations

import argparse
import io
import json
import os
import re
import sys
import tempfile
import zipfile
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Set

import fitz
from PIL import Image
from rapidocr_onnxruntime import RapidOCR

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from scripts.ali_sf_attachment_text_backfill import (  # noqa: E402
    latest_rows_by_id,
    merge_prefer_existing,
    parse_attachment_fields,
)
from utils.jd_detail_parser import clean_text  # noqa: E402

DEFAULT_DETAIL_JSONL = PROJECT_ROOT / r"output\ali_sf_bj_full_h5_20260616\阿里法拍房_北京_详情回填_API.jsonl"
DEFAULT_ATTACHMENT_JSONL = PROJECT_ROOT / r"output\ali_sf_bj_full_h5_20260616\阿里法拍房_北京_附件正文回填_API.jsonl"
DEFAULT_OUTPUT_JSONL = PROJECT_ROOT / r"output\ali_sf_bj_full_h5_20260616\阿里法拍房_北京_附件正文OCR回填_API.jsonl"
DEFAULT_CACHE_ROOT = PROJECT_ROOT / r"output\ali_sf_bj_attachment_text_cache"

OCR_ENGINE: Optional[RapidOCR] = None
OFFICE_SUFFIXES = {".doc", ".docx", ".xls", ".xlsx"}


def get_ocr() -> RapidOCR:
    global OCR_ENGINE
    if OCR_ENGINE is None:
        OCR_ENGINE = RapidOCR()
    return OCR_ENGINE


def normalize_id(value: Any) -> str:
    text = str(value or "").strip()
    if text.endswith(".0") and text[:-2].isdigit():
        text = text[:-2]
    return text


def split_joined(value: Any) -> List[str]:
    text = clean_text(str(value or ""))
    return [part.strip() for part in re.split(r"[；;|,\n]+", text) if part.strip()]


def candidate_ids(detail_jsonl: Path, attachment_jsonl: Path, explicit_ids: Set[str], limit: int) -> List[str]:
    if explicit_ids:
        return list(explicit_ids)[:limit if limit > 0 else None]
    details = latest_rows_by_id(detail_jsonl)
    attachments = latest_rows_by_id(attachment_jsonl)
    result: List[str] = []
    for item_id, row in attachments.items():
        if row.get("附件正文抓取状态") != "失败":
            continue
        detail = details.get(item_id, {})
        names = split_joined(detail.get("附件名称") or row.get("附件名称"))
        if not any(name.lower().endswith((".pdf", ".docx")) for name in names):
            continue
        result.append(item_id)
        if limit > 0 and len(result) >= limit:
            break
    return result


def ocr_image(image: Image.Image, max_side: int = 2200) -> str:
    image = image.convert("RGB")
    if max(image.size) > max_side:
        scale = max_side / max(image.size)
        image = image.resize((max(1, int(image.width * scale)), max(1, int(image.height * scale))))
    fd, path = tempfile.mkstemp(suffix=".png")
    os.close(fd)
    try:
        image.save(path)
        result, _ = get_ocr()(path)
    finally:
        try:
            os.remove(path)
        except OSError:
            pass
    lines = []
    for item in result or []:
        if len(item) >= 2 and item[1]:
            lines.append(str(item[1]))
    return clean_text("\n".join(lines))


def ocr_pdf(path: Path, max_pages: int, scale: float) -> str:
    texts: List[str] = []
    with fitz.open(str(path)) as doc:
        for page in doc[:max_pages]:
            pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
            text = ocr_image(Image.open(io.BytesIO(pix.tobytes("png"))))
            if text:
                texts.append(text)
    return clean_text("\n".join(texts))


def ocr_docx_images(path: Path, max_images: int) -> str:
    texts: List[str] = []
    with zipfile.ZipFile(path) as archive:
        names = [
            name
            for name in archive.namelist()
            if name.startswith("word/media/") and name.lower().endswith((".png", ".jpg", ".jpeg"))
        ]
        for name in names[:max_images]:
            try:
                text = ocr_image(Image.open(io.BytesIO(archive.read(name))))
            except Exception:
                continue
            if text:
                texts.append(text)
    return clean_text("\n".join(texts))


def office_to_pdf(path: Path, output_dir: Path) -> Optional[Path]:
    suffix = path.suffix.lower()
    if suffix not in OFFICE_SUFFIXES:
        return None
    try:
        import pythoncom  # type: ignore
        import win32com.client  # type: ignore
    except Exception:
        return None

    pdf_path = output_dir / (path.stem + ".pdf")
    output_dir.mkdir(parents=True, exist_ok=True)
    if pdf_path.exists():
        try:
            pdf_path.unlink()
        except OSError:
            pass

    pythoncom.CoInitialize()
    app = None
    try:
        if suffix in {".doc", ".docx"}:
            app = win32com.client.DispatchEx("Word.Application")
            app.Visible = False
            app.DisplayAlerts = 0
            doc = app.Documents.Open(str(path), ReadOnly=True, AddToRecentFiles=False, Visible=False)
            try:
                doc.ExportAsFixedFormat(str(pdf_path), 17)
            finally:
                doc.Close(False)
        else:
            app = win32com.client.DispatchEx("Excel.Application")
            app.Visible = False
            app.DisplayAlerts = False
            wb = app.Workbooks.Open(str(path), ReadOnly=True)
            try:
                wb.ExportAsFixedFormat(0, str(pdf_path))
            finally:
                wb.Close(False)
    except Exception:
        return None
    finally:
        try:
            if app is not None:
                app.Quit()
        except Exception:
            pass
        pythoncom.CoUninitialize()

    return pdf_path if pdf_path.exists() else None


def ocr_office_via_pdf(path: Path, max_pdf_pages: int) -> str:
    with tempfile.TemporaryDirectory(prefix="ali_office_pdf_") as tmpdir:
        pdf_path = office_to_pdf(path.resolve(), Path(tmpdir).resolve())
        if not pdf_path:
            return ""
        return ocr_pdf(pdf_path, max_pages=max_pdf_pages, scale=2.0)


def local_files(cache_root: Path, item_id: str, max_mb: float) -> List[Path]:
    item_dir = cache_root / item_id
    if not item_dir.exists():
        return []
    files = []
    for path in item_dir.iterdir():
        if not path.is_file() or path.suffix.lower() not in {".pdf", ".doc", ".docx", ".xls", ".xlsx"}:
            continue
        if max_mb > 0 and path.stat().st_size > max_mb * 1024 * 1024:
            continue
        files.append(path)
    useful = ["评估", "估价", "调查", "报告", "房屋", "不动产", "调解", "执行"]
    files.sort(key=lambda p: (0 if any(k in p.name for k in useful) else 1, p.suffix.lower() != ".docx", p.name))
    return files


def process_one(item_id: str, detail: Dict[str, Any], cache_root: Path, max_mb: float, max_files: int, max_pdf_pages: int, max_docx_images: int) -> Dict[str, Any]:
    texts: List[str] = []
    parsed_files: List[Dict[str, Any]] = []
    errors: List[str] = []
    for path in local_files(cache_root, item_id, max_mb)[:max_files]:
        try:
            suffix = path.suffix.lower()
            if suffix == ".pdf":
                text = ocr_pdf(path, max_pages=max_pdf_pages, scale=2.0)
            elif suffix == ".docx":
                text = ocr_docx_images(path, max_images=max_docx_images)
                pdf_text = ocr_office_via_pdf(path, max_pdf_pages=max_pdf_pages)
                if len(pdf_text) > len(text):
                    text = pdf_text
            elif suffix in {".doc", ".xls", ".xlsx"}:
                text = ocr_office_via_pdf(path, max_pdf_pages=max_pdf_pages)
            else:
                text = ""
            parsed_files.append({"path": str(path), "size": path.stat().st_size, "ocr_text_length": len(text)})
            if text:
                texts.append(f"【{path.name}】\n{text}")
        except Exception as exc:  # noqa: BLE001
            errors.append(f"{path.name}:{exc}")
    combined = clean_text("\n".join(texts))
    title = clean_text(str(detail.get("标的物名称") or detail.get("详情标题") or detail.get("标题") or ""))
    fields = parse_attachment_fields(combined, title) if combined else {}
    merged = merge_prefer_existing(detail, fields, force_empty_only=True)
    merged.update(
        {
            "标的物ID": item_id,
            "附件正文OCR时间": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "附件正文OCR状态": "成功" if combined else "失败",
            "附件正文OCR错误": "；".join(errors),
            "附件正文OCR文件数": len(parsed_files),
            "附件正文OCR原文": json.dumps(parsed_files, ensure_ascii=False, separators=(",", ":")),
            "附件正文文本": combined[:12000] if combined else "",
            "附件正文解析字段": json.dumps(fields, ensure_ascii=False, separators=(",", ":")) if fields else "",
        }
    )
    return merged


def append_jsonl(path: Path, row: Dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("a", encoding="utf-8") as handle:
        handle.write(json.dumps(row, ensure_ascii=False, separators=(",", ":")) + "\n")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="OCR image-only cached Ali attachments.")
    parser.add_argument("--detail-jsonl", default=str(DEFAULT_DETAIL_JSONL))
    parser.add_argument("--attachment-jsonl", default=str(DEFAULT_ATTACHMENT_JSONL))
    parser.add_argument("--output-jsonl", default=str(DEFAULT_OUTPUT_JSONL))
    parser.add_argument("--cache-root", default=str(DEFAULT_CACHE_ROOT))
    parser.add_argument("--ids", default="")
    parser.add_argument("--limit", type=int, default=100)
    parser.add_argument("--workers", type=int, default=2)
    parser.add_argument("--max-mb", type=float, default=25)
    parser.add_argument("--max-files", type=int, default=1)
    parser.add_argument("--max-pdf-pages", type=int, default=2)
    parser.add_argument("--max-docx-images", type=int, default=6)
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    detail_jsonl = Path(args.detail_jsonl)
    attachment_jsonl = Path(args.attachment_jsonl)
    output_jsonl = Path(args.output_jsonl)
    cache_root = Path(args.cache_root)
    details = latest_rows_by_id(detail_jsonl)
    explicit_ids = {normalize_id(part) for part in args.ids.split(",") if normalize_id(part)}
    ids = candidate_ids(detail_jsonl, attachment_jsonl, explicit_ids, args.limit)
    print(f"ocr_plan candidates={len(ids)} workers={args.workers}", flush=True)
    done = 0
    success = 0
    with ThreadPoolExecutor(max_workers=max(1, args.workers)) as executor:
        future_map = {
            executor.submit(
                process_one,
                item_id,
                details.get(item_id, {"标的物ID": item_id}),
                cache_root,
                args.max_mb,
                args.max_files,
                args.max_pdf_pages,
                args.max_docx_images,
            ): item_id
            for item_id in ids
        }
        for future in as_completed(future_map):
            row = future.result()
            append_jsonl(output_jsonl, row)
            done += 1
            if row.get("附件正文OCR状态") == "成功":
                success += 1
            if done % 10 == 0 or done == len(ids):
                print(f"ocr_progress done={done}/{len(ids)} success={success}", flush=True)
    print(f"ocr_done done={done} success={success} output={output_jsonl}", flush=True)


if __name__ == "__main__":
    main()
