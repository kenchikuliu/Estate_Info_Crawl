import json
import math
import re
import subprocess
import zipfile
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple

import pandas as pd
import requests
from pypdf import PdfReader

GUANGDONG_CITY_NAMES = [
    "广州市", "深圳市", "珠海市", "汕头市", "佛山市", "韶关市", "湛江市", "肇庆市",
    "江门市", "茂名市", "惠州市", "梅州市", "汕尾市", "河源市", "阳江市", "清远市",
    "东莞市", "中山市", "潮州市", "揭阳市", "云浮市",
]

INVALID_FIELD_MARKERS = [
    "通知", "通知】", "承担交易环节", "历史欠缴税费", "竞买成功后", "办理过户手续",
    "权利类型", "所有权取得方式", "发生变更而产生", "费用由买受人", "详见附件",
]


SECTION_ALIASES = {
    "拍品名称": ["拍品名称", "标的物名称", "标的名称"],
    "权利来源": ["权利来源", "处置依据", "拍卖依据"],
    "权证情况": ["权证情况"],
    "拍品所有人": ["拍品所有人", "标的物所有人", "标的所有人", "权利人", "产权人"],
    "拍品现状": ["拍品现状", "标的物现状", "标的现状"],
    "租赁情况": ["租赁情况", "租赁或占有情况", "是否出租或承包", "租赁情况/占用情况"],
    "钥匙/占用情况": ["钥匙/占用情况", "钥 匙", "钥匙", "占用情况", "钥匙情况"],
    "户籍/工商注册": ["户籍/工商注册", "户口情况", "户籍情况", "工商注册"],
    "欠费情况": ["欠费情况"],
    "权利限制状况及抵押状况": ["权利限制状况及抵押状况", "权利限制情况", "抵押情况"],
    "成交后提供的文件": ["成交后提供的文件", "成交后提供 的文件", "提供的文件"],
    "拍品介绍": ["拍品介绍", "标的物介绍", "拍卖标的介绍"],
    "房屋权属状况": ["房屋权属状况"],
    "土地权属状况": ["土地权属状况"],
}

SURVEY_LABEL_MAPPING = {
    "标的名称": "标的物名称",
    "拍品名称": "标的物名称",
    "标的物名称": "标的物名称",
    "标的物名称、坐落": "标的物名称",
    "拍卖对象": "标的物名称",
    "产权证号": "权证情况",
    "权属证号": "权证情况",
    "权属证明号码": "权证情况",
    "权证情况": "权证情况",
    "不动产证号": "权证情况",
    "不动产权证书号": "权证情况",
    "房产证号": "权证情况",
    "房地产权证号": "权证情况",
    "标的所有人": "被执行人",
    "标的物所有人": "被执行人",
    "拍品所有人": "被执行人",
    "权利人": "被执行人",
    "产权人": "被执行人",
    "被执行人": "被执行人",
    "钥匙": "钥匙",
    "钥匙情况": "钥匙",
    "钥匙/占用情况": "钥匙",
    "占有使用或租赁情况": "钥匙",
    "占用情况": "钥匙",
    "租赁情况/占用情况": "钥匙",
    "使用情况": "钥匙",
    "是否已腾空": "腾空情况",
    "户籍/工商注册": "户籍注册",
    "户口情况": "户籍注册",
    "户籍情况": "户籍注册",
    "工商注册": "户籍注册",
    "欠费情况": "欠费情况",
    "欠费": "欠费情况",
    "提供的文件": "提供文件",
    "成交后提供的文件": "提供文件",
    "建筑面积": "建筑面积",
    "面积": "建筑面积",
    "房屋类型": "房屋类型",
    "性质": "房屋类型",
    "房屋性质": "房屋类型",
    "房屋用途": "房屋用途",
    "用途": "房屋用途",
    "规划用途": "房屋用途",
    "使用性质": "房屋用途",
    "总层数": "总层数",
    "所在楼层": "所在层",
    "所在层": "所在层",
    "楼层": "所在层",
    "竣工时间": "竣工时间",
    "竣工日期": "竣工时间",
    "建成时间": "竣工时间",
    "建成年代": "竣工时间",
    "购买时间": "购买时间",
    "购买日期": "购买时间",
    "购置时间": "购买时间",
    "取得时间": "购买时间",
    "取得日期": "购买时间",
    "询价反馈日期": "核准日期",
    "议价反馈日期": "核准日期",
    "核准日期": "核准日期",
    "所有权来源": "所有权来源",
    "所有权取得方式": "所有权来源",
    "转移方式": "所有权来源",
    "取得方式": "所有权来源",
    "土地用途": "土地用途",
    "土地性质": "土地性质",
    "权利性质": "土地性质",
    "使用期限": "使用期限",
    "宗地使用期限": "使用期限",
    "土地使用年限": "使用期限",
    "拍卖依据": "权利来源",
    "拍卖依据文号": "权利来源",
    "案号": "权利来源",
    "案 号": "权利来源",
    "权利限制情况": "权利限制状况及抵押状况",
    "权利限制状况及抵押状况": "权利限制状况及抵押状况",
    "抵押情况": "权利限制状况及抵押状况",
}


def clean_text(text: str) -> str:
    if not text:
        return ""
    text = text.replace("\r", "\n")
    text = re.sub(r"\u3000", " ", text)
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{2,}", "\n", text)
    return text.strip()


def _is_empty_like(text: str) -> bool:
    normalized = clean_text(text).lower()
    return normalized in {"", "nan", "none", "null", "n/a"}


def _normalize_label(label: str) -> str:
    normalized = clean_text(label)
    normalized = re.sub(r"[\s:：；;，,。\.、]+", "", normalized)
    return normalized


NORMALIZED_SURVEY_LABEL_MAPPING = {
    _normalize_label(label): field for label, field in SURVEY_LABEL_MAPPING.items()
}


def _label_to_field(label: str) -> Optional[str]:
    normalized = _normalize_label(label)
    if not normalized:
        return None
    return NORMALIZED_SURVEY_LABEL_MAPPING.get(normalized)


def _header_cell_to_field(label: str) -> Optional[str]:
    normalized = _normalize_label(label)
    if not normalized:
        return None
    if "产权证号" in normalized or "不动产权证" in normalized:
        return "权证情况"
    if "房屋用途" in normalized or normalized == "用途":
        return "房屋用途"
    if normalized in {"权利人", "产权人", "标的所有人"}:
        return "被执行人"
    if "座落" in normalized or "坐落" in normalized:
        return "标的物名称"
    if normalized == "结构":
        return "房屋类型"
    if "所在楼层" in normalized or "所在层" in normalized:
        return "所在层"
    if "建筑面积" in normalized:
        return "建筑面积"
    if "土地用途" in normalized:
        return "土地用途"
    if "权利性质" in normalized or "土地性质" in normalized:
        return "土地性质"
    return _label_to_field(label)


def _looks_like_data_row(cells: List[str]) -> bool:
    if not cells:
        return False
    first = _normalize_label(cells[0])
    if first in {"合计", "小计"}:
        return False
    return bool(re.fullmatch(r"[0-9一二三四五六七八九十]+", first))


def _parse_header_value_table(rows: List[List[str]]) -> Dict[str, str]:
    result: Dict[str, str] = {}
    for index, row in enumerate(rows[:-1]):
        headers = [clean_text(cell) for cell in row if clean_text(cell)]
        fields = [_header_cell_to_field(cell) for cell in headers]
        if len([field for field in fields if field]) < 3:
            continue

        for next_row in rows[index + 1 : index + 4]:
            values = [clean_text(cell) for cell in next_row if clean_text(cell)]
            if len(values) < 2 or not _looks_like_data_row(values):
                continue

            offset = 1 if len(values) == len(headers) + 1 and not fields[0] else 0
            for header_index, field in enumerate(fields):
                if not field:
                    continue
                value_index = header_index + offset
                if value_index >= len(values):
                    continue
                normalized = normalize_field_value(field, values[value_index])
                if not normalized or normalized in {"/", "无"}:
                    continue
                if field not in result:
                    result[field] = normalized
            break
    return result


def _alias_pairs() -> List[Tuple[str, str]]:
    pairs = []
    for canonical, aliases in SECTION_ALIASES.items():
        for alias in aliases:
            pairs.append((canonical, alias))
    pairs.sort(key=lambda item: len(item[1]), reverse=True)
    return pairs


ALIAS_PAIRS = _alias_pairs()


def parse_intro_sections(text: str) -> Dict[str, str]:
    normalized = clean_text(text)
    if not normalized:
        return {}

    matches = []
    for canonical, alias in ALIAS_PAIRS:
        for match in re.finditer(re.escape(alias), normalized):
            matches.append((match.start(), canonical, alias))

    matches.sort(key=lambda item: item[0])
    sections: Dict[str, str] = {}

    for idx, (start, canonical, alias) in enumerate(matches):
        if canonical in sections:
            continue
        content_start = start + len(alias)
        content_end = len(normalized)
        for next_start, _, _ in matches[idx + 1:]:
            if next_start > start:
                content_end = next_start
                break
        content = normalized[content_start:content_end].strip(" ：:\n")
        if content:
            sections[canonical] = content

    return sections


def extract_pdf_text(pdf_path: str) -> str:
    try:
        reader = PdfReader(pdf_path)
        texts = [page.extract_text() or "" for page in reader.pages]
        return clean_text("\n".join(texts))
    except Exception:
        return ""


def extract_pdf_fields(text: str) -> Dict[str, str]:
    normalized = clean_text(text)
    patterns = {
        "建筑面积": [r"建筑面积[:：]?\s*([^\n；;，,。]+)", r"面积[:：]?\s*([^\n；;，,。]+)"],
        "房屋类型": [r"房屋类型[:：]?\s*([^\n；;，,。]+)"],
        "房屋用途": [r"房屋用途[:：]?\s*([^\n；;，,。]+)", r"用途[:：]?\s*([^\n；;，,。]+)"],
        "总层数": [r"总层数[:：]?\s*([^\n；;，,。]+)"],
        "所在层": [r"所在层[:：]?\s*([^\n；;，,。]+)", r"位于第([^\n；;，,。]+)"],
        "竣工时间": [
            r"(?:竣工时间|竣工日期|建成时间|建成年代)[:：]?\s*([^\n；;，,。]+)",
        ],
        "购买时间": [
            r"(?:购买时间|购买日期|购置时间|取得时间|取得日期)[:：]?\s*([^\n；;，,。]+)",
        ],
        "土地性质": [r"土地性质[:：]?\s*([^\n；;，,。]+)"],
        "土地用途": [r"土地用途[:：]?\s*([^\n；;，,。]+)"],
        "使用期限": [r"使用期限[:：]?\s*([^\n；;。]+)", r"终止日期[:：]?\s*([^\n；;。]+)"],
        "是否占用": [r"占用情况[:：]?\s*([^\n；;。]+)", r"有人占用[^\n。]*"],
        "欠费": [r"欠费情况[:：]?\s*([^\n]+)"],
        "查封/抵押": [r"查封[^。\n]*", r"抵押[^。\n]*"],
    }
    result: Dict[str, str] = {}
    for field, regexes in patterns.items():
        for regex in regexes:
            match = re.search(regex, normalized)
            if match:
                result[field] = clean_text(match.group(0) if field == "查封/抵押" else match.group(1))
                break
    return result


def extract_docx_text(docx_path: str) -> str:
    try:
        with zipfile.ZipFile(docx_path) as archive:
            xml = archive.read("word/document.xml").decode("utf-8", errors="ignore")
        text = re.sub(r"</w:p>", "\n", xml)
        text = re.sub(r"<[^>]+>", "", text)
        return clean_text(text)
    except Exception:
        return ""


def extract_doc_text(doc_path: str) -> str:
    for command in (["antiword", doc_path], ["antiword.exe", doc_path]):
        try:
            result = subprocess.run(
                command,
                stdout=subprocess.PIPE,
                stderr=subprocess.DEVNULL,
                check=False,
                timeout=30,
            )
            for encoding in ("utf-8", "gbk", "latin1"):
                text = result.stdout.decode(encoding, errors="ignore")
                normalized = clean_text(text)
                if len(normalized) >= 20:
                    return normalized
        except Exception:
            continue
    try:
        data = Path(doc_path).read_bytes()
        text = data.decode("utf-8", errors="ignore")
        normalized = clean_text(text)
        if len(normalized) < 20:
            text = data.decode("gbk", errors="ignore")
            normalized = clean_text(text)
        return normalized
    except Exception:
        return ""


def extract_spreadsheet_text(file_path: str) -> str:
    try:
        excel = pd.ExcelFile(file_path)
        lines: List[str] = []
        for sheet_name in excel.sheet_names:
            df = pd.read_excel(file_path, sheet_name=sheet_name, header=None).fillna("")
            for row in df.values.tolist():
                cells = [clean_text(str(cell)) for cell in row if clean_text(str(cell))]
                if cells:
                    lines.append(" ".join(cells))
        return clean_text("\n".join(lines))
    except Exception:
        return ""


def extract_attachment_text(file_path: str) -> str:
    suffix = Path(file_path).suffix.lower()
    if suffix == ".pdf":
        return extract_pdf_text(file_path)
    if suffix == ".doc":
        return extract_doc_text(file_path)
    if suffix == ".docx":
        return extract_docx_text(file_path)
    if suffix in {".xls", ".xlsx"}:
        return extract_spreadsheet_text(file_path)
    if suffix in {".txt", ".csv"}:
        try:
            return clean_text(Path(file_path).read_text(encoding="utf-8", errors="ignore"))
        except Exception:
            return ""
    return ""


def extract_line_labeled_fields(text: str) -> Dict[str, str]:
    normalized = clean_text(text)
    if not normalized:
        return {}

    lines = [clean_text(line).strip(" ：:\t") for line in normalized.split("\n")]
    lines = [line for line in lines if line]
    result: Dict[str, str] = {}
    index = 0
    while index < len(lines):
        canonical = _label_to_field(lines[index])
        if not canonical:
            index += 1
            continue

        values: List[str] = []
        next_index = index + 1
        while next_index < len(lines):
            if _label_to_field(lines[next_index]):
                break
            values.append(lines[next_index])
            next_index += 1

        value = clean_text(" ".join(values))
        if value:
            if value in {"/", "无"} and result.get(canonical):
                index = max(next_index, index + 1)
                continue
            result[canonical] = value
        index = max(next_index, index + 1)

    return result


def extract_labeled_fields(text: str) -> Dict[str, str]:
    normalized = clean_text(text)
    if not normalized:
        return {}

    field_patterns = {
        "标的物名称": [
            r"(?:拍品名称|标的物名称|标的物名称、坐落|拍卖对象)[:：]?\s*([^\n]+)",
            r"位于\s*([\s\S]{1,180}?)(?=[（(]|房产证号|房地产权证号|，建筑面积|,建筑面积)",
        ],
        "权利来源": [
            r"权利来源[:：]?\s*([^\n]+)",
            r"处置依据[:：]?\s*([^\n]+)",
            r"拍卖依据[:：]?\s*([^\n]+)",
            r"拍卖依据文号[:：]?\s*([^\n]+)",
            r"案\s*号[:：]?\s*([^\n]+)",
            r"所有权来源[:：]?\s*([^\n]+)",
        ],
        "权证情况": [
            r"(?:房产证号|房地产权证号|权属证号|权证情况|产权证号|不动产证号|不动产权证书号|权属证明号码)[:：]?\s*([\s\S]{1,100}?号)",
            r"(?:权属证号|权证情况|产权证号|不动产证号|不动产权证书号)[:：]?\s*([^\n]+)",
        ],
        "被执行人": [
            r"(?:拍品所有人|标的所有人|标的物所有人|权利人|产权人)[:：]?\s*([^\n]+)",
            r"被执行人[:：]?\s*([^\n]+)",
        ],
        "钥匙": [
            r"(?:钥匙/占用情况|钥匙情况|钥匙)[:：]?\s*([^\n]+)",
            r"钥\s*匙[:：]?\s*([^\n]+)",
            r"(?:租赁情况/占用情况|占有使用或租赁情况|占用情况|占有使用情况)[:：]?\s*([^\n]+)",
        ],
        "户籍注册": [
            r"(?:户籍/工商注册|户口情况|户籍情况|工商注册)[:：]?\s*([^\n]+)",
        ],
        "欠费情况": [
            r"欠费情况[:：]?\s*([^\n]+)",
            r"欠费[:：]?\s*([^\n]+)",
            r"(截至[^\n]*?尚欠[^\n]+)",
        ],
        "提供文件": [
            r"(?:成交后提供的文件|提供的文件)[:：]?\s*([^\n]+)",
        ],
        "建筑面积": [
            r"建筑面积(?:为)?[:：]?\s*([0-9]+(?:\.[0-9]+)?\s*(?:平方米|平米|㎡))",
            r"建筑面积[:：]?\s*([^\n；;，,。]+)",
        ],
        "房屋类型": [
            r"房屋性质[:：]?\s*([^\n；;，,。]+)",
            r"性质[:：]?\s*([^\n；;，,。]+)",
            r"房屋类型[:：]?\s*([^\n；;，,。]+)",
            r"权利性质[:：]?\s*([^\n；;，,。]+)",
        ],
        "房屋用途": [
            r"房屋用途[:：]?\s*([^\n；;，,。]+)",
            r"用途[:：]?\s*([^\n；;，,。]+)",
            r"规划用途[:：]?\s*([^\n；;，,。]+)",
            r"使用性质[:：]?\s*([^\n；;，,。]+)",
        ],
        "总层数": [
            r"(?:总层数|共)\s*([0-9一二三四五六七八九十百]+层)",
        ],
        "竣工时间": [
            r"(?:竣工时间|竣工日期|建成时间|建成年代)[:：]?\s*([^\n；;，,。]+)",
            r"建成于\s*([0-9]{4}年(?:[0-9]{1,2}月)?)",
        ],
        "购买时间": [
            r"(?:购买时间|购买日期|购置时间|取得时间|取得日期)[:：]?\s*([^\n；;，,。]+)",
            r"(?:于|在)\s*([0-9]{4}年(?:[0-9]{1,2}月(?:[0-9]{1,2}日)?)?)\s*(?:购买|购置|取得)",
        ],
        "核准日期": [
            r"核准日期[:：]?\s*([^\n]+)",
        ],
        "所有权来源": [
            r"所有权来源[:：]?\s*([^\n]+)",
            r"所有权取得方式[:：]?\s*([^\n]+)",
            r"转移方式[:：]?\s*([^\n]+)",
            r"取得方式[:：]?\s*([^\n]+)",
        ],
        "土地用途": [
            r"土地用途[:：]?\s*([^\n；;，,。]+)",
        ],
        "土地性质": [
            r"(?:土地性质|权利性质)[:：]?\s*([^\n；;，,。]+)",
        ],
        "使用期限": [
            r"(?:使用期限|终止日期)[:：]?\s*([^\n；;。]+)",
            r"宗地使用期限[:：]?\s*([^\n；;。]+)",
            r"土地使用年限[:：]?\s*([^\n；;。]+)",
        ],
        "所在层": [
            r"(?:所在楼层|所在层|楼层)[:：]?\s*([^\n；;，,。]+)",
            r"([0-9一二三四五六七八九十]+层，共[0-9一二三四五六七八九十]+层[^\n]*)",
        ],
    }

    result: Dict[str, str] = {}
    for field, patterns in field_patterns.items():
        for pattern in patterns:
            match = re.search(pattern, normalized)
            if match:
                value = clean_text(match.group(1))
                if value:
                    result[field] = value
                    break

    result.update(extract_line_labeled_fields(normalized))
    return result


def normalize_field_value(field: str, value: str) -> str:
    cleaned = clean_text(value)
    if _is_empty_like(cleaned):
        return ""
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    cleaned = re.sub(r"\s+([㎡平方米层年月日号座栋单元室])", r"\1", cleaned)
    cleaned = re.sub(r"([第共])\s+", r"\1", cleaned)

    if field == "建筑面积":
        match = re.search(r"([0-9]+(?:\.[0-9]+)?\s*(?:平方米|平米|㎡))", cleaned)
        if match:
            return match.group(1)
    if field == "房屋用途":
        cleaned = cleaned.replace("房屋用途", "").replace("用途", "")
        cleaned = re.sub(r"^(?:及|为|用途为)\s*", "", cleaned)
        cleaned = re.sub(r"^性质\s*", "", cleaned)
        cleaned = cleaned.replace("规划", "").strip()
    if field == "房屋类型":
        cleaned = cleaned.replace("土地性质及用途", "").replace("房屋：", "").replace("土地：", "")
        cleaned = re.sub(r"^(?:为|属)\s*", "", cleaned)
        cleaned = cleaned.strip()
    if field == "土地性质":
        cleaned = cleaned.replace("土地性质及用途", "").replace("权利性质", "").replace("土地：", "")
        cleaned = re.sub(r"^(?:及用途|性质|为|属)\s*", "", cleaned).strip()
    if field == "土地用途":
        cleaned = cleaned.replace("土地用途", "").replace("用途", "")
        cleaned = cleaned.replace("房屋规划用途", "").replace("房屋用途", "")
        cleaned = re.sub(r"^(?:及|为|用途为)\s*", "", cleaned).strip()
    if field == "所在层":
        cleaned = cleaned.replace("位于第", "")
    if field in {"竣工时间", "购买时间", "核准日期"}:
        cleaned = cleaned.replace("建成年代", "").replace("建成时间", "")
        cleaned = cleaned.replace("竣工时间", "").replace("竣工日期", "")
        cleaned = cleaned.replace("购买时间", "").replace("购买日期", "")
        cleaned = cleaned.replace("购置时间", "").replace("取得时间", "").replace("取得日期", "")
        cleaned = re.sub(r"^(?:为|于|在)\s*", "", cleaned).strip(" ：:")
    if field == "权利来源":
        tokens = [token for token in re.split(r"\s+", cleaned) if token]
        deduped = []
        for token in tokens:
            if token not in deduped:
                deduped.append(token)
        cleaned = " ".join(deduped)
    if field in {"标的物名称", "权证情况"}:
        parts = [part for part in re.split(r"\s+", cleaned) if part]
        if len(parts) >= 2 and len(set(parts)) == 1:
            cleaned = parts[0]
    if field == "被执行人":
        cleaned = re.sub(r"[（(]被执行人.*", "", cleaned).strip()
    return cleaned


def sanitize_structured_value(field: str, value: str) -> str:
    cleaned = normalize_field_value(field, value)
    if not cleaned:
        return ""

    if any(marker in cleaned for marker in INVALID_FIELD_MARKERS):
        if field in {"被执行人", "房屋类型"}:
            return ""

    if field == "被执行人":
        if len(cleaned) > 80:
            return ""
        tokens = [token for token in re.split(r"\s+", cleaned) if token]
        deduped_tokens: List[str] = []
        seen = set()
        for token in tokens:
            if token not in seen:
                seen.add(token)
                deduped_tokens.append(token)
        cleaned = " ".join(deduped_tokens)
        if any(token in cleaned for token in ["拍卖", "税费", "法院", "竞买", "通知"]):
            return ""
    if field == "房屋类型":
        for pattern in [
            r"(市场化商品房)",
            r"(商品房)",
            r"(住宅)",
            r"(公寓)",
            r"(别墅)",
            r"(工业)",
            r"(商业)",
            r"(办公)",
            r"(厂房)",
        ]:
            match = re.search(pattern, cleaned)
            if match:
                return match.group(1)
        if len(cleaned) > 40:
            return ""
    if field == "房屋用途":
        for pattern in [
            r"(住宅)",
            r"(商业)",
            r"(办公)",
            r"(工业)",
            r"(车位)",
            r"(仓储)",
            r"(城镇住宅用地)",
        ]:
            match = re.search(pattern, cleaned)
            if match:
                return match.group(1)
        if len(cleaned) > 30:
            return ""
    if field == "土地性质":
        for pattern in [
            r"(出让)",
            r"(划拨)",
            r"(国有建设用地使用权)",
            r"(国有土地（出让）)",
            r"(国有土地使用权)",
        ]:
            match = re.search(pattern, cleaned)
            if match:
                return match.group(1)
        if len(cleaned) > 40:
            return ""
    if field == "土地用途":
        for pattern in [
            r"(城镇住宅用地)",
            r"(二类居住用地)",
            r"(住宅)",
            r"(商业、住宅)",
            r"(商业)",
            r"(办公)",
            r"(工业)",
        ]:
            match = re.search(pattern, cleaned)
            if match:
                return match.group(1)
        if len(cleaned) > 40:
            return ""
    if field == "权证情况" and len(cleaned) > 120:
        return ""
    if field in {"竣工时间", "购买时间", "核准日期"}:
        match = re.search(r"([0-9]{4}年(?:[0-9]{1,2}月(?:[0-9]{1,2}日)?)?|[0-9]{4}-[0-9]{1,2}(?:-[0-9]{1,2})?)", cleaned)
        if match:
            return match.group(1)
        if len(cleaned) > 40:
            return ""
    if field == "标的物名称":
        cleaned = re.sub(r"^、坐落\s*", "", cleaned)
    return cleaned


def postprocess_structured_fields(data: Dict[str, str]) -> Dict[str, str]:
    result = dict(data)
    for field in [
        "标的物名称", "权证情况", "被执行人", "钥匙", "户籍注册", "欠费情况", "提供文件",
        "建筑面积", "房屋类型", "房屋用途", "总层数", "所在层", "竣工时间", "购买时间", "土地性质", "土地用途",
        "权利来源", "所有权来源", "使用期限", "核准日期", "腾空情况",
    ]:
        if field in result:
            result[field] = sanitize_structured_value(field, str(result[field]))
    return result


def parse_survey_table_rows(rows: List[List[str]]) -> Dict[str, str]:
    result: Dict[str, str] = _parse_header_value_table(rows)
    for row in rows:
        cells = [clean_text(cell) for cell in row if clean_text(cell)]
        if len(cells) < 2:
            continue
        index = 0
        while index < len(cells):
            canonical = _label_to_field(cells[index])
            if not canonical:
                index += 1
                continue
            values: List[str] = []
            next_index = index + 1
            while next_index < len(cells) and not _label_to_field(cells[next_index]):
                values.append(cells[next_index])
                next_index += 1
            value = " ".join(values)
            normalized = normalize_field_value(canonical, value)
            if normalized:
                if normalized in {"/", "无"} and result.get(canonical):
                    index = max(next_index, index + 1)
                    continue
                if not result.get(canonical):
                    result[canonical] = normalized
            index = max(next_index, index + 1)

    return result


def extract_rights_status_text(text: str) -> Dict[str, str]:
    normalized = clean_text(text)
    result = {
        "权利限制状况及抵押状况": "",
        "房屋权属状况": "",
        "土地权属状况": "",
    }
    if not normalized:
        return result

    rights_match = re.search(
        r"(权利限制状况及\s*抵押状况|权利限制状况及抵押状况|权利限制情况|抵押情况)[:：]?\s*(.+?)(?=(成交后提供|拍品介绍|房屋权属状况|土地权属状况|$))",
        normalized,
        re.S,
    )
    if rights_match:
        result["权利限制状况及抵押状况"] = clean_text(rights_match.group(2))

    house_match = re.search(
        r"(?:房屋权属状况)[:：]?\s*(.+?)(?=(?:土地权属状况|2、土地权属状况|土地宗地号|$))",
        normalized,
        re.S,
    )
    if house_match:
        result["房屋权属状况"] = clean_text(house_match.group(1))

    land_match = re.search(
        r"(?:土地权属状况)[:：]?\s*(.+?)(?=(?:特别提示|竞买公告|竞买须知|$))",
        normalized,
        re.S,
    )
    if land_match:
        result["土地权属状况"] = clean_text(land_match.group(1))

    if not result["房屋权属状况"]:
        alt_house_match = re.search(
            r"(?:标的物名称、坐落|标的物情况|房产概况及拍卖标的|拍卖标的物情况)[:：]?\s*(.+?)(?=(?:特别提示|税、费|竞买人条件|土地使用期限|使用期限|$))",
            normalized,
            re.S,
        )
        if alt_house_match:
            result["房屋权属状况"] = clean_text(alt_house_match.group(1))

    if not result["土地权属状况"]:
        alt_land_match = re.search(
            r"(?:土地使用期限|土地用途|宗地（丘）面积|土地宗地号)[:：]?\s*(.+?)(?=(?:评估价|起拍价|特别提示|竞买人条件|$))",
            normalized,
            re.S,
        )
        if alt_land_match:
            result["土地权属状况"] = clean_text(alt_land_match.group(0))

    return result


def _extract_region_by_suffix(address: str, suffixes: List[str]) -> Optional[str]:
    for suffix in suffixes:
        index = address.find(suffix)
        if index > 0:
            start = index
            while start > 0 and address[start - 1] not in "省市区县镇乡街道盟旗 ":
                start -= 1
            candidate = address[start:index + len(suffix)]
            if len(candidate) >= len(suffix) + 1:
                return candidate
    return None


def extract_region_fields(address: str) -> Dict[str, str]:
    normalized = clean_text(address)
    result = {"省": "", "市": "", "区县": "", "城市": "", "格式化地址": normalized}
    if not normalized:
        return result

    if normalized.startswith("广东"):
        result["省"] = "广东"
        for city_name in GUANGDONG_CITY_NAMES:
            if city_name in normalized:
                result["市"] = city_name
                result["城市"] = city_name
                city_index = normalized.find(city_name)
                remaining = normalized[city_index + len(city_name):]
                district_match = re.match(r"([\u4e00-\u9fa5]{1,12}(?:区|县|市|旗|镇|街道))", remaining)
                if district_match:
                    district = district_match.group(1)
                    if district.endswith("街道") or district.endswith("镇"):
                        second_match = re.match(r"([\u4e00-\u9fa5]{1,12}(?:区|县|市|旗))", remaining)
                        if second_match:
                            district = second_match.group(1)
                    if city_name in {"东莞市", "中山市"} and (district == city_name or "名下位于" in district):
                        district = ""
                    if district.startswith(city_name) or "位于" in district:
                        district = ""
                    result["区县"] = district
                return result

    province_match = re.match(r"((?:北京|天津|上海|重庆)市|(?:内蒙古|广西|宁夏|新疆|西藏)自治区|(?:香港|澳门)特别行政区|[\u4e00-\u9fa5]{2,8}省)", normalized)
    if province_match:
        result["省"] = province_match.group(1)
        remaining = normalized[len(result["省"]):]
    else:
        short_province_match = re.match(r"([\u4e00-\u9fa5]{2,3})(?=[\u4e00-\u9fa5]{2,6}市)", normalized)
        remaining = normalized
        if short_province_match:
            result["省"] = short_province_match.group(1)
            remaining = normalized[len(result["省"]):]

    city_match = re.match(r"([\u4e00-\u9fa5]{2,12}市)", remaining)
    if city_match:
        result["市"] = city_match.group(1)
    elif result["省"] in {"北京市", "天津市", "上海市", "重庆市"}:
        result["市"] = result["省"]
    else:
        city = _extract_region_by_suffix(normalized, ["市", "自治州", "地区", "盟"])
        if city:
            result["市"] = city

    district = _extract_region_by_suffix(normalized, ["区", "县", "市", "旗"])
    if district and district != result["市"]:
        result["区县"] = district

    result["城市"] = result["市"] or result["城市"]
    return result


def filter_image_urls(urls: Iterable[str]) -> List[str]:
    cleaned = []
    seen = set()
    invalid_keywords = ["logo", "icon", "qrcode", "二维码", "tps-", "TB1", "png"]
    for url in urls:
        if not url:
            continue
        normalized = url.strip()
        if normalized.startswith("//"):
            normalized = "https:" + normalized
        lower = normalized.lower()
        if any(keyword.lower() in lower for keyword in invalid_keywords):
            continue
        if normalized in seen:
            continue
        seen.add(normalized)
        cleaned.append(normalized)
    return cleaned


def dump_json(value) -> str:
    return json.dumps(value, ensure_ascii=False)


def list_pdf_files(folder_path: str) -> List[str]:
    folder = Path(folder_path)
    if not folder.exists():
        return []
    return [str(path) for path in folder.glob("*.pdf")]


def list_attachment_files(folder_path: str) -> List[str]:
    folder = Path(folder_path)
    if not folder.exists():
        return []
    paths: List[str] = []
    for suffix in ("*.pdf", "*.doc", "*.docx", "*.xls", "*.xlsx", "*.txt", "*.csv"):
        paths.extend(str(path) for path in folder.glob(suffix))
    return paths


POI_COLUMN_SPECS = {
    "交通-地铁站": {
        "tags": [("railway", "station"), ("station", "subway"), ("subway", "yes")],
        "radius": 3000,
    },
    "交通-公交站": {
        "tags": [("highway", "bus_stop"), ("public_transport", "platform"), ("amenity", "bus_station")],
        "radius": 2000,
    },
    "教育-幼儿园": {
        "tags": [("amenity", "kindergarten")],
        "radius": 3000,
    },
    "教育-小学": {
        "tags": [("amenity", "school")],
        "radius": 3000,
        "name_keywords": ["小学"],
    },
    "教育-中学": {
        "tags": [("amenity", "school")],
        "radius": 4000,
        "name_keywords": ["中学", "高中", "初中"],
    },
    "购物-商场": {
        "tags": [("shop", "mall"), ("building", "retail"), ("landuse", "retail")],
        "radius": 5000,
    },
    "购物-超市": {
        "tags": [("shop", "supermarket"), ("shop", "convenience")],
        "radius": 3000,
    },
    "医疗-综合医院": {
        "tags": [("amenity", "hospital")],
        "radius": 5000,
    },
    "医疗-诊所": {
        "tags": [("amenity", "clinic"), ("healthcare", "clinic")],
        "radius": 3000,
    },
    "公园-公园": {
        "tags": [("leisure", "park"), ("boundary", "protected_area")],
        "radius": 5000,
    },
}


def haversine_distance_meters(lat1: float, lon1: float, lat2: float, lon2: float) -> float:
    radius = 6371000.0
    phi1 = math.radians(lat1)
    phi2 = math.radians(lat2)
    d_phi = math.radians(lat2 - lat1)
    d_lambda = math.radians(lon2 - lon1)
    a = math.sin(d_phi / 2) ** 2 + math.cos(phi1) * math.cos(phi2) * math.sin(d_lambda / 2) ** 2
    return 2 * radius * math.atan2(math.sqrt(a), math.sqrt(1 - a))


def format_distance(distance_meters: float) -> str:
    if distance_meters < 1000:
        return f"{int(round(distance_meters))}m"
    return f"{distance_meters / 1000:.2f}km"


def extract_coordinates(raw_value: str) -> Tuple[Optional[float], Optional[float]]:
    normalized = clean_text(raw_value)
    if not normalized:
        return None, None
    parts = re.split(r"[,，\s]+", normalized)
    numbers = []
    for part in parts:
        try:
            numbers.append(float(part))
        except Exception:
            continue
    if len(numbers) >= 2:
        return numbers[0], numbers[1]
    return None, None


def normalize_address_for_geocode(address: str) -> str:
    normalized = clean_text(address)
    if not normalized:
        return ""
    replacements = [
        (r".*?名下位于", ""),
        (r".*?位于", ""),
        (r"的不动产.*$", ""),
        (r"房产一处.*$", ""),
        (r"房产.*$", ""),
        (r"住宅.*$", ""),
        (r"[（(].*?[)）]", ""),
    ]
    candidate = normalized
    for pattern, repl in replacements:
        candidate = re.sub(pattern, repl, candidate)
    candidate = re.sub(r"^广东东莞市东莞市", "广东省东莞市", candidate)
    candidate = re.sub(r"^广东广州市", "广东省广州市", candidate)
    candidate = re.sub(r"^广东佛山市", "广东省佛山市", candidate)
    candidate = re.sub(r"^广东深圳市", "广东省深圳市", candidate)
    candidate = re.sub(r"^广东珠海市", "广东省珠海市", candidate)
    candidate = re.sub(r"\s+", "", candidate)
    return candidate or normalized


def geocode_address(address: str, timeout: int = 20) -> Dict[str, str]:
    normalized = normalize_address_for_geocode(address)
    if not normalized:
        return {}

    headers = {"User-Agent": "EstateInfoCrawl/1.0 (geocoding enrichment)"}
    try:
        response = requests.get(
            "https://geocode.arcgis.com/arcgis/rest/services/World/GeocodeServer/findAddressCandidates",
            params={"f": "json", "maxLocations": 1, "singleLine": normalized},
            headers=headers,
            timeout=timeout,
        )
        response.raise_for_status()
        items = response.json().get("candidates", [])
    except Exception:
        return {}

    if not items:
        return {}

    item = items[0]
    display_address = clean_text(item.get("address", ""))
    region_fields = extract_region_fields(display_address or normalized)
    location = item.get("location", {}) or {}
    return {
        "经度": str(location.get("x", "")),
        "纬度": str(location.get("y", "")),
        "格式化地址": display_address or normalized,
        "省": region_fields.get("省", ""),
        "市": region_fields.get("市", ""),
        "区县": region_fields.get("区县", ""),
    }


def build_overpass_query(lat: float, lon: float, radius: int, tags: List[Tuple[str, str]]) -> str:
    parts = []
    for key, value in tags:
        parts.append(f'node(around:{radius},{lat},{lon})["{key}"="{value}"];')
        parts.append(f'way(around:{radius},{lat},{lon})["{key}"="{value}"];')
        parts.append(f'relation(around:{radius},{lat},{lon})["{key}"="{value}"];')
    return "[out:json][timeout:25];(" + "".join(parts) + ");out center;"


def nearest_poi_distance(lat: float, lon: float, spec: Dict[str, object], timeout: int = 30) -> str:
    query = build_overpass_query(lat, lon, int(spec["radius"]), list(spec["tags"]))
    try:
        response = requests.post(
            "https://overpass-api.de/api/interpreter",
            data=query.encode("utf-8"),
            headers={"User-Agent": "EstateInfoCrawl/1.0 (poi enrichment)"},
            timeout=timeout,
        )
        response.raise_for_status()
        payload = response.json()
    except Exception:
        return ""

    name_keywords = [str(keyword) for keyword in spec.get("name_keywords", [])]
    distances: List[float] = []
    for element in payload.get("elements", []):
        poi_lat = element.get("lat", element.get("center", {}).get("lat"))
        poi_lon = element.get("lon", element.get("center", {}).get("lon"))
        if poi_lat is None or poi_lon is None:
            continue
        name = clean_text((element.get("tags", {}) or {}).get("name", ""))
        if name_keywords and name and not any(keyword in name for keyword in name_keywords):
            continue
        distance = haversine_distance_meters(lat, lon, float(poi_lat), float(poi_lon))
        distances.append(distance)

    if not distances:
        return ""
    return format_distance(min(distances))


def enrich_poi_distances(lat: float, lon: float, timeout: int = 45) -> Dict[str, str]:
    max_radius = max(int(spec["radius"]) for spec in POI_COLUMN_SPECS.values())
    tag_pairs = set()
    for spec in POI_COLUMN_SPECS.values():
        for tag in spec["tags"]:
            tag_pairs.add(tag)

    query = build_overpass_query(lat, lon, max_radius, list(tag_pairs))
    try:
        response = requests.post(
            "https://overpass-api.de/api/interpreter",
            data=query.encode("utf-8"),
            headers={"User-Agent": "EstateInfoCrawl/1.0 (poi enrichment)"},
            timeout=timeout,
        )
        response.raise_for_status()
        payload = response.json()
    except Exception:
        return {column: "" for column in POI_COLUMN_SPECS}

    results = {column: "" for column in POI_COLUMN_SPECS}
    min_distances: Dict[str, Optional[float]] = {column: None for column in POI_COLUMN_SPECS}

    for element in payload.get("elements", []):
        tags = element.get("tags", {}) or {}
        poi_lat = element.get("lat", element.get("center", {}).get("lat"))
        poi_lon = element.get("lon", element.get("center", {}).get("lon"))
        if poi_lat is None or poi_lon is None:
            continue
        distance = haversine_distance_meters(lat, lon, float(poi_lat), float(poi_lon))
        name = clean_text(tags.get("name", ""))

        for column, spec in POI_COLUMN_SPECS.items():
            if distance > float(spec["radius"]):
                continue
            matched = False
            for key, value in spec["tags"]:
                if tags.get(key) == value:
                    matched = True
                    break
            if not matched:
                continue
            name_keywords = [str(keyword) for keyword in spec.get("name_keywords", [])]
            if name_keywords and name and not any(keyword in name for keyword in name_keywords):
                continue
            current = min_distances[column]
            if current is None or distance < current:
                min_distances[column] = distance

    for column, distance in min_distances.items():
        if distance is not None:
            results[column] = format_distance(distance)
    return results
