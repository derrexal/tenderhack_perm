from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime
from decimal import Decimal, ROUND_HALF_UP
from io import BytesIO
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple
import re
import threading
import zipfile
import xml.etree.ElementTree as ET

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Mm, Pt
from fastapi import FastAPI, HTTPException
from fastapi.responses import Response
from pydantic import BaseModel, Field

DATA_DIR = Path(__file__).parent / "raw_data"
CONTRACTS_FILE = DATA_DIR / "TenderHack_Контракты_20260313.xlsx"
STE_FILE = DATA_DIR / "TenderHack_СТЕ_20260313.xlsx"

NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
SAMPLE_LIMIT = 5
DEFAULT_CURRENCY = "RUB"


class Ste(BaseModel):
    id: Optional[int] = None
    name: Optional[str] = None
    category: Optional[str] = None
    manufacturer: Optional[str] = None
    characteristics: Optional[str] = None


class Contract(BaseModel):
    id: Optional[int] = None
    procurementName: Optional[str] = None
    procurementMethod: Optional[str] = None
    initialContractValue: Optional[str] = None
    contractValueAfterSigning: Optional[str] = None
    reductionPercent: Optional[str] = None
    vatRate: Optional[str] = None
    contractSigningDate: Optional[str] = None
    buyerInn: Optional[str] = None
    buyerRegion: Optional[str] = None
    supplierInn: Optional[str] = None
    supplierRegion: Optional[str] = None


class ContractItemWithSte(BaseModel):
    id: Optional[int] = None
    steId: Optional[int] = None
    steItemName: Optional[str] = None
    quantity: Optional[str] = None
    unit: Optional[str] = None
    unitPrice: Optional[str] = None
    steName: Optional[str] = None
    steCategory: Optional[str] = None
    steManufacturer: Optional[str] = None
    steCharacteristics: Optional[str] = None


class JustificationRequest(BaseModel):
    contract: Contract
    items: List[ContractItemWithSte]
    currency: str = Field(default=DEFAULT_CURRENCY, description="Код валюты, по умолчанию RUB")
    maxSamples: int = Field(default=3, description="Сколько примеров контрактов показывать в документе")


class ContractReport(BaseModel):
    contract: Contract
    items: List[ContractItemWithSte]


class BatchJustificationRequest(BaseModel):
    contracts: List[ContractReport]
    currency: str = Field(default=DEFAULT_CURRENCY, description="Код валюты, по умолчанию RUB")
    maxSamples: int = Field(default=3, description="Сколько примеров контрактов показывать в документе")
    reportTitle: Optional[str] = Field(
        default=None, description="Заголовок отчета. Если не задан, используется стандартный."
    )


class StePriceJustificationRow(BaseModel):
    contractId: Optional[str] = None
    procurementMethod: Optional[str] = None
    initialContractValue: Optional[str] = None
    contractValueAfterSigning: Optional[str] = None
    reductionPercent: Optional[str] = None
    contractSigningDate: Optional[str] = None
    buyerInn: Optional[str] = None
    supplierInn: Optional[str] = None
    steId: Optional[int] = None
    steItemName: Optional[str] = None
    unitPrice: Optional[str] = None


class StePriceJustificationRequest(BaseModel):
    items: List[StePriceJustificationRow]
    currency: str = Field(default=DEFAULT_CURRENCY, description="Код валюты, по умолчанию RUB")
    reportTitle: Optional[str] = Field(
        default=None, description="Заголовок отчета. Если не задан, используется стандартный."
    )
    steId: Optional[int] = Field(default=None, description="Идентификатор СТЕ (если нужно зафиксировать)")
    steItemName: Optional[str] = Field(default=None, description="Наименование СТЕ (если нужно зафиксировать)")
    signerName: Optional[str] = Field(default=None, description="ФИО подписанта")
    signerTitle: Optional[str] = Field(default=None, description="Должность подписанта")


@dataclass
class PriceSample:
    contract_id: str
    signed_at: datetime
    unit_price: Decimal


@dataclass
class PriceStats:
    count: int = 0
    sum_price: Decimal = Decimal("0")
    min_price: Optional[Decimal] = None
    max_price: Optional[Decimal] = None
    min_date: Optional[datetime] = None
    max_date: Optional[datetime] = None
    samples: List[PriceSample] = None

    def __post_init__(self) -> None:
        if self.samples is None:
            self.samples = []

    def update(self, price: Decimal, signed_at: Optional[datetime], contract_id: Optional[str]) -> None:
        self.count += 1
        self.sum_price += price
        if self.min_price is None or price < self.min_price:
            self.min_price = price
        if self.max_price is None or price > self.max_price:
            self.max_price = price
        if signed_at:
            if self.min_date is None or signed_at < self.min_date:
                self.min_date = signed_at
            if self.max_date is None or signed_at > self.max_date:
                self.max_date = signed_at
        if signed_at and contract_id:
            self.samples.append(PriceSample(contract_id=contract_id, signed_at=signed_at, unit_price=price))
            self.samples.sort(key=lambda s: s.signed_at, reverse=True)
            if len(self.samples) > SAMPLE_LIMIT:
                self.samples = self.samples[:SAMPLE_LIMIT]

    @property
    def avg_price(self) -> Optional[Decimal]:
        if not self.count:
            return None
        return self.sum_price / Decimal(self.count)


@dataclass
class ContractReportData:
    contract: Contract
    items: List[ContractItemWithSte]
    analysis: List[Dict[str, object]]
    total_value: Decimal


class RawDataRepository:
    def __init__(self, contracts_path: Path, ste_path: Path) -> None:
        self.contracts_path = contracts_path
        self.ste_path = ste_path
        self._loaded = False
        self._lock = threading.Lock()
        self.ste_by_id: Dict[str, Ste] = {}
        self.stats_by_ste_id: Dict[str, PriceStats] = {}
        self.stats_by_ste_id_method: Dict[Tuple[str, str], PriceStats] = {}
        self.stats_by_name: Dict[str, PriceStats] = {}
        self.stats_by_category: Dict[str, PriceStats] = {}

    def ensure_loaded(self) -> None:
        if self._loaded:
            return
        with self._lock:
            if self._loaded:
                return
            if not self.contracts_path.exists() or not self.ste_path.exists():
                raise FileNotFoundError("raw_data файлы не найдены")
            self.ste_by_id = load_ste_catalog(self.ste_path)
            (
                self.stats_by_ste_id,
                self.stats_by_ste_id_method,
                self.stats_by_name,
                self.stats_by_category,
            ) = load_contract_stats(self.contracts_path, self.ste_by_id)
            self._loaded = True


DATA_REPO = RawDataRepository(CONTRACTS_FILE, STE_FILE)


def normalize_text(value: Optional[str]) -> str:
    if not value:
        return ""
    text = value.lower().strip()
    text = re.sub(r"\s+", " ", text)
    return text


def normalize_name(value: Optional[str]) -> str:
    if not value:
        return ""
    text = value.lower()
    text = re.sub(r"[^0-9a-zа-яё\s]", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def parse_decimal(value: Optional[str]) -> Optional[Decimal]:
    if value is None:
        return None
    if isinstance(value, Decimal):
        return value
    if isinstance(value, (int, float)):
        return Decimal(str(value))
    text = str(value).strip()
    if not text:
        return None
    text = text.replace(" ", "").replace(",", ".")
    text = re.sub(r"[^0-9.\-]", "", text)
    if not text or text in {"-", "."}:
        return None
    try:
        return Decimal(text)
    except Exception:
        return None


def parse_datetime(value: Optional[str]) -> Optional[datetime]:
    if not value:
        return None
    if isinstance(value, datetime):
        return value
    text = str(value).strip()
    if not text:
        return None
    for fmt in ("%Y-%m-%d %H:%M:%S.%f", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d"):
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            continue
    try:
        return datetime.fromisoformat(text)
    except ValueError:
        return None


def format_decimal(value: Optional[Decimal], places: int = 2) -> str:
    if value is None:
        return "—"
    quant = Decimal("1") if places == 0 else Decimal("0." + "0" * (places - 1) + "1")
    amount = value.quantize(quant, rounding=ROUND_HALF_UP)
    formatted = f"{amount:,.{places}f}".replace(",", " ")
    return formatted.replace(".", ",")


def format_optional_decimal(value: Optional[str], places: int = 2) -> str:
    if value is None:
        return ""
    parsed = parse_decimal(value)
    if parsed is None:
        return str(value)
    return format_decimal(parsed, places)


def format_optional_date(value: Optional[str]) -> str:
    if not value:
        return ""
    parsed = parse_datetime(value)
    if parsed is None:
        return str(value)
    return parsed.strftime("%d.%m.%Y")


def currency_label(code: str) -> str:
    code_up = (code or "").upper()
    if code_up in {"RUB", "RUR", "РУБ", "RUBLE"}:
        return "руб."
    return code


def rtf_escape(text: str) -> str:
    escaped = []
    for ch in text:
        if ch in {"\\", "{", "}"}:
            escaped.append("\\" + ch)
        elif ch == "\n":
            escaped.append("\\par ")
        elif ord(ch) < 128:
            escaped.append(ch)
        else:
            escaped.append(f"\\u{ord(ch)}?")
    return "".join(escaped)


def _ensure_rfonts(style, font_name: str) -> None:
    r_pr = style._element.get_or_add_rPr()
    r_fonts = r_pr.find(qn("w:rFonts"))
    if r_fonts is None:
        r_fonts = OxmlElement("w:rFonts")
        r_pr.append(r_fonts)
    r_fonts.set(qn("w:ascii"), font_name)
    r_fonts.set(qn("w:hAnsi"), font_name)
    r_fonts.set(qn("w:eastAsia"), font_name)
    r_fonts.set(qn("w:cs"), font_name)


def _ensure_paragraph_style(
    styles,
    name: str,
    font_name: str,
    font_size: int,
    bold: bool = False,
    alignment: Optional[int] = None,
    line_spacing: Optional[float] = None,
    first_line_indent_cm: Optional[float] = None,
    space_before_pt: float = 0,
    space_after_pt: float = 0,
    based_on: Optional[str] = None,
) -> None:
    if name in styles:
        style = styles[name]
    else:
        style = styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
    if based_on:
        style.base_style = styles[based_on]
    font = style.font
    font.name = font_name
    font.size = Pt(font_size)
    font.bold = bold
    _ensure_rfonts(style, font_name)

    pf = style.paragraph_format
    if alignment is not None:
        pf.alignment = alignment
    if line_spacing is not None:
        pf.line_spacing = line_spacing
    if first_line_indent_cm is not None:
        pf.first_line_indent = Cm(first_line_indent_cm)
    pf.space_before = Pt(space_before_pt)
    pf.space_after = Pt(space_after_pt)


def _ensure_table_style(styles, name: str, base: str = "Table Grid") -> None:
    if name in styles:
        return
    style = styles.add_style(name, WD_STYLE_TYPE.TABLE)
    style.base_style = styles[base]
    style.hidden = False
    style.quick_style = True
    style.priority = 99

    tbl_style_pr = OxmlElement("w:tblStylePr")
    tbl_style_pr.set(qn("w:type"), "firstRow")

    r_pr = OxmlElement("w:rPr")
    b = OxmlElement("w:b")
    r_pr.append(b)
    tbl_style_pr.append(r_pr)

    tc_pr = OxmlElement("w:tcPr")
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), "E6E6E6")
    tc_pr.append(shd)
    tbl_style_pr.append(tc_pr)

    style._element.append(tbl_style_pr)


def _set_repeat_table_header(row) -> None:
    tr = row._tr
    tr_pr = tr.get_or_add_trPr()
    header = OxmlElement("w:tblHeader")
    header.set(qn("w:val"), "true")
    tr_pr.append(header)


class RtfBuilder:
    def __init__(self, landscape: bool = False, margin_twips: int = 1134) -> None:
        if landscape:
            paperw, paperh = 16840, 11907
        else:
            paperw, paperh = 11907, 16840
        self.parts: List[str] = [
            "{\\rtf1\\ansi\\deff0"
            "{\\fonttbl{\\f0 Calibri;}}"
            "{\\colortbl ;\\red230\\green230\\blue230;}"
            f"\\paperw{paperw}\\paperh{paperh}"
            f"\\margl{margin_twips}\\margr{margin_twips}"
            f"\\margt{margin_twips}\\margb{margin_twips}"
            "\\viewkind4\\uc1\\pard\\fs22 "
        ]

    def add_paragraph(
        self,
        text: str,
        bold: bool = False,
        size: Optional[int] = None,
        align: str = "left",
        space_after: int = 120,
    ) -> None:
        prefix = ""
        suffix = ""
        align_code = {"left": "\\ql", "center": "\\qc", "right": "\\qr", "justify": "\\qj"}.get(
            align, "\\ql"
        )
        if bold:
            prefix += "\\b "
            suffix += "\\b0 "
        if size is not None:
            prefix += f"\\fs{size} "
            suffix += "\\fs22 "
        spacing = f"\\sa{space_after} "
        self.parts.append(
            f"{align_code}{spacing}" + prefix + rtf_escape(text) + suffix + "\\par "
        )

    def finish(self) -> bytes:
        self.parts.append("}")
        return "".join(self.parts).encode("ascii", errors="ignore")

    def add_page_break(self) -> None:
        self.parts.append("\\page ")

    def add_table(
        self,
        rows: List[List[str]],
        col_widths: List[int],
        header_rows: int = 1,
        header_shading: bool = True,
        borders: bool = True,
        font_size: int = 20,
    ) -> None:
        if not rows or not col_widths:
            return
        col_positions = []
        acc = 0
        for width in col_widths:
            acc += width
            col_positions.append(acc)

        total_cols = len(col_widths)
        border_defs = ""
        if borders:
            border_defs = "\\clbrdrt\\brdrs\\clbrdrl\\brdrs\\clbrdrb\\brdrs\\clbrdrr\\brdrs"

        for row_idx, row in enumerate(rows):
            padded_row = row + [""] * max(0, total_cols - len(row))
            row_rtf = ["\\trowd\\trgaph108\\trleft0"]
            if row_idx < header_rows:
                row_rtf.append("\\trhdr")
            for pos in col_positions:
                cell_defs = border_defs
                if header_shading and row_idx < header_rows:
                    cell_defs += "\\clcbpat1"
                row_rtf.append(f"{cell_defs}\\cellx{pos}")
            row_rtf.append(" ")
            for cell in padded_row[:total_cols]:
                cell_text = rtf_escape(cell)
                if row_idx < header_rows:
                    row_rtf.append(
                        f"\\intbl\\qc\\b\\fs{font_size} {cell_text}\\b0\\fs22\\cell "
                    )
                else:
                    row_rtf.append(
                        f"\\intbl\\ql\\fs{font_size} {cell_text}\\fs22\\cell "
                    )
            row_rtf.append("\\row ")
            self.parts.append("".join(row_rtf))


def col_to_idx(col: str) -> int:
    idx = 0
    for ch in col:
        idx = idx * 26 + (ord(ch) - 64)
    return idx - 1


def load_shared_strings(z: zipfile.ZipFile) -> List[str]:
    if "xl/sharedStrings.xml" not in z.namelist():
        return []
    strings: List[str] = []
    with z.open("xl/sharedStrings.xml") as f:
        for _, elem in ET.iterparse(f, events=("end",)):
            if elem.tag == f"{NS}si":
                text = "".join(
                    t.text or "" for t in elem.findall(f".//{NS}t")
                )
                strings.append(text)
                elem.clear()
    return strings


def iter_xlsx_rows(path: Path) -> Iterable[List[str]]:
    with zipfile.ZipFile(path) as z:
        shared = load_shared_strings(z)
        with z.open("xl/worksheets/sheet1.xml") as f:
            context = ET.iterparse(f, events=("end",))
            for _, elem in context:
                if elem.tag != f"{NS}row":
                    continue
                row_cells: Dict[int, str] = {}
                for cell in elem.findall(f"{NS}c"):
                    ref = cell.get("r")
                    if not ref:
                        continue
                    col = "".join(ch for ch in ref if ch.isalpha())
                    col_idx = col_to_idx(col)
                    cell_type = cell.get("t")
                    value = ""
                    if cell_type == "inlineStr":
                        is_elem = cell.find(f"{NS}is")
                        if is_elem is not None:
                            value = "".join(
                                t.text or "" for t in is_elem.findall(f".//{NS}t")
                            )
                    else:
                        v = cell.find(f"{NS}v")
                        if v is not None and v.text is not None:
                            value = v.text
                            if cell_type == "s":
                                try:
                                    value = shared[int(value)]
                                except Exception:
                                    pass
                    row_cells[col_idx] = value
                if row_cells:
                    max_idx = max(row_cells)
                    row = [""] * (max_idx + 1)
                    for idx, val in row_cells.items():
                        row[idx] = val
                else:
                    row = []
                yield row
                elem.clear()


def safe_get(row: List[str], idx: Optional[int]) -> str:
    if idx is None:
        return ""
    if idx >= len(row):
        return ""
    return row[idx]


def load_ste_catalog(path: Path) -> Dict[str, Ste]:
    rows = iter_xlsx_rows(path)
    header = next(rows, [])
    index = {name: idx for idx, name in enumerate(header)}
    ste_map: Dict[str, Ste] = {}
    for row in rows:
        ste_id = safe_get(row, index.get("Идентификатор СТЕ"))
        if not ste_id:
            continue
        ste = Ste(
            id=int(ste_id) if str(ste_id).isdigit() else None,
            name=safe_get(row, index.get("Наименование СТЕ")) or None,
            category=safe_get(row, index.get("Категория")) or None,
            manufacturer=safe_get(row, index.get("Производитель")) or None,
            characteristics=safe_get(row, index.get("характеристики СТЕ")) or None,
        )
        ste_map[str(ste_id)] = ste
    return ste_map


def load_contract_stats(
    path: Path,
    ste_map: Dict[str, Ste],
) -> Tuple[Dict[str, PriceStats], Dict[Tuple[str, str], PriceStats], Dict[str, PriceStats], Dict[str, PriceStats]]:
    rows = iter_xlsx_rows(path)
    header = next(rows, [])
    index = {name: idx for idx, name in enumerate(header)}
    stats_by_ste: Dict[str, PriceStats] = {}
    stats_by_ste_method: Dict[Tuple[str, str], PriceStats] = {}
    stats_by_name: Dict[str, PriceStats] = {}
    stats_by_category: Dict[str, PriceStats] = {}

    for row in rows:
        ste_id = safe_get(row, index.get("Идентификатор СТЕ по контракту"))
        ste_name = safe_get(row, index.get("Наименование позиции СТЕ"))
        unit_price_raw = safe_get(row, index.get("Цена за единицу"))
        if not unit_price_raw:
            continue
        unit_price = parse_decimal(unit_price_raw)
        if unit_price is None:
            continue
        contract_id = safe_get(row, index.get("Идентификатор контракта"))
        signed_at = parse_datetime(safe_get(row, index.get("Дата заключения контракта")))
        procurement_method = safe_get(row, index.get("Способ закупки"))

        if ste_id:
            ste_key = str(ste_id)
            stats_by_ste.setdefault(ste_key, PriceStats()).update(
                unit_price, signed_at, contract_id
            )
            if procurement_method:
                method_key = (ste_key, procurement_method)
                stats_by_ste_method.setdefault(method_key, PriceStats()).update(
                    unit_price, signed_at, contract_id
                )
            ste_info = ste_map.get(ste_key)
            if ste_info and ste_info.category:
                cat_key = normalize_text(ste_info.category)
                if cat_key:
                    stats_by_category.setdefault(cat_key, PriceStats()).update(
                        unit_price, signed_at, contract_id
                    )

        if ste_name:
            name_key = normalize_name(ste_name)
            if name_key:
                stats_by_name.setdefault(name_key, PriceStats()).update(
                    unit_price, signed_at, contract_id
                )

    return stats_by_ste, stats_by_ste_method, stats_by_name, stats_by_category


def enrich_item(item: ContractItemWithSte, ste_map: Dict[str, Ste]) -> ContractItemWithSte:
    if item.steId is None:
        return item
    ste = ste_map.get(str(item.steId))
    if not ste:
        return item
    if not item.steName:
        item.steName = ste.name
    if not item.steCategory:
        item.steCategory = ste.category
    if not item.steManufacturer:
        item.steManufacturer = ste.manufacturer
    if not item.steCharacteristics:
        item.steCharacteristics = ste.characteristics
    return item


def choose_stats(
    item: ContractItemWithSte,
    repo: RawDataRepository,
    procurement_method: Optional[str],
) -> Tuple[Optional[PriceStats], str]:
    if item.steId is not None:
        ste_key = str(item.steId)
        if procurement_method:
            stats = repo.stats_by_ste_id_method.get((ste_key, procurement_method))
            if stats and stats.count:
                return stats, "СТЕ + способ закупки"
        stats = repo.stats_by_ste_id.get(ste_key)
        if stats and stats.count:
            return stats, "СТЕ"
    name_key = normalize_name(item.steItemName or item.steName)
    if name_key:
        stats = repo.stats_by_name.get(name_key)
        if stats and stats.count:
            return stats, "наименование позиции"
    cat_key = normalize_text(item.steCategory)
    if cat_key:
        stats = repo.stats_by_category.get(cat_key)
        if stats and stats.count:
            return stats, "категория"
    return None, "нет данных"


def analyze_contract(
    contract: Contract,
    items: List[ContractItemWithSte],
    repo: RawDataRepository,
    max_samples: int,
) -> ContractReportData:
    enriched_items: List[ContractItemWithSte] = [
        enrich_item(item, repo.ste_by_id) for item in items
    ]

    analysis: List[Dict[str, object]] = []
    total_value = Decimal("0")
    sample_limit = max(0, min(max_samples, SAMPLE_LIMIT))

    for item in enriched_items:
        stats, basis = choose_stats(item, repo, contract.procurementMethod)
        quantity = parse_decimal(item.quantity) or Decimal("0")
        unit = item.unit or ""
        fallback_price = parse_decimal(item.unitPrice) or Decimal("0")
        if stats and stats.avg_price is not None:
            unit_price = stats.avg_price
        else:
            unit_price = fallback_price
        cost = quantity * unit_price
        total_value += cost

        analysis.append(
            {
                "name": item.steItemName or item.steName or "Без наименования",
                "quantity": quantity,
                "unit": unit,
                "unit_price": unit_price,
                "cost": cost,
                "stats": stats,
                "basis": basis,
                "fallback": format_decimal(fallback_price),
                "max_samples": sample_limit,
            }
        )

    return ContractReportData(
        contract=contract,
        items=enriched_items,
        analysis=analysis,
        total_value=total_value,
    )


def append_procurement_section(builder: RtfBuilder, contract: Contract, title: str) -> None:
    builder.add_paragraph(title, bold=True)
    if contract.procurementName:
        builder.add_paragraph(f"Наименование закупки: {contract.procurementName}")
    if contract.procurementMethod:
        builder.add_paragraph(f"Способ закупки: {contract.procurementMethod}")
    if contract.id is not None:
        builder.add_paragraph(f"Идентификатор контракта: {contract.id}")
    if contract.buyerInn or contract.buyerRegion:
        buyer = " ".join(part for part in [contract.buyerInn, contract.buyerRegion] if part)
        builder.add_paragraph(f"Заказчик: {buyer}")
    if contract.supplierInn or contract.supplierRegion:
        supplier = " ".join(
            part for part in [contract.supplierInn, contract.supplierRegion] if part
        )
        builder.add_paragraph(f"Поставщик (при наличии): {supplier}")
    if contract.contractSigningDate:
        builder.add_paragraph(f"Дата заключения контракта: {contract.contractSigningDate}")
    if contract.vatRate:
        builder.add_paragraph(f"Ставка НДС: {contract.vatRate}")


def append_items_section(builder: RtfBuilder, items: List[ContractItemWithSte], title: str) -> None:
    builder.add_paragraph(title, bold=True)
    for idx, item in enumerate(items, start=1):
        name = item.steItemName or item.steName or "Без наименования"
        qty = parse_decimal(item.quantity) or Decimal("0")
        unit = item.unit or ""
        ste_id = item.steId if item.steId is not None else "—"
        builder.add_paragraph(f"{idx}. {name}")
        builder.add_paragraph(f"Количество: {format_decimal(qty, 3)} {unit}. СТЕ: {ste_id}.")
        if item.steCategory or item.steManufacturer:
            cat = item.steCategory or "—"
            mfr = item.steManufacturer or "—"
            builder.add_paragraph(f"Категория: {cat}. Производитель: {mfr}.")
        if item.steCharacteristics:
            builder.add_paragraph(f"Характеристики: {item.steCharacteristics}")


def append_market_section(
    builder: RtfBuilder,
    analysis: List[Dict[str, object]],
    currency: str,
    title: str,
) -> None:
    builder.add_paragraph(title, bold=True)
    for idx, item_info in enumerate(analysis, start=1):
        builder.add_paragraph(f"{idx}. {item_info['name']}")
        basis = item_info["basis"]
        stats: Optional[PriceStats] = item_info["stats"]
        if stats and stats.count:
            period = ""
            if stats.min_date and stats.max_date:
                period = (
                    f" (период {stats.min_date.strftime('%d.%m.%Y')}"
                    f" — {stats.max_date.strftime('%d.%m.%Y')})"
                )
            builder.add_paragraph(
                f"Найдено сопоставимых контрактов: {stats.count}{period}. Основание: {basis}."
            )
            builder.add_paragraph(
                "Минимальная цена: "
                f"{format_decimal(stats.min_price)} {currency_label(currency)}; "
                "Средняя цена: "
                f"{format_decimal(stats.avg_price)} {currency_label(currency)}; "
                "Максимальная цена: "
                f"{format_decimal(stats.max_price)} {currency_label(currency)}."
            )
            samples = stats.samples[: item_info["max_samples"]]
            if samples:
                sample_lines = []
                for sample in samples:
                    sample_lines.append(
                        f"№{sample.contract_id} от {sample.signed_at.strftime('%d.%m.%Y')}: "
                        f"{format_decimal(sample.unit_price)} {currency_label(currency)}"
                    )
                builder.add_paragraph("Примеры контрактов: " + "; ".join(sample_lines))
        else:
            fallback = item_info["fallback"]
            builder.add_paragraph(
                "Сопоставимые контракты не найдены. "
                f"В качестве ориентира использована цена: {fallback} {currency_label(currency)}."
            )


def append_calc_section(
    builder: RtfBuilder,
    analysis: List[Dict[str, object]],
    total_value: Decimal,
    currency: str,
    title: str,
    total_label: str,
    total_bold: bool = True,
) -> None:
    builder.add_paragraph(title, bold=True)
    for idx, item_info in enumerate(analysis, start=1):
        qty = item_info["quantity"]
        unit = item_info["unit"]
        unit_price = item_info["unit_price"]
        cost = item_info["cost"]
        builder.add_paragraph(
            f"{idx}. {item_info['name']}: {format_decimal(qty, 3)} {unit} × "
            f"{format_decimal(unit_price)} = {format_decimal(cost)} {currency_label(currency)}"
        )
    builder.add_paragraph(
        f"{total_label}{format_decimal(total_value)} {currency_label(currency)}",
        bold=total_bold,
    )


def append_conclusion(builder: RtfBuilder, title: str) -> None:
    builder.add_paragraph(title, bold=True)
    builder.add_paragraph(
        "На основании анализа данных о ранее заключенных контрактах и справочника СТЕ "
        "рекомендуется установить начальную (максимальную) цену контракта на уровне, "
        "указанном в разделе расчета."
    )


def build_document(
    contract: Contract,
    items: List[ContractItemWithSte],
    analysis: List[Dict[str, object]],
    total_value: Decimal,
    currency: str,
) -> bytes:
    builder = RtfBuilder()
    today = date.today().strftime("%d.%m.%Y")

    builder.add_paragraph(
        "Обоснование начальной (максимальной) цены контракта для котировочной сессии",
        bold=True,
        size=32,
    )
    builder.add_paragraph(f"Дата формирования: {today}")

    append_procurement_section(builder, contract, "1. Сведения о закупке")
    append_items_section(builder, items, "2. Объект закупки")
    append_market_section(builder, analysis, currency, "3. Анализ рынка")
    append_calc_section(
        builder,
        analysis,
        total_value,
        currency,
        "4. Расчет НМЦК",
        "Итого НМЦК: ",
    )
    append_conclusion(builder, "5. Вывод")

    return builder.finish()


def build_batch_document(
    reports: List[ContractReportData],
    currency: str,
    report_title: Optional[str],
) -> bytes:
    builder = RtfBuilder()
    today = date.today().strftime("%d.%m.%Y")
    title = report_title or "Сводное обоснование начальной (максимальной) цены контрактов"

    builder.add_paragraph(title, bold=True, size=32)
    builder.add_paragraph(f"Дата формирования: {today}")
    builder.add_paragraph(f"Количество контрактов: {len(reports)}")

    total_all = sum((report.total_value for report in reports), Decimal("0"))
    builder.add_paragraph(
        f"Суммарная НМЦК: {format_decimal(total_all)} {currency_label(currency)}",
        bold=True,
    )

    builder.add_paragraph("1. Сводка по контрактам", bold=True)
    for idx, report in enumerate(reports, start=1):
        contract = report.contract
        name = contract.procurementName or "Без наименования"
        method = contract.procurementMethod or "—"
        items_count = len(report.items)
        builder.add_paragraph(f"{idx}. {name}")
        builder.add_paragraph(
            f"Способ закупки: {method}. Позиции: {items_count}. "
            f"НМЦК: {format_decimal(report.total_value)} {currency_label(currency)}."
        )

    builder.add_paragraph("2. Детализация по контрактам", bold=True)
    for idx, report in enumerate(reports, start=1):
        name = report.contract.procurementName or "Без наименования"
        builder.add_paragraph(f"Контракт {idx}: {name}", bold=True, size=28)
        append_procurement_section(builder, report.contract, "Сведения о закупке")
        append_items_section(builder, report.items, "Объект закупки")
        append_market_section(builder, report.analysis, currency, "Анализ рынка")
        append_calc_section(
            builder,
            report.analysis,
            report.total_value,
            currency,
            "Расчет НМЦК",
            "Итого НМЦК по контракту: ",
        )
        append_conclusion(builder, "Вывод")
        if idx != len(reports):
            builder.add_page_break()

    return builder.finish()


def build_ste_price_document(
    request: StePriceJustificationRequest,
    rows: List[StePriceJustificationRow],
) -> bytes:
    builder = RtfBuilder(landscape=True, margin_twips=720)
    today = date.today().strftime("%d.%m.%Y")
    title = request.reportTitle or "Обоснование начальной цены стандартной товарной единицы"

    ste_id_values = {row.steId for row in rows if row.steId is not None}
    ste_name_values = {row.steItemName for row in rows if row.steItemName}
    ste_id = request.steId or (next(iter(ste_id_values)) if ste_id_values else None)
    ste_name = request.steItemName or (next(iter(ste_name_values)) if ste_name_values else None)

    prices = []
    dates = []
    for row in rows:
        price = parse_decimal(row.unitPrice)
        if price is not None:
            prices.append(price)
        dt = parse_datetime(row.contractSigningDate)
        if dt:
            dates.append(dt)

    count = len(prices)
    avg_price = sum(prices, Decimal("0")) / Decimal(count) if count else None
    min_price = min(prices) if prices else None
    max_price = max(prices) if prices else None
    period = ""
    if dates:
        period = f"{min(dates).strftime('%d.%m.%Y')} — {max(dates).strftime('%d.%m.%Y')}"

    builder.add_paragraph(title, bold=True, size=32, align="center", space_after=240)
    builder.add_paragraph(f"Дата формирования: {today}", align="center", space_after=240)
    info_lines = []
    if ste_id is not None or ste_name:
        ste_line = "СТЕ: "
        if ste_id is not None:
            ste_line += str(ste_id)
        if ste_name:
            ste_line += f" — {ste_name}"
        info_lines.append(ste_line)
    if period:
        info_lines.append(f"Период данных: {period}")
    info_lines.append(f"Количество сопоставимых контрактов: {len(rows)}")
    for line in info_lines:
        builder.add_paragraph(line, align="left", space_after=120)

    if avg_price is not None:
        builder.add_paragraph(
            "Рекомендованная средняя цена: "
            f"{format_decimal(avg_price)} {currency_label(request.currency)}",
            bold=True,
            align="left",
            space_after=240,
        )

    if len(ste_id_values) > 1:
        builder.add_paragraph(
            "Внимание: входные данные содержат несколько идентификаторов СТЕ.",
            bold=True,
            space_after=120,
        )
    if len(ste_name_values) > 1:
        builder.add_paragraph(
            "Внимание: входные данные содержат несколько наименований СТЕ.",
            bold=True,
            space_after=120,
        )

    table_rows: List[List[str]] = [
        [
            "№",
            "ID контракта",
            "Способ закупки",
            "НМЦК",
            "Цена после заключения",
            "% снижения",
            "Дата заключения",
            "ИНН заказчика",
            "ИНН поставщика",
            "СТЕ",
            "Позиция СТЕ",
            "Цена за единицу",
        ]
    ]

    for idx, row in enumerate(rows, start=1):
        table_rows.append(
            [
                str(idx),
                row.contractId or "",
                row.procurementMethod or "",
                format_optional_decimal(row.initialContractValue),
                format_optional_decimal(row.contractValueAfterSigning),
                row.reductionPercent or "",
                format_optional_date(row.contractSigningDate),
                row.buyerInn or "",
                row.supplierInn or "",
                str(row.steId) if row.steId is not None else "",
                row.steItemName or "",
                format_optional_decimal(row.unitPrice),
            ]
        )

    col_widths = [400, 900, 1500, 1000, 1100, 800, 1100, 1100, 1100, 800, 2000, 1000]
    builder.add_table(
        table_rows,
        col_widths,
        header_rows=1,
        header_shading=True,
        borders=True,
        font_size=18,
    )

    if avg_price is not None:
        builder.add_paragraph(
            "Средняя цена: "
            f"{format_decimal(avg_price)} {currency_label(request.currency)}; "
            "Минимальная цена: "
            f"{format_decimal(min_price)} {currency_label(request.currency)}; "
            "Максимальная цена: "
            f"{format_decimal(max_price)} {currency_label(request.currency)}.",
            space_after=240,
        )

    signer_title = request.signerTitle or "Должность"
    signer_name = request.signerName or "ФИО"
    builder.add_paragraph("", space_after=120)
    builder.add_paragraph(
        f"{signer_title}: ____________________ /{signer_name}/",
        align="right",
        space_after=0,
    )

    return builder.finish()


def build_ste_price_docx(
    request: StePriceJustificationRequest,
    rows: List[StePriceJustificationRow],
) -> bytes:
    today = date.today().strftime("%d.%m.%Y")
    title = request.reportTitle or "Обоснование начальной цены стандартной товарной единицы"

    ste_id_values = {row.steId for row in rows if row.steId is not None}
    ste_name_values = {row.steItemName for row in rows if row.steItemName}
    ste_id = request.steId or (next(iter(ste_id_values)) if ste_id_values else None)
    ste_name = request.steItemName or (next(iter(ste_name_values)) if ste_name_values else None)

    prices = []
    dates = []
    for row in rows:
        price = parse_decimal(row.unitPrice)
        if price is not None:
            prices.append(price)
        dt = parse_datetime(row.contractSigningDate)
        if dt:
            dates.append(dt)

    count = len(prices)
    avg_price = sum(prices, Decimal("0")) / Decimal(count) if count else None
    min_price = min(prices) if prices else None
    max_price = max(prices) if prices else None
    period = ""
    if dates:
        period = f"{min(dates).strftime('%d.%m.%Y')} — {max(dates).strftime('%d.%m.%Y')}"

    doc = Document()
    section = doc.sections[0]
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width = Mm(297)
    section.page_height = Mm(210)
    section.left_margin = Mm(30)
    section.right_margin = Mm(10)
    section.top_margin = Mm(20)
    section.bottom_margin = Mm(20)

    styles = doc.styles
    _ensure_paragraph_style(
        styles,
        "Основной текст",
        font_name="Times New Roman",
        font_size=14,
        alignment=WD_ALIGN_PARAGRAPH.JUSTIFY,
        line_spacing=1.5,
        first_line_indent_cm=1.25,
        space_before_pt=0,
        space_after_pt=0,
    )
    _ensure_paragraph_style(
        styles,
        "Заголовок документа",
        font_name="Times New Roman",
        font_size=16,
        bold=True,
        alignment=WD_ALIGN_PARAGRAPH.CENTER,
        line_spacing=1.0,
        first_line_indent_cm=0,
        space_before_pt=0,
        space_after_pt=12,
        based_on="Основной текст",
    )
    _ensure_paragraph_style(
        styles,
        "Подзаголовок",
        font_name="Times New Roman",
        font_size=14,
        bold=True,
        alignment=WD_ALIGN_PARAGRAPH.LEFT,
        line_spacing=1.0,
        first_line_indent_cm=0,
        space_before_pt=6,
        space_after_pt=6,
        based_on="Основной текст",
    )
    _ensure_paragraph_style(
        styles,
        "Реквизиты",
        font_name="Times New Roman",
        font_size=12,
        alignment=WD_ALIGN_PARAGRAPH.RIGHT,
        line_spacing=1.0,
        first_line_indent_cm=0,
        space_before_pt=0,
        space_after_pt=12,
        based_on="Основной текст",
    )
    _ensure_paragraph_style(
        styles,
        "Список",
        font_name="Times New Roman",
        font_size=14,
        alignment=WD_ALIGN_PARAGRAPH.LEFT,
        line_spacing=1.5,
        first_line_indent_cm=0,
        space_before_pt=0,
        space_after_pt=0,
        based_on="Основной текст",
    )
    _ensure_paragraph_style(
        styles,
        "Подпись",
        font_name="Times New Roman",
        font_size=14,
        alignment=WD_ALIGN_PARAGRAPH.RIGHT,
        line_spacing=1.0,
        first_line_indent_cm=0,
        space_before_pt=12,
        space_after_pt=0,
        based_on="Основной текст",
    )
    _ensure_paragraph_style(
        styles,
        "Приложение",
        font_name="Times New Roman",
        font_size=12,
        alignment=WD_ALIGN_PARAGRAPH.LEFT,
        line_spacing=1.0,
        first_line_indent_cm=0,
        space_before_pt=6,
        space_after_pt=6,
        based_on="Основной текст",
    )
    _ensure_paragraph_style(
        styles,
        "Табличный текст",
        font_name="Times New Roman",
        font_size=11,
        alignment=WD_ALIGN_PARAGRAPH.LEFT,
        line_spacing=1.0,
        first_line_indent_cm=0,
        space_before_pt=0,
        space_after_pt=0,
        based_on="Основной текст",
    )
    _ensure_table_style(styles, "Таблица")

    doc.add_paragraph(title, style="Заголовок документа")
    doc.add_paragraph(f"Дата формирования: {today}", style="Реквизиты")

    if ste_id is not None or ste_name:
        ste_line = "СТЕ: "
        if ste_id is not None:
            ste_line += str(ste_id)
        if ste_name:
            ste_line += f" — {ste_name}"
        doc.add_paragraph(ste_line, style="Основной текст")
    if period:
        doc.add_paragraph(f"Период данных: {period}", style="Основной текст")
    doc.add_paragraph(
        f"Количество сопоставимых контрактов: {len(rows)}", style="Основной текст"
    )

    if avg_price is not None:
        doc.add_paragraph(
            "Рекомендованная средняя цена: "
            f"{format_decimal(avg_price)} {currency_label(request.currency)}",
            style="Подзаголовок",
        )

    if len(ste_id_values) > 1:
        doc.add_paragraph(
            "Внимание: входные данные содержат несколько идентификаторов СТЕ.",
            style="Основной текст",
        )
    if len(ste_name_values) > 1:
        doc.add_paragraph(
            "Внимание: входные данные содержат несколько наименований СТЕ.",
            style="Основной текст",
        )

    headers = [
        "№",
        "ID контракта",
        "Способ закупки",
        "НМЦК",
        "Цена после заключения",
        "% снижения",
        "Дата заключения",
        "ИНН заказчика",
        "ИНН поставщика",
        "СТЕ",
        "Позиция СТЕ",
        "Цена за единицу",
    ]

    table = doc.add_table(rows=1, cols=len(headers))
    table.style = "Таблица"
    table.autofit = True
    header_cells = table.rows[0].cells
    for idx, text in enumerate(headers):
        header_cells[idx].text = text
        for paragraph in header_cells[idx].paragraphs:
            paragraph.style = "Табличный текст"
    _set_repeat_table_header(table.rows[0])

    for idx, row in enumerate(rows, start=1):
        cells = table.add_row().cells
        values = [
            str(idx),
            row.contractId or "",
            row.procurementMethod or "",
            format_optional_decimal(row.initialContractValue),
            format_optional_decimal(row.contractValueAfterSigning),
            row.reductionPercent or "",
            format_optional_date(row.contractSigningDate),
            row.buyerInn or "",
            row.supplierInn or "",
            str(row.steId) if row.steId is not None else "",
            row.steItemName or "",
            format_optional_decimal(row.unitPrice),
        ]
        for cell, value in zip(cells, values):
            cell.text = value
            for paragraph in cell.paragraphs:
                paragraph.style = "Табличный текст"

    if avg_price is not None:
        doc.add_paragraph(
            "Средняя цена: "
            f"{format_decimal(avg_price)} {currency_label(request.currency)}; "
            "Минимальная цена: "
            f"{format_decimal(min_price)} {currency_label(request.currency)}; "
            "Максимальная цена: "
            f"{format_decimal(max_price)} {currency_label(request.currency)}.",
            style="Основной текст",
        )

    signer_title = request.signerTitle or "Должность"
    signer_name = request.signerName or "ФИО"
    doc.add_paragraph(
        f"{signer_title}: ____________________ /{signer_name}/",
        style="Подпись",
    )

    buf = BytesIO()
    doc.save(buf)
    return buf.getvalue()


app = FastAPI(title="TenderHack Price Justification")


@app.post("/api/v1/ste-price-justification/doc")
async def generate_ste_price_justification(payload: StePriceJustificationRequest) -> Response:
    if not payload.items:
        raise HTTPException(status_code=400, detail="items не должен быть пустым")

    doc_content = build_ste_price_docx(payload, payload.items)
    filename = "ste_price_justification.docx"
    headers = {"Content-Disposition": f"attachment; filename={filename}"}
    return Response(
        content=doc_content,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers=headers,
    )


if __name__ == "__main__":
    import uvicorn

    uvicorn.run(app, host="0.0.0.0", port=8000)
