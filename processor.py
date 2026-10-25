from __future__ import annotations

import io
import re
import unicodedata
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any
from zoneinfo import ZoneInfo

import pandas as pd
import xlsxwriter


BANGKOK_TZ = ZoneInfo("Asia/Bangkok")
SUPPLIER_ALIASES = {
    "White Warehouse",
    "Supplier-White Warehouse",
}


class ProcessingError(RuntimeError):
    pass


@dataclass
class OutputResult:
    content: bytes
    source_rows: int
    filtered_rows: int
    document_count: int
    matched_rows: int
    unmatched_items: list[str]
    branch_codes: list[str]
    document_numbers: list[str]


def _text(value: Any) -> str:
    if value is None or pd.isna(value):
        return ""
    value = str(value).strip()
    if value.endswith(".0") and value[:-2].isdigit():
        value = value[:-2]
    return value


def _identifier(value: Any, remove_leading_zeroes: bool = False) -> str:
    value = _text(value)
    if remove_leading_zeroes and value.isdigit():
        return value.lstrip("0") or "0"
    return value


def _key(value: Any) -> str:
    return unicodedata.normalize("NFKC", _text(value)).casefold()


def _excel_engine(filename: str) -> str:
    suffix = Path(filename).suffix.lower()
    if suffix == ".xls":
        return "xlrd"
    if suffix == ".xlsx":
        return "openpyxl"
    raise ProcessingError("รองรับเฉพาะไฟล์ .xls และ .xlsx")


def _read_excel_bytes(content: bytes, filename: str) -> pd.DataFrame:
    try:
        return pd.read_excel(
            io.BytesIO(content),
            header=None,
            dtype=object,
            engine=_excel_engine(filename),
        )
    except Exception as exc:
        raise ProcessingError(f"ไม่สามารถอ่านไฟล์ AMS ได้: {exc}") from exc


def _find_header_row(frame: pd.DataFrame, required: set[str]) -> int:
    for row_index in range(min(25, len(frame))):
        values = {_text(value) for value in frame.iloc[row_index].tolist()}
        if required.issubset(values):
            return row_index
    raise ProcessingError(
        "ไม่พบ Header ของ AMS: ลำดับ, รหัสทรัพย์สิน และชื่อ Supplier"
    )


def _make_unique(names: list[str]) -> list[str]:
    counts: dict[str, int] = {}
    result: list[str] = []
    for index, name in enumerate(names, start=1):
        base = name or f"column_{index}"
        counts[base] = counts.get(base, 0) + 1
        result.append(base if counts[base] == 1 else f"{base}_{counts[base]}")
    return result


def _build_headers(frame: pd.DataFrame, header_row: int) -> tuple[list[str], int]:
    top = [_text(value) for value in frame.iloc[header_row].tolist()]
    sub = (
        [_text(value) for value in frame.iloc[header_row + 1].tolist()]
        if header_row + 1 < len(frame)
        else [""] * len(top)
    )

    has_second_header = any(
        value in {"รหัสสาขา", "ชื่อสาขา", "ประเภทร้าน", "เลขที่เอกสาร"}
        for value in sub
    )
    if not has_second_header:
        return _make_unique(top), header_row + 1

    group_names = {"ร้านต้นทาง", "ร้านปลายทาง"}
    current_group = ""
    headers: list[str] = []

    for top_value, sub_value in zip(top, sub):
        if top_value:
            current_group = top_value if top_value in group_names else ""

        if current_group and sub_value:
            headers.append(f"{current_group}_{sub_value}")
        elif top_value and top_value not in group_names:
            headers.append(top_value)
        elif sub_value:
            headers.append(sub_value)
        else:
            headers.append("")

    return _make_unique(headers), header_row + 2


def _resolve_column(columns: list[str], aliases: list[str], required: bool = True) -> str:
    normalized = {_key(column): column for column in columns}
    for alias in aliases:
        if _key(alias) in normalized:
            return normalized[_key(alias)]
    if required:
        raise ProcessingError(f"ไม่พบคอลัมน์ที่ต้องใช้: {', '.join(aliases)}")
    return ""


def parse_ams(content: bytes, filename: str) -> pd.DataFrame:
    raw = _read_excel_bytes(content, filename)
    header_row = _find_header_row(
        raw,
        {"ลำดับ", "รหัสทรัพย์สิน", "ชื่อ Supplier"},
    )
    headers, data_start = _build_headers(raw, header_row)
    data = raw.iloc[data_start:].copy()
    data.columns = headers
    data = data.dropna(how="all").reset_index(drop=True)

    columns = list(data.columns)
    mapping = {
        "source_sequence": _resolve_column(columns, ["ลำดับ"]),
        "transaction_date": _resolve_column(
            columns, ["วันที่ดำเนินการ"], required=False
        ),
        "branch_code": _resolve_column(
            columns, ["ร้านต้นทาง_รหัสสาขา", "รหัสสาขา"]
        ),
        "branch_name": _resolve_column(
            columns, ["ร้านต้นทาง_ชื่อสาขา", "ชื่อสาขา"]
        ),
        "region": _resolve_column(
            columns, ["ร้านต้นทาง_ภาค", "ภาค"], required=False
        ),
        "asset_code": _resolve_column(columns, ["รหัสทรัพย์สิน"]),
        "sno": _resolve_column(columns, ["Sno.", "Sno"], required=False),
        "item_name": _resolve_column(columns, ["รายการทรัพย์สิน"]),
        "supplier_name": _resolve_column(columns, ["ชื่อ Supplier"]),
        "document_no": _resolve_column(
            columns, ["ร้านปลายทาง_เลขที่เอกสาร", "เลขที่เอกสาร"]
        ),
        "source_note": _resolve_column(columns, ["หมายเหตุ"], required=False),
    }

    output = pd.DataFrame(index=data.index)
    for canonical_name, source_name in mapping.items():
        output[canonical_name] = data[source_name] if source_name else ""
    output["_source_order"] = range(len(output))

    for column in output.columns:
        if column != "_source_order":
            output[column] = output[column].map(_text)

    output = output[
        (output["asset_code"] != "") | (output["item_name"] != "")
    ].reset_index(drop=True)
    return output


def load_accessory_master(master_path: Path) -> dict[str, dict[str, str]]:
    if not master_path.exists():
        raise ProcessingError(f"ไม่พบไฟล์ Master: {master_path.name}")

    try:
        sheets = pd.read_excel(
            master_path,
            sheet_name=None,
            header=None,
            dtype=object,
            engine=_excel_engine(master_path.name),
        )
    except Exception as exc:
        raise ProcessingError(f"ไม่สามารถอ่านไฟล์ Master ได้: {exc}") from exc

    preferred_names = ["Lookup", "รับเข้าคลัง"]
    selected: pd.DataFrame | None = None
    for sheet_name in preferred_names:
        if sheet_name in sheets:
            selected = sheets[sheet_name]
            break
    if selected is None:
        selected = next(iter(sheets.values()))

    header_row = -1
    for row_index in range(min(30, len(selected))):
        values = [_text(value) for value in selected.iloc[row_index].tolist()]
        if "รายการทรัพย์สิน" in values:
            header_row = row_index
            break
    if header_row < 0:
        raise ProcessingError("Master ไม่มีคอลัมน์ 'รายการทรัพย์สิน'")

    headers = _make_unique(
        [_text(value) for value in selected.iloc[header_row].tolist()]
    )
    table = selected.iloc[header_row + 1 :].copy()
    table.columns = headers

    item_col = _resolve_column(list(table.columns), ["รายการทรัพย์สิน"])
    accessory_col = _resolve_column(
        list(table.columns),
        ["อุปกรณ์ส่วนควบ", "จำนวนอุปกรณ์ส่วนควบ"],
        required=False,
    )
    note_col = _resolve_column(list(table.columns), ["หมายเหตุ"], required=False)
    image_col = _resolve_column(list(table.columns), ["รูปภาพ"], required=False)

    lookup: dict[str, dict[str, str]] = {}
    for _, row in table.iterrows():
        item_name = _text(row[item_col])
        if not item_name:
            continue
        key = _key(item_name)
        record = {
            "item_name": item_name,
            "accessory": _text(row[accessory_col]) if accessory_col else "",
            "note": _text(row[note_col]) if note_col else "",
            "image": _text(row[image_col]) if image_col else "",
        }
        if key in lookup and lookup[key] != record:
            raise ProcessingError(
                f"Master มีชื่ออุปกรณ์ซ้ำแต่ข้อมูลไม่ตรงกัน: {item_name}"
            )
        lookup[key] = record
    return lookup


def prepare_data(
    ams_data: pd.DataFrame,
    accessory_lookup: dict[str, dict[str, str]],
) -> tuple[pd.DataFrame, list[str]]:
    supplier_keys = {_key(value) for value in SUPPLIER_ALIASES}
    filtered = ams_data[
        ams_data["supplier_name"].map(_key).isin(supplier_keys)
    ].copy()
    if filtered.empty:
        found = sorted(
            value for value in ams_data["supplier_name"].map(_text).unique() if value
        )
        sample = ", ".join(found[:8]) or "ไม่พบค่า Supplier"
        raise ProcessingError(
            "ไม่พบข้อมูล Supplier-White Warehouse ในไฟล์ AMS "
            f"(ค่าที่พบ: {sample})"
        )

    matched: list[bool] = []
    accessories: list[str] = []
    lookup_notes: list[str] = []
    lookup_images: list[str] = []
    unmatched: list[str] = []

    for item_name in filtered["item_name"].tolist():
        record = accessory_lookup.get(_key(item_name))
        matched.append(record is not None)
        if record is None:
            accessories.append("")
            lookup_notes.append("")
            lookup_images.append("")
            if item_name and item_name not in unmatched:
                unmatched.append(item_name)
        else:
            accessories.append(record["accessory"])
            lookup_notes.append(record["note"])
            lookup_images.append(record["image"])

    filtered["_matched"] = matched
    filtered["_accessory"] = accessories
    filtered["_lookup_note"] = lookup_notes
    filtered["_lookup_image"] = lookup_images
    filtered = filtered.sort_values("_source_order", kind="stable").reset_index(drop=True)
    return filtered, unmatched


def _unique_values(values: pd.Series) -> list[str]:
    result: list[str] = []
    for value in values.map(_text).tolist():
        if value and value not in result:
            result.append(value)
    return result


def _sheet_name(document_no: str, used: set[str]) -> str:
    suffix = re.sub(r"[^0-9A-Za-zก-๙_-]+", "", document_no.replace("/", ""))
    suffix = suffix[-18:] or "NoDoc"
    base = f"Checklist_{suffix}"[:31]
    candidate = base
    number = 2
    while candidate.casefold() in used:
        ending = f"_{number}"
        candidate = f"{base[:31-len(ending)]}{ending}"
        number += 1
    used.add(candidate.casefold())
    return candidate


def _branch_name(value: Any) -> str:
    value = _text(value)
    return re.sub(r"^SE\s+", "", value, flags=re.IGNORECASE).strip()


def _row_note(row: pd.Series) -> str:
    return _text(row["source_note"]) or _text(row["_lookup_note"])


def _add_checklist_sheet(
    workbook: xlsxwriter.Workbook,
    sheet_name: str,
    rows: pd.DataFrame,
    document_no: str,
    master_dir: Path,
    created_at: datetime,
) -> None:
    worksheet = workbook.add_worksheet(sheet_name)
    worksheet.hide_gridlines(2)
    worksheet.set_landscape()
    worksheet.set_paper(9)
    worksheet.fit_to_pages(1, 0)
    worksheet.set_margins(0.2, 0.2, 0.3, 0.3)

    widths = [8.4, 15.1, 11.8, 6.9, 11.0, 46.0, 72.0, 24.0, 19.0, 19.6]
    for column, width in enumerate(widths):
        worksheet.set_column(column, column, width)

    base = {
        "font_name": "Arial",
        "font_size": 10,
        "border": 1,
        "border_color": "#333333",
        "valign": "vcenter",
    }
    body_center = workbook.add_format({**base, "align": "center"})
    body_left = workbook.add_format({**base, "align": "left"})
    body_wrap = workbook.add_format({**base, "align": "left", "text_wrap": True})
    manual_green = workbook.add_format(
        {**base, "align": "center", "bg_color": "#E2F0D9"}
    )
    manual_yellow = workbook.add_format(
        {**base, "align": "center", "bg_color": "#FFFACD"}
    )
    manual_gray = workbook.add_format(
        {**base, "align": "center", "bg_color": "#D9D9D9"}
    )
    header = workbook.add_format(
        {
            **base,
            "bold": True,
            "align": "center",
            "text_wrap": True,
        }
    )
    title = workbook.add_format(
        {"font_name": "Arial", "font_size": 16, "bold": True, "align": "center"}
    )
    note_format = workbook.add_format(
        {"font_name": "Arial", "font_size": 10, "font_color": "#FF0000"}
    )
    signature = workbook.add_format(
        {
            "font_name": "Arial",
            "font_size": 10,
            "align": "center",
            "valign": "vcenter",
            "text_wrap": True,
            "border": 2,
        }
    )

    worksheet.merge_range(0, 0, 0, 9, "Check Sheet อุปกรณ์ร้านปิด", title)
    worksheet.set_row(0, 20)

    branch_codes = _unique_values(rows["branch_code"])
    branch_names = [_branch_name(value) for value in _unique_values(rows["branch_name"])]
    regions = _unique_values(rows["region"])
    header_text = (
        f"รหัสสาขา {', '.join(branch_codes)}    "
        f"ชื่อสาขา {', '.join(branch_names)}    "
        f"ภาค {', '.join(regions)}    "
        f"เลขที่เอกสาร {document_no or '-'}    "
        f"วันที่ {created_at.strftime('%d/%m/%Y')}"
    )
    worksheet.insert_textbox(
        2,
        1,
        header_text,
        {
            "width": 1280,
            "height": 48,
            "font": {"name": "Arial", "size": 13, "bold": True},
            "align": {"horizontal": "center", "vertical": "middle"},
            "line": {"color": "#000000", "width": 1.5},
            "fill": {"color": "#FFFFFF", "transparency": 100},
            "object_position": 1,
        },
    )

    headers = [
        "ลำดับ",
        "รหัสทรัพย์สิน",
        "มีบาร์โค้ด",
        "Sno",
        "มีอุปกรณ์",
        "รายการทรัพย์สิน",
        "จำนวนอุปกรณ์ส่วนควบ",
        "S/N Printer POS",
        "Location",
        "หมายเหตุ",
    ]
    worksheet.set_row(6, 22.25)
    for column, value in enumerate(headers):
        worksheet.write_string(6, column, value, header)

    first_data_row = 7
    for index, (_, row) in enumerate(rows.iterrows(), start=1):
        excel_row = first_data_row + index - 1
        accessory = _text(row["_accessory"])
        height = 65 if "\n" in accessory else (30 if len(accessory) > 85 else 18)
        worksheet.set_row(excel_row, height)
        worksheet.write_number(excel_row, 0, index, body_center)
        worksheet.write_string(
            excel_row,
            1,
            _identifier(row["asset_code"], remove_leading_zeroes=True),
            body_center,
        )
        worksheet.write_blank(excel_row, 2, None, manual_green)
        worksheet.write_string(excel_row, 3, _text(row["sno"]), body_center)
        worksheet.write_blank(excel_row, 4, None, manual_yellow)
        worksheet.write_string(excel_row, 5, _text(row["item_name"]), body_left)
        worksheet.write_string(excel_row, 6, accessory, body_wrap)
        worksheet.write_blank(excel_row, 7, None, manual_gray)
        worksheet.write_blank(excel_row, 8, None, body_center)
        worksheet.write_string(excel_row, 9, _row_note(row), body_wrap)

        image_relative = _text(row["_lookup_image"])
        if image_relative:
            image_path = master_dir / image_relative
            if image_path.exists():
                worksheet.insert_image(
                    excel_row,
                    6,
                    str(image_path),
                    {
                        "x_offset": 440,
                        "y_offset": 4,
                        "x_scale": 0.44,
                        "y_scale": 0.44,
                        "object_position": 1,
                    },
                )

    last_data_row = first_data_row + len(rows) - 1
    note_row = last_data_row + 2
    worksheet.merge_range(
        note_row,
        0,
        note_row,
        9,
        "หมายเหตุ : 1.อุปกรณ์ปั๊มน้ำไม่รับเข้าคลัง White Warehouse "
        "เนื่องจากเป็นอุปกรณ์ของ CP Retailink",
        note_format,
    )
    worksheet.merge_range(
        note_row + 1,
        0,
        note_row + 1,
        9,
        "                 2. อุปกรณ์ที่รหัสทรัพย์สินขึ้นต้นด้วย 980XXXXX "
        "คืออุปกรณ์เครื่องเช่าไม่รับเข้าคลัง White Warehouse ทุกกรณี",
        note_format,
    )

    signature_row = note_row + 5
    worksheet.merge_range(
        signature_row,
        1,
        signature_row + 6,
        3,
        "ลงชื่อ _______________________\n\nSupplier ขนส่ง _______________\n\n"
        "วันที่ _______________________\nเบอร์โทร ____________________",
        signature,
    )
    worksheet.merge_range(
        signature_row,
        6,
        signature_row + 6,
        7,
        "ลงชื่อ _______________________\n\nAll Now\n\nวันที่ _______________________",
        signature,
    )
    worksheet.print_area(0, 0, signature_row + 6, 9)


def _add_allnow_sheet(
    workbook: xlsxwriter.Workbook,
    rows: pd.DataFrame,
    created_at: datetime,
) -> None:
    worksheet = workbook.add_worksheet("AllNow")
    worksheet.hide_gridlines(2)
    worksheet.set_landscape()
    worksheet.set_paper(9)
    worksheet.fit_to_pages(1, 0)
    worksheet.set_margins(0.2, 0.2, 0.3, 0.3)

    widths = {
        0: 2.5,
        1: 8.0,
        2: 12.1,
        3: 29.8,
        4: 16.8,
        5: 31.8,
        6: 12.0,
        7: 28.1,
        8: 8.0,
        9: 12.6,
        10: 16.2,
        11: 12.8,
    }
    for column, width in widths.items():
        worksheet.set_column(column, column, width)

    base = {
        "font_name": "Arial",
        "font_size": 10,
        "border": 1,
        "border_color": "#333333",
        "valign": "vcenter",
    }
    center = workbook.add_format({**base, "align": "center"})
    left = workbook.add_format({**base, "align": "left"})
    wrap = workbook.add_format({**base, "align": "left", "text_wrap": True})
    header = workbook.add_format(
        {**base, "bold": True, "align": "center", "text_wrap": True}
    )
    title = workbook.add_format(
        {"font_name": "Arial", "font_size": 16, "bold": True, "align": "center"}
    )
    info = workbook.add_format(
        {"font_name": "Arial", "font_size": 10, "align": "center", "text_wrap": True}
    )
    signature = workbook.add_format(
        {
            "font_name": "Arial",
            "font_size": 10,
            "align": "center",
            "valign": "vcenter",
            "text_wrap": True,
            "border": 1,
        }
    )

    worksheet.merge_range(0, 1, 0, 11, "ใบรับอุปกรณ์เข้า Stock Pool - WWH", title)
    worksheet.write_string(1, 8, f"วันที่ {created_at.strftime('%d/%m/%y')}", info)
    documents = _unique_values(rows["document_no"])
    document_text = ", ".join(documents) if documents else "-"
    worksheet.merge_range(1, 10, 1, 11, f"เลขที่เอกสาร {document_text}", info)

    headers = [
        "ลำดับที่",
        "รหัสทรัพย์สิน",
        "ชื่ออุปกรณ์",
        "รุ่น",
        "อุปกรณ์ส่วนควบ",
        "สาขา ",
        "ชื่อสาขา",
        "จำนวน",
        "Location",
        "สถานะการรับเข้า",
        "หมายเหตุ",
    ]
    for offset, value in enumerate(headers, start=1):
        worksheet.write_string(2, offset, value, header)

    first_data_row = 3
    for index, (_, row) in enumerate(rows.iterrows(), start=1):
        excel_row = first_data_row + index - 1
        accessory = _text(row["_accessory"])
        worksheet.set_row(excel_row, 30 if len(accessory) > 55 or "\n" in accessory else 18)
        worksheet.write_number(excel_row, 1, index, center)
        worksheet.write_string(
            excel_row,
            2,
            _identifier(row["asset_code"], remove_leading_zeroes=True),
            center,
        )
        worksheet.write_string(excel_row, 3, _text(row["item_name"]), left)
        worksheet.write_string(excel_row, 4, _text(row["sno"]), center)
        worksheet.write_string(excel_row, 5, accessory, wrap)
        worksheet.write_string(excel_row, 6, _text(row["branch_code"]), center)
        worksheet.write_string(excel_row, 7, _branch_name(row["branch_name"]), left)
        worksheet.write_number(excel_row, 8, 1, center)
        worksheet.write_blank(excel_row, 9, None, center)
        worksheet.write_string(excel_row, 10, "อุปกรณ์ร้านปิด", center)
        worksheet.write_string(excel_row, 11, _row_note(row), wrap)

    last_data_row = first_data_row + len(rows) - 1
    signature_row = last_data_row + 3
    worksheet.merge_range(
        signature_row,
        3,
        signature_row + 4,
        4,
        "ลงชื่อ ______________________\nSupplier ขนส่ง _______________\n"
        "วันที่ ______________________\nเบอร์โทร ____________________",
        signature,
    )
    worksheet.merge_range(
        signature_row,
        7,
        signature_row + 4,
        8,
        "ลงชื่อ ______________________\nAll Now\nวันที่ ______________________",
        signature,
    )
    worksheet.print_area(0, 1, signature_row + 4, 11)


def build_output(
    ams_content: bytes,
    ams_filename: str,
    master_path: Path,
) -> OutputResult:
    source = parse_ams(ams_content, ams_filename)
    lookup = load_accessory_master(master_path)
    filtered, unmatched = prepare_data(source, lookup)

    documents = _unique_values(filtered["document_no"])
    document_groups = documents or [""]
    output = io.BytesIO()
    created_at = datetime.now(BANGKOK_TZ)

    with xlsxwriter.Workbook(output, {"in_memory": True}) as workbook:
        workbook.set_properties(
            {
                "title": "AMS Checklist and AllNow",
                "subject": "Generated from AMS",
                "company": "CP ALL",
            }
        )
        used_sheet_names: set[str] = {"allnow"}
        for document_no in document_groups:
            group = filtered[filtered["document_no"] == document_no].copy()
            sheet_name = _sheet_name(document_no, used_sheet_names)
            _add_checklist_sheet(
                workbook,
                sheet_name,
                group,
                document_no,
                master_path.parent,
                created_at,
            )
        _add_allnow_sheet(workbook, filtered, created_at)

    return OutputResult(
        content=output.getvalue(),
        source_rows=len(source),
        filtered_rows=len(filtered),
        document_count=len(document_groups),
        matched_rows=int(filtered["_matched"].sum()),
        unmatched_items=unmatched,
        branch_codes=_unique_values(filtered["branch_code"]),
        document_numbers=documents,
    )
