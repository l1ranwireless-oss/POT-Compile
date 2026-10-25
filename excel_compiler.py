import csv
import io
import itertools
import re
from pathlib import Path

import xlrd
from openpyxl import Workbook, load_workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.utils import get_column_letter

MAX_FILES = 20
MAX_UPLOAD_BYTES = 500 * 1024 * 1024
ALLOWED_EXTENSIONS = {".xlsx", ".xls", ".csv"}


def _safe_filename(value):
    name = str(value or "").replace("\\", "/").rsplit("/", 1)[-1]
    name = re.sub(r"[^A-Za-z0-9._-]", "_", name)
    if not name or Path(name).suffix.lower() not in ALLOWED_EXTENSIONS:
        raise ValueError(f"Format file tidak didukung: {name or 'tanpa nama'}")
    return name


def _trim_row(row):
    values = list(row)
    while values and values[-1] in (None, ""):
        values.pop()
    return values


def _header_key(value):
    return " ".join(str(value or "").strip().casefold().split())


def _iter_csv_rows(filename, source):
    stream = io.BytesIO(source) if isinstance(source, bytes) else open(source, "rb")
    try:
        sample = stream.read(8192)
        try:
            sample_text = sample.decode("utf-8-sig")
            encoding = "utf-8-sig"
        except UnicodeDecodeError:
            try:
                sample_text = sample.decode("cp1252")
                encoding = "cp1252"
            except UnicodeDecodeError as error:
                raise ValueError(f"Encoding CSV tidak didukung: {filename}") from error
        try:
            dialect = csv.Sniffer().sniff(sample_text, delimiters=",;\t|")
        except csv.Error:
            dialect = csv.excel
        stream.seek(0)
        text_stream = io.TextIOWrapper(stream, encoding=encoding, newline="")
        for row in csv.reader(text_stream, dialect):
            values = _trim_row(row)
            if values:
                yield values, [False] * len(values), None
    finally:
        stream.close()


def _read_xlsx(filename, source):
    try:
        workbook = load_workbook(
            io.BytesIO(source) if isinstance(source, bytes) else source,
            read_only=True,
            data_only=False,
            keep_links=False,
        )
    except Exception as error:
        raise ValueError(f"File Excel tidak dapat dibaca: {filename}") from error
    if not workbook.worksheets:
        workbook.close()
        raise ValueError(f"Workbook tidak memiliki sheet: {filename}")
    for worksheet in workbook.worksheets:
        rows = worksheet.iter_rows()
        for cells in rows:
            if any(cell.value not in (None, "") for cell in cells):
                source_rows = itertools.chain((cells,), rows)
                break
        else:
            continue
        break
    else:
        workbook.close()
        raise ValueError(f"Header tidak ditemukan pada {filename}")

    def iter_rows():
        for row in source_rows:
            cells = list(row)
            while cells and cells[-1].value in (None, ""):
                cells.pop()
            yield (
                [cell.value for cell in cells],
                [cell.data_type == "f" for cell in cells],
                cells,
            )

    return iter_rows(), workbook, worksheet


def _iter_xls_rows(filename, source):
    try:
        if isinstance(source, bytes):
            workbook = xlrd.open_workbook(file_contents=source, on_demand=True)
        else:
            workbook = xlrd.open_workbook(filename=str(source), on_demand=True)
        if not workbook.nsheets:
            raise ValueError(f"Workbook tidak memiliki sheet: {filename}")
        sheet = workbook.sheet_by_index(0)
        def iter_rows():
            for row_index in range(sheet.nrows):
                row = []
                for cell in sheet.row(row_index):
                    if cell.ctype == xlrd.XL_CELL_EMPTY:
                        row.append(None)
                    elif cell.ctype == xlrd.XL_CELL_DATE:
                        row.append(xlrd.xldate_as_datetime(cell.value, workbook.datemode))
                    else:
                        row.append(cell.value)
                values = _trim_row(row)
                yield values, [False] * len(values), None

        return iter_rows(), workbook, None
    except Exception as error:
        if isinstance(error, ValueError):
            raise
        raise ValueError(f"File Excel tidak dapat dibaca: {filename}") from error


def _validate_header(values, filename):
    headers = _trim_row(values)
    keys = [_header_key(value) for value in headers]
    if not headers or any(not key for key in keys):
        raise ValueError(f"Header kosong ditemukan pada {filename}")
    if len(set(keys)) != len(keys):
        raise ValueError(f"Header duplikat ditemukan pada {filename}")
    return headers, keys


def _read_source(filename, source):
    extension = Path(filename).suffix.lower()
    if extension == ".csv":
        rows = _iter_csv_rows(filename, source)
        workbook = worksheet = None
    elif extension == ".xlsx":
        rows, workbook, worksheet = _read_xlsx(filename, source)
    else:
        rows, workbook, worksheet = _iter_xls_rows(filename, source)

    try:
        header_index = -1
        header_cells = None
        for header_index, (values, _, cells) in enumerate(rows):
            if any(value not in (None, "") for value in values):
                headers, keys = _validate_header(values, filename)
                header_cells = cells
                break
        else:
            raise ValueError(f"Header tidak ditemukan pada {filename}")

        def iter_data_rows():
            for values, formulas, cells in rows:
                if not any(value not in (None, "") for value in values):
                    continue
                if len(values) > len(headers) and any(value not in (None, "") for value in values[len(headers):]):
                    raise ValueError(f"Data memiliki kolom tanpa header pada {filename}")
                yield values[:len(headers)], formulas[:len(headers)], cells

        data_rows = iter_data_rows()
        first_data_row = next(data_rows, None)
        return {
            "filename": filename,
            "extension": extension,
            "headers": headers,
            "keys": keys,
            "rows": itertools.chain((first_data_row,), data_rows) if first_data_row else iter(()),
            "workbook": workbook,
            "worksheet": worksheet,
            "sheet_title": worksheet.title if worksheet else Path(filename).stem,
            "header_index": header_index,
            "header_cells": header_cells,
            "data_style_cells": first_data_row[2] if first_data_row else None,
        }
    except Exception:
        _close_workbook(workbook)
        raise


def _close_workbook(workbook):
    if workbook is None:
        return
    if hasattr(workbook, "close"):
        workbook.close()
    elif hasattr(workbook, "release_resources"):
        workbook.release_resources()


def _copy_cell_style(source_cell, target_cell):
    if source_cell is None or not source_cell.has_style:
        return
    for attribute in ("font", "fill", "border", "alignment", "number_format", "protection"):
        setattr(target_cell, attribute, getattr(source_cell, attribute))


def _input_size(source):
    return len(source) if isinstance(source, bytes) else Path(source).stat().st_size


def compile_workbooks(files, output_path=None):
    if not 2 <= len(files) <= MAX_FILES:
        raise ValueError(f"Pilih antara 2 sampai {MAX_FILES} file")

    prepared = []
    total_size = 0
    for filename, source in files:
        safe_name = _safe_filename(filename)
        if not isinstance(source, (bytes, str, Path)):
            raise ValueError(f"File tidak valid: {safe_name}")
        if _input_size(source) <= 0:
            raise ValueError(f"File kosong atau tidak valid: {safe_name}")
        total_size += _input_size(source)
        if total_size > MAX_UPLOAD_BYTES:
            raise ValueError("Total ukuran file melebihi batas 500 MB")
        prepared.append((safe_name, source))

    records = []
    output_workbook = None
    try:
        for filename, source in prepared:
            records.append(_read_source(filename, source))

        template_index = next(
            (index for index, record in enumerate(records) if Path(record["filename"]).suffix.lower() == ".xlsx"),
            None,
        )
        if template_index is None:
            raise ValueError("Sertakan minimal satu file .xlsx sebagai template agar format tetap terjaga")
        template = records[template_index]
        expected_keys = template["keys"]
        for record in records:
            if set(record["keys"]) != set(expected_keys):
                missing = [key for key in expected_keys if key not in record["keys"]]
                extra = [key for key in record["keys"] if key not in expected_keys]
                detail = []
                if missing:
                    detail.append("kolom hilang: " + ", ".join(missing))
                if extra:
                    detail.append("kolom tambahan: " + ", ".join(extra))
                raise ValueError(f"Header {record['filename']} tidak cocok ({'; '.join(detail)})")

        output_workbook = Workbook(write_only=True)
        template_title = template["sheet_title"]
        template_title = re.sub(r"[\\/*?:\[\]]", "_", template_title)[:31] or "Compiled"
        worksheet = output_workbook.create_sheet(template_title)
        header_row = 1
        header_cells = []
        for index, value in enumerate(template["headers"]):
            cell = WriteOnlyCell(worksheet, value=value)
            source_cell = template["header_cells"][index] if template["header_cells"] else None
            _copy_cell_style(source_cell, cell)
            header_cells.append(cell)
        worksheet.append(header_cells)
        template_styles = template["data_style_cells"]
        output_row_count = 1

        for record in records:
            source_columns = {key: index for index, key in enumerate(record["keys"])}
            for values, formulas, _ in record["rows"]:
                output_row = []
                for column, key in enumerate(expected_keys):
                    source_index = source_columns[key]
                    value = values[source_index] if source_index < len(values) else None
                    cell = WriteOnlyCell(worksheet, value=value)
                    if template_styles and column < len(template_styles):
                        _copy_cell_style(template_styles[column], cell)
                    if isinstance(value, str) and value.startswith("=") and not (
                        source_index < len(formulas) and formulas[source_index]
                    ):
                        cell.data_type = "s"
                    output_row.append(cell)
                worksheet.append(output_row)
                output_row_count += 1

            worksheet.auto_filter.ref = f"A{header_row}:{get_column_letter(len(expected_keys))}{output_row_count}"
        if output_path is None:
            output = io.BytesIO()
            output_workbook.save(output)
            return output.getvalue()
        output_workbook.save(output_path)
        return Path(output_path)
    finally:
        if output_workbook:
            output_workbook.close()
        closed = set()
        for record in records:
            workbook = record["workbook"]
            if workbook and id(workbook) not in closed:
                _close_workbook(workbook)
                closed.add(id(workbook))
