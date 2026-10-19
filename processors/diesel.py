"""Заполнение DIESEL по буквам столбцов, с проверкой размера и неоднозначности."""

import csv
from collections import defaultdict
from datetime import date, datetime
from pathlib import Path
from posixpath import join, normpath
from xml.dom import minidom
from xml.etree import ElementTree
from zipfile import ZipFile

from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string
from openpyxl.utils.datetime import to_excel
from openpyxl.styles.numbers import is_date_format

from core.base_processor import BaseProcessor, ProcessingResult


class DieselProcessor(BaseProcessor):
    XML_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    REFERENCE_SHEET = "Спецификация исходная"
    TARGET_SHEET = "нач"
    SOURCE = {"article": "B", "size": "K", "code": "J", "certificate": "AD", "date": "AE"}
    TARGET = {"article": "A", "size": "L", "code": "I", "certificate": "AC", "date": "AD"}

    def get_name(self) -> str:
        return "Обработка DIESEL"

    def get_description(self) -> str:
        return "Код ТН ВЭД, номер сертификата и дата из сезонного заказа DIESEL"

    @staticmethod
    def _key(value) -> str:
        return "" if value is None else str(value).strip().upper()

    @classmethod
    def _size(cls, value) -> str:
        text = cls._key(value)
        try:
            return format(float(text.replace(",", ".")), "g")
        except ValueError:
            return text

    def load_reference(self, reference_file: Path, sheet_name: str = REFERENCE_SHEET) -> dict:
        reference_file = Path(reference_file)
        if not self.validate_input_file(reference_file):
            raise ValueError("Для эталона нужен существующий файл .xlsx")
        book = load_workbook(reference_file, read_only=True, data_only=True)
        try:
            if sheet_name not in book.sheetnames:
                raise ValueError(f"В эталоне нет листа «{sheet_name}». Доступны: {', '.join(book.sheetnames)}")
            sheet = book[sheet_name]
            columns = {key: column_index_from_string(letter) - 1 for key, letter in self.SOURCE.items()}
            if sheet.max_column < column_index_from_string(self.SOURCE["date"]):
                raise ValueError("В эталоне нужны столбцы B, J, K, AD, AE")
            reference = defaultdict(set)
            for row in sheet.iter_rows(min_row=3, max_col=columns["date"] + 1, values_only=True):
                article = self._key(row[columns["article"]])
                if not article:
                    continue
                values = tuple(row[columns[key]] for key in ("code", "certificate", "date"))
                reference[article].add((self._size(row[columns["size"]]), values))
            if not reference:
                raise ValueError("В эталоне не найдены артикулы в B начиная со строки 3")
            return dict(reference)
        finally:
            book.close()

    def _match(self, reference: dict, article: str, size: str):
        candidates = reference.get(article)
        fallback = not candidates and "_" in article
        if fallback:
            candidates = reference.get(article.rsplit("_", 1)[0])
        if not candidates:
            return None, "Артикул не найден в эталоне", False
        values = {item[1] for item in candidates}
        if fallback or len(values) > 1:
            if not size:
                return None, "Для выбора данных нужен размер в L", False
            values = {data for candidate_size, data in candidates if candidate_size == size}
            if not values:
                return None, "В эталоне не найден этот размер", False
        if len(values) != 1:
            return None, "Разные данные эталона для одного артикула и размера", False
        data = next(iter(values))
        code, certificate, document_date = data
        code_text = str(code).strip() if code is not None else ""
        if not (code_text.isascii() and code_text.isdigit() and len(code_text) == 10):
            return None, "В эталоне нет корректного десятизначного кода ТН ВЭД", False
        if not isinstance(certificate, str) or not certificate.strip():
            return None, "В эталоне не заполнен номер сертификата", False
        if not isinstance(document_date, (datetime, date)):
            return None, "В эталоне не заполнена дата или она не является датой Excel", False
        return data, "", bool(fallback)

    def _save_updates(self, input_path, output_path, sheet_name, updates, book):
        """Меняем только нужные ячейки: обычное сохранение теряет кэш формул Excel."""
        with ZipFile(input_path) as source:
            manifest = ElementTree.fromstring(source.read("xl/workbook.xml"))
            sheet = next(item for item in manifest.find(f"{{{self.XML_NS}}}sheets") if item.attrib["name"] == sheet_name)
            relationship_id = sheet.attrib["{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"]
            relationships = ElementTree.fromstring(source.read("xl/_rels/workbook.xml.rels"))
            target = next(item.attrib["Target"] for item in relationships if item.attrib["Id"] == relationship_id)
            sheet_path = target.lstrip("/") if target.startswith("/") else normpath(join("xl", target))
            document = minidom.parseString(source.read(sheet_path))
            prefix = document.documentElement.prefix

            def element(name):
                return document.createElementNS(self.XML_NS, f"{prefix}:{name}" if prefix else name)

            rows = {int(node.getAttribute("r")): node for node in document.getElementsByTagNameNS(self.XML_NS, "row")}
            cells = {node.getAttribute("r"): node for node in document.getElementsByTagNameNS(self.XML_NS, "c")}
            styles = minidom.parseString(source.read("xl/styles.xml"))
            cell_styles = styles.getElementsByTagNameNS(self.XML_NS, "cellXfs")[0]
            style_nodes = [node for node in cell_styles.childNodes if node.nodeType == node.ELEMENT_NODE]
            added_styles = {}
            for coordinate, value in updates.items():
                cell = cells.get(coordinate)
                if cell is None:
                    cell = element("c")
                    cell.setAttribute("r", coordinate)
                    row = rows[book[sheet_name][coordinate].row]
                    before = next((node for node in row.childNodes if node.nodeType == node.ELEMENT_NODE
                                   and node.localName == "c" and book[sheet_name][node.getAttribute("r")].column
                                   > book[sheet_name][coordinate].column), None)
                    row.insertBefore(cell, before)
                for node in list(cell.childNodes):
                    if node.nodeType == node.ELEMENT_NODE and node.localName in {"f", "v", "is"}:
                        cell.removeChild(node)
                if isinstance(value, (datetime, date)):
                    if not is_date_format(book[sheet_name][coordinate].number_format):
                        style_id = int(cell.getAttribute("s") or 0)
                        if style_id not in added_styles:
                            style = style_nodes[style_id].cloneNode(deep=True)
                            style.setAttribute("numFmtId", "14")
                            style.setAttribute("applyNumberFormat", "1")
                            added_styles[style_id] = len(style_nodes)
                            style_nodes.append(style)
                            cell_styles.appendChild(style)
                        cell.setAttribute("s", str(added_styles[style_id]))
                    value = to_excel(value, epoch=book.epoch)
                if isinstance(value, str):
                    cell.setAttribute("t", "inlineStr")
                    inline = element("is")
                    text = element("t")
                    text.setAttribute("xml:space", "preserve")
                    text.appendChild(document.createTextNode(value))
                    inline.appendChild(text)
                    cell.appendChild(inline)
                else:
                    cell.setAttribute("t", "n")
                    number = element("v")
                    number.appendChild(document.createTextNode(str(value)))
                    cell.appendChild(number)
            replacements = {sheet_path: document.toxml(encoding="utf-8")}
            if added_styles:
                cell_styles.setAttribute("count", str(len(style_nodes)))
                replacements["xl/styles.xml"] = styles.toxml(encoding="utf-8")
            with self.file_manager.atomic_output(output_path) as temporary_path:
                with ZipFile(temporary_path, "w") as destination:
                    destination.comment = source.comment
                    for entry in source.infolist():
                        destination.writestr(entry, replacements.get(entry.filename, source.read(entry.filename)))

    def process_file(self, input_path: Path, **kwargs) -> ProcessingResult:
        book = None
        try:
            input_path = Path(input_path)
            if not self.validate_input_file(input_path):
                raise ValueError("Для спецификации нужен существующий файл .xlsx")
            reference_file = kwargs.get("reference_file")
            if reference_file and Path(reference_file).resolve() == input_path.resolve():
                raise ValueError("Эталон и заполняемая спецификация должны быть разными файлами")
            reference = kwargs.get("reference_data")
            if reference is None:
                if not reference_file:
                    raise ValueError("Сначала выберите эталон DIESEL")
                reference = self.load_reference(reference_file)

            book = load_workbook(input_path)
            sheet_name = kwargs.get("sheet_name", self.TARGET_SHEET)
            if sheet_name not in book.sheetnames:
                if len(book.sheetnames) != 1:
                    raise ValueError(f"Нет листа «{sheet_name}»; укажите лист спецификации явно")
                sheet_name = book.sheetnames[0]
            sheet = book[sheet_name]
            columns = {key: column_index_from_string(letter) for key, letter in self.TARGET.items()}
            if sheet.max_column < columns["date"]:
                raise ValueError("В спецификации нужны столбцы A, I, L, AC, AD")

            issues = []
            updates = {}
            filled = matched_without_suffix = total = 0
            for row_number in range(3, sheet.max_row + 1):
                article = self._key(sheet.cell(row_number, columns["article"]).value)
                if not article:
                    continue
                total += 1
                size = self._size(sheet.cell(row_number, columns["size"]).value)
                values, reason, fallback = self._match(reference, article, size)
                if reason:
                    issues.append({"Лист": sheet_name, "Строка": row_number, "Артикул": article,
                                   "Размер": size, "Причина": reason})
                    continue
                for key, value in zip(("code", "certificate", "date"), values):
                    cell = sheet.cell(row_number, columns[key])
                    updates[cell.coordinate] = value
                filled += 1
                matched_without_suffix += fallback

            output_path = self.file_manager.generate_output_path(input_path, "_diesel")
            details = {"total_rows": total, "filled_rows": filled, "unresolved_rows": len(issues),
                       "matched_without_suffix": matched_without_suffix, "issues": issues, "warnings": []}
            if issues:
                report_path = output_path.with_suffix(".issues.csv")
                with self.file_manager.atomic_output(report_path) as temporary_path:
                    with temporary_path.open("w", encoding="utf-8-sig", newline="") as handle:
                        writer = csv.DictWriter(handle, fieldnames=list(issues[0]), delimiter=";")
                        writer.writeheader()
                        writer.writerows(issues)
                details["report_file"] = str(report_path)
                details["warnings"].append(f"Без изменений оставлено строк: {len(issues)}. Причины — в отчёте.")
            if not filled:
                return ProcessingResult(False, input_path, message="Не заполнено ни одной строки", details=details)
            self._save_updates(input_path, output_path, sheet_name, updates, book)
            if self.settings.auto_archive and not self.file_manager.archive_file(input_path):
                details["warnings"].append("Не удалось создать архивную копию исходника; результат сохранён.")
            return ProcessingResult(
                True, input_path, output_path,
                f"Заполнено {filled} из {total} строк; по артикулу без суффикса и размеру: {matched_without_suffix}",
                details,
            )
        except Exception as error:
            self.logger.exception("Ошибка обработки DIESEL: %s", input_path)
            return ProcessingResult(False, input_path, message=f"Ошибка DIESEL: {error}")
        finally:
            if book is not None:
                book.close()
