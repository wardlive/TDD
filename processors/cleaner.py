"""Очистка первого листа с сохранением остальных листов, формул и оформления."""

from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

from core.base_processor import BaseProcessor, ProcessingResult


class FileCleaner(BaseProcessor):
    def get_name(self) -> str:
        return "Очистка файлов"

    def get_description(self) -> str:
        return "Апострофы и D → . в Price, Net Weight, Gross Weight (резерв: AB, AJ, AK)"

    def process_file(self, input_path: Path, **kwargs) -> ProcessingResult:
        book = None
        try:
            if not self.validate_input_file(input_path):
                raise ValueError("Для очистки нужен существующий файл .xlsx")
            book = load_workbook(input_path)
            sheet = book.worksheets[0]
            columns = self._find_target_columns(sheet)
            changes = {"apostrophes_removed": 0, "d_to_dot_replacements": 0,
                       "processed_cells": 0, "target_columns_found": columns}
            indices = {column["index"] for column in columns}
            for row in sheet.iter_rows(min_row=2):
                for cell in row:
                    if cell.data_type == "f" or not isinstance(cell.value, str):
                        continue
                    value = cell.value
                    if value.startswith("'"):
                        value = value[1:]
                        changes["apostrophes_removed"] += 1
                        changes["processed_cells"] += 1
                    if cell.column in indices and "D" in value:
                        value = value.replace("D", ".")
                        changes["d_to_dot_replacements"] += 1
                        changes["processed_cells"] += 1
                    if value != cell.value:
                        cell.value = value
                        cell.data_type = "s"
            output_path = self.file_manager.generate_output_path(input_path, "_cleaned")
            with self.file_manager.atomic_output(output_path) as temporary_path:
                book.save(temporary_path)
            self.file_manager.archive_file(input_path)
            return ProcessingResult(
                True, input_path, output_path,
                f"Апострофы: {changes['apostrophes_removed']}, D → .: {changes['d_to_dot_replacements']}",
                changes,
            )
        except Exception as error:
            self.logger.exception("Ошибка очистки %s", input_path)
            return ProcessingResult(False, input_path, message=f"Ошибка очистки: {error}")
        finally:
            if book is not None:
                book.close()

    @staticmethod
    def _find_target_columns(sheet) -> list:
        patterns = ("price", "стоимость", "цена", "net weight", "nett weight", "нетто", "gross weight", "брутто")
        columns = [cell.column for cell in sheet[1]
                   if any(pattern in str(cell.value).lower() for pattern in patterns)]
        if not columns:
            columns = [index for index in (28, 36, 37) if index <= sheet.max_column]
        return [{"index": index, "letter": get_column_letter(index), "name": str(sheet.cell(1, index).value)}
                for index in columns]
