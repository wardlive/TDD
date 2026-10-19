import csv
import hashlib
import unittest
from datetime import datetime
from pathlib import Path
from tempfile import TemporaryDirectory
from zipfile import ZipFile

from openpyxl import Workbook, load_workbook

from config.settings import Settings
from core.file_manager import FileManager
from processors.diesel import DieselProcessor


class DieselTests(unittest.TestCase):
    def setUp(self):
        self.temp = TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.settings = Settings(data_dir=self.root / 'data', auto_archive=False)
        self.files = FileManager(self.settings)
        self.files.ensure_directories()
        self.processor = DieselProcessor(self.files, self.settings)
        self.reference = self.root / 'reference.xlsx'
        book = Workbook()
        sheet = book.active
        sheet.title = DieselProcessor.REFERENCE_SHEET
        for column in ('B', 'J', 'K', 'AD', 'AE'):
            sheet[f'{column}1'] = 'Заголовок намеренно не совпадает'
        for n, (article, size, code, cert, date) in enumerate([
            ('EXACT', 37, '0123456789', 'CERT-1', datetime(2025, 10, 31)),
            ('SHOE', 37, 6403519100, 'CERT-2', datetime(2026, 6, 1)),
            ('SHOE', 38, 6403519900, 'CERT-2', datetime(2026, 6, 1)),
            ('AMBIGUOUS', 37, 6403519100, 'CERT-3', datetime(2026, 6, 1)),
            ('AMBIGUOUS', 37, 6403519900, 'CERT-3', datetime(2026, 6, 1)),
            ('INCOMPLETE', 37, 6403519100, None, None),
            ('EXACT', 37, '0123456789', 'CERT-1', datetime(2025, 10, 31)),
            ('SHOE_100', 37, 6403519900, 'EXACT-SUFFIX', datetime(2026, 6, 1)),
        ], 3):
            for column, value in zip(('B', 'K', 'J', 'AD', 'AE'), (article, size, code, cert, date)):
                sheet[f'{column}{n}'] = value
        book.save(self.reference)
        book.close()
        self.target = self.root / 'target.xlsx'
        book = Workbook()
        sheet = book.active
        sheet.title = 'нач'
        sheet['AD2'] = 'Любое название'
        for row, (article, size) in enumerate([
            ('exact', '37.0'), ('SHOE_900', 38), ('SHOE', 37),
            ('MISSING', 37), ('AMBIGUOUS', 37), ('INCOMPLETE', 37),
            ('SHOE_900', 40), ('SHOE_100', 37),
        ], 3):
            sheet[f'A{row}'] = article
            sheet[f'L{row}'] = size
            sheet[f'B{row}'] = f'Сохранить описание {row}'
            sheet[f'AD{row}'].number_format = 'dd.mm.yyyy'
        sheet['E3'] = '=1+2'
        sheet['N3'] = 3.8900000000000006
        sheet.freeze_panes = 'A3'
        book.create_sheet('Не изменять')['A1'] = '=5+5'
        book.save(self.target)
        book.close()
        with ZipFile(self.target) as archive:
            parts = [(item, archive.read(item.filename)) for item in archive.infolist()]
        with ZipFile(self.target, 'w') as archive:
            for item, data in parts:
                if item.filename == 'xl/worksheets/sheet1.xml':
                    data = data.replace(b'<f>1+2</f><v></v>', b'<f>1+2</f><v>3</v>')
                    data = data.replace(b'<f>1+2</f><v />', b'<f>1+2</f><v>3</v>')
                    data = data.replace(b'<f>1+2</f><v/>', b'<f>1+2</f><v>3</v>')
                archive.writestr(item, data)

    def test_matching_and_preservation(self):
        original_hash = hashlib.sha256(self.target.read_bytes()).hexdigest()
        result = self.processor.process_file(self.target, reference_file=self.reference)
        self.assertTrue(result.success, result.message)
        self.assertEqual(result.details['filled_rows'], 4)
        self.assertEqual(result.details['unresolved_rows'], 4)
        self.assertEqual(result.details['matched_without_suffix'], 1)
        self.assertEqual(hashlib.sha256(self.target.read_bytes()).hexdigest(), original_hash)
        with ZipFile(self.target) as before, ZipFile(result.output_file) as after:
            self.assertEqual(before.namelist(), after.namelist())
            self.assertEqual([name for name in before.namelist() if before.read(name) != after.read(name)],
                             ['xl/worksheets/sheet1.xml'])
        for data_only in (False, True):
            before = load_workbook(self.target, data_only=data_only)
            after = load_workbook(result.output_file, data_only=data_only)
            try:
                self.assertEqual(after['нач']['I3'].value, '0123456789')
                self.assertEqual(after['нач']['I4'].value, 6403519900)
                self.assertEqual(after['нач']['I5'].value, 6403519100)
                self.assertEqual(after['нач']['AC10'].value, 'EXACT-SUFFIX')
                self.assertEqual(after['нач']['AD3'].value, datetime(2025, 10, 31))
                for row in before['нач']:
                    for cell in row:
                        if cell.column not in (9, 29, 30):
                            self.assertEqual(cell.value, after['нач'][cell.coordinate].value, cell.coordinate)
                self.assertEqual(after['нач']['E3'].value, 3 if data_only else '=1+2')
            finally:
                before.close()
                after.close()
        with open(result.details['report_file'], encoding='utf-8-sig', newline='') as handle:
            issues = list(csv.DictReader(handle, delimiter=';'))
        self.assertEqual([item['Артикул'] for item in issues], ['MISSING', 'AMBIGUOUS', 'INCOMPLETE', 'SHOE_900'])

    def test_date_style_added_if_template_has_none(self):
        book = load_workbook(self.target)
        book['нач']['AD3'].number_format = 'General'
        book.save(self.target)
        book.close()
        result = self.processor.process_file(self.target, reference_file=self.reference)
        self.assertTrue(result.success, result.message)
        book = load_workbook(result.output_file)
        try:
            self.assertEqual(book['нач']['AD3'].value, datetime(2025, 10, 31))
        finally:
            book.close()

    def test_reference_cannot_be_target(self):
        with self.assertLogs('DieselProcessor', level='ERROR'):
            result = self.processor.process_file(self.reference, reference_file=self.reference)
        self.assertFalse(result.success)
        self.assertIsNone(result.output_file)

    def test_no_matches_creates_report_without_result_workbook(self):
        result = self.processor.process_file(self.target, reference_data={})
        self.assertFalse(result.success)
        self.assertIsNone(result.output_file)
        self.assertTrue(Path(result.details['report_file']).is_file())
        self.assertEqual(list(self.settings.output_dir.glob('*.xlsx')), [])

    def test_fallback_always_checks_size_and_conflicts_never_choose_first(self):
        reference = self.processor.load_reference(self.reference)
        self.assertIsNone(self.processor._match(reference, 'EXACT_900', '38')[0])
        self.assertIsNone(self.processor._match(reference, 'AMBIGUOUS', '37')[0])
        self.assertIsNone(self.processor._match(reference, 'INCOMPLETE', '37')[0])
        self.assertIsNone(self.processor._match(reference, 'SHOE_900', '')[0])


if __name__ == '__main__':
    unittest.main()
