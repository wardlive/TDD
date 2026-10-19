import contextlib
import io
import unittest
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import patch

import pandas as pd
from openpyxl import Workbook, load_workbook

from cli.menu import choose
from config.settings import Settings
from core.file_manager import FileManager
from processors.cleaner import FileCleaner
from processors.compressor import SpecificationCompressor
from processors.merger import SheetMerger


class ProcessingTests(unittest.TestCase):
    def setUp(self):
        self.temp = TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.settings = Settings(data_dir=Path(self.temp.name), auto_archive=False)
        self.files = FileManager(self.settings)
        self.files.ensure_directories()

    def test_box_shares_and_totals_with_nonsequential_indices(self):
        processor = SpecificationCompressor(self.files, self.settings)
        frame = pd.DataFrame([[0.0] * 22 for _ in range(3)], index=[3, 7, 9])
        frame[0] = ['A', 'A', 'B']
        frame[10] = [1, 2, 1]
        frame[11] = [37.0, 38.0, 40.5]
        frame[13] = [1.0, 3.0, 2.0]
        frame[16] = [1.5, 3.5, 2.5]
        frame[19] = [10.0, 20.0, 30.0]
        frame[21] = [1, 1, 2]
        shares = processor._calculate_box_percentages(frame)
        self.assertEqual(shares.iloc[:, 14].tolist(), [0.25, 0.75, 1.0])
        compressed = processor._compress_specifications(shares, True)
        self.assertEqual(compressed.iloc[0, 10], 3)
        self.assertEqual(compressed.iloc[0, 13], 4)
        self.assertEqual(compressed.iloc[0, 16], 5)
        self.assertEqual(compressed.iloc[0, 19], 30)
        self.assertEqual(compressed.iloc[0, 11], '37-38')
        self.assertEqual(compressed.iloc[1, 11], '40.5')
        frame[13] = 0.0
        with self.assertRaises(ValueError):
            processor._calculate_box_percentages(frame)

    def test_read_engine_does_not_follow_write_engine(self):
        self.settings.excel_engine = 'xlsxwriter'
        path = self.settings.input_dir / 'spec.xlsx'
        book = Workbook()
        sheet = book.active
        sheet.append(['Итоги'])
        sheet.append([f'col{i}' for i in range(22)])
        sheet.append(['A'] + [1] * 21)
        book.save(path)
        book.close()
        result = SpecificationCompressor(self.files, self.settings).process_file(path)
        self.assertTrue(result.success, result.message)
        self.assertTrue(result.output_file.is_file())

    def test_cleaner_preserves_sheets_styles_and_formulas(self):
        path = self.settings.input_dir / 'clean.xlsx'
        book = Workbook()
        sheet = book.active
        sheet['A1'] = 'Article'
        sheet['AB1'] = 'Price'
        sheet['AJ1'] = 'Net Weight'
        sheet['AK1'] = 'Gross Weight'
        sheet['A2'] = "'KEEPD"
        sheet['AB2'] = "'12D50"
        sheet['AJ2'] = '1D25'
        sheet['AK2'] = '2D00'
        sheet['B2'] = '=1+2'
        sheet['C2'] = "'=1+2"
        sheet['A2'].number_format = '@'
        book.create_sheet('Другой')['A1'] = 'KEEP'
        book.save(path)
        book.close()
        result = FileCleaner(self.files, self.settings).process_file(path)
        self.assertTrue(result.success, result.message)
        actual = load_workbook(result.output_file)
        try:
            self.assertEqual(actual.active['A2'].value, 'KEEPD')
            self.assertEqual(actual.active['AB2'].value, '12.50')
            self.assertEqual(actual.active['AJ2'].value, '1.25')
            self.assertEqual(actual.active['AK2'].value, '2.00')
            self.assertEqual(actual.active['B2'].value, '=1+2')
            self.assertEqual(actual.active['C2'].data_type, 's')
            self.assertEqual(actual.active['A2'].number_format, '@')
            self.assertEqual(actual['Другой']['A1'].value, 'KEEP')
        finally:
            actual.close()

    def test_merger_wide_columns_and_case_insensitive_invoice(self):
        path = self.settings.input_dir / 'merge.xlsx'
        book = Workbook()
        sheet = book.active
        sheet.title = 'pl inv 123'
        sheet.append([f'col{i}' for i in range(27)])
        sheet.append(['ITEM'] + ['1'] * 26)
        sheet.append(['Total pieces'] + ['99'] * 26)
        book.create_sheet('Не трогать')['A1'] = 'KEEP'
        book.save(path)
        book.close()
        result = SheetMerger(self.files, self.settings).process_file(path)
        self.assertTrue(result.success, result.message)
        actual = load_workbook(result.output_file)
        try:
            self.assertEqual(actual['Общий']['A3'].value, '123')
            self.assertEqual(actual['Общий']['R1'].value, '=SUBTOTAL(9,R3:R3)')
            self.assertEqual(actual['Общий'].max_row, 3)
            self.assertEqual(actual['Не трогать']['A1'].value, 'KEEP')
        finally:
            actual.close()

    def test_atomic_output_removes_partial_file_on_interrupt(self):
        path = self.settings.output_dir / 'result.xlsx'
        with self.assertRaises(KeyboardInterrupt):
            with self.files.atomic_output(path) as temporary:
                temporary.write_bytes(b'partial')
                raise KeyboardInterrupt
        self.assertFalse(path.exists())
        self.assertEqual(list(self.settings.output_dir.iterdir()), [])

    def test_numeric_menu_multiselect_and_exit_without_tty(self):
        with patch('cli.menu.TerminalMenu', None), contextlib.redirect_stdout(io.StringIO()):
            with patch('builtins.input', return_value='0'):
                self.assertEqual(choose(['[1] Run', '[0] Exit'], title='test'), 1)
            with patch('builtins.input', return_value='2,1,2'):
                self.assertEqual(choose(['first', 'second'], title='test', multiple=True), (1, 0))
            with patch('builtins.input', side_effect=EOFError):
                self.assertIsNone(choose(['[0] Exit'], title='test'))


if __name__ == '__main__':
    unittest.main()
