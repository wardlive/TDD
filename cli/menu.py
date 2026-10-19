"""Интерактивное меню TDD-Suite: стрелки, цифровые команды и выбор файлов."""

import logging
import os
import sys
from pathlib import Path

from halo import Halo

if os.name != "nt":
    from simple_term_menu import TerminalMenu
else:
    TerminalMenu = None

from processors import DieselProcessor, FileCleaner, SheetMerger, SpecificationCompressor

logger = logging.getLogger(__name__)

ASCII_LOGO = (
    " _____ ____  ____ ",
    "|_   _|  _ \\|  _ \\ ",
    "  | | | | | | | | |",
    "  | | | |_| | |_| |",
    "  |_| |____/|____/ ",
)
MAIN_ENTRIES = (
    "[1] Сжатие       Группировка товаров и расчёт долей",
    "[2] Очистка      Апострофы и замена D → .",
    "[3] Объединение  Листы PL INV → Общий",
    "[4] DIESEL       ТН ВЭД, сертификаты и даты",
    "[5] Настройки    Параметры обработки",
    "[6] Статистика   Файлы и результаты",
    "[0] Выход",
)


def compose_header(copy) -> str:
    width = max(map(len, ASCII_LOGO))
    return "\n" + "\n".join(
        f"  {left:<{width}}      {right}".rstrip()
        for left, right in zip(ASCII_LOGO, copy)
    ) + "\n"


def choose(entries, *, title, multiple=False, cursor_index=0):
    """Единое управление; цифровой ввод также работает в Windows и без TTY."""
    if TerminalMenu is not None and sys.stdin.isatty() and sys.stdout.isatty() and os.environ.get("TERM") != "dumb":
        controls = "↑↓ Выбор   Пробел Отметить   Enter Продолжить   Esc Назад" if multiple else "↑↓ Выбор   Enter Открыть   Цифры Быстрый выбор   0 Назад/выход"
        return TerminalMenu(
            entries, title=title, menu_cursor="▸ ",
            menu_cursor_style=("fg_cyan", "bold"), menu_highlight_style=("fg_cyan", "bold"),
            shortcut_key_highlight_style=("fg_cyan", "bold"),
            cycle_cursor=True, clear_screen=True, clear_menu_on_exit=True,
            quit_keys=("escape", "q", "Q"), raise_error_on_interrupt=True,
            status_bar=(" ", f"  {controls}  ", " "), status_bar_style=("fg_yellow", "bold"),
            cursor_index=cursor_index, multi_select=multiple, show_multi_select_hint=False,
            multi_select_select_on_accept=False, multi_select_empty_ok=False,
        ).show()
    print(title)
    for index, entry in enumerate(entries, 1):
        print(f"  {index}. {entry}" if multiple else f"  {entry}")
    while True:
        try:
            raw = input("\nНомера файлов через запятую (0 — назад): " if multiple else "\nВаш выбор: ").strip()
        except EOFError:
            return None
        if raw.lower() in {"q", "", "escape"} or (multiple and raw == "0"):
            return None
        if multiple:
            try:
                selected = tuple(dict.fromkeys(int(value.strip()) - 1 for value in raw.split(",")))
                if selected and all(0 <= index < len(entries) for index in selected):
                    return selected
            except ValueError:
                pass
        else:
            for index, entry in enumerate(entries):
                if entry.startswith(f"[{raw}]"):
                    return index
        print("Неверный выбор. Укажите номер из списка.")


def pause():
    try:
        input("\nEnter — вернуться в меню…")
    except EOFError:
        pass


class MainMenu:
    def __init__(self, file_manager, settings):
        self.file_manager = file_manager
        self.settings = settings
        self.browse_directory = settings.input_dir
        self.processors = {
            "compressor": SpecificationCompressor(file_manager, settings),
            "cleaner": FileCleaner(file_manager, settings),
            "merger": SheetMerger(file_manager, settings),
            "diesel": DieselProcessor(file_manager, settings),
        }

    def _header(self):
        return compose_header((
            "", "Таможня даёт добро!", "Обработка Excel-документов",
            f"Входящие: {len(self.file_manager.get_input_files())} файл(ов)", "Результаты: data/output/",
        ))

    def run(self):
        cursor = 0
        actions = (self._compression_menu, lambda: self._run_processor("cleaner"),
                   lambda: self._run_processor("merger"), self._run_diesel,
                   self._settings_menu, self._show_statistics)
        while True:
            selected = choose(MAIN_ENTRIES, title=self._header(), cursor_index=cursor)
            if selected is None or selected == 6:
                print("До свидания!")
                return
            cursor = selected
            try:
                actions[selected]()
            except KeyboardInterrupt:
                print("\nОбработка прервана. Возврат в главное меню.")
                pause()
            except Exception as error:
                logger.exception("Ошибка операции")
                print(f"\n✘ {error}\nПодробности: {self.settings.data_dir / 'tdd_suite.log'}")
                pause()

    def _browse_files(self, *, title, multiple=False, exclude=None):
        directory = self.browse_directory
        while True:
            files = self.file_manager.get_excel_files(directory)
            if exclude:
                files = [path for path in files if path.resolve() != Path(exclude).resolve()]
            entries = ["[1] Выбрать файлы из списка", "[2] Другая папка или полный путь", "[0] Назад"]
            selected = choose(entries, title=f"\n  {title}\n  Папка: {directory}\n  Файлов .xlsx: {len(files)}\n")
            if selected is None or selected == 2:
                return []
            if selected == 1:
                try:
                    raw = input("Путь к папке или файлу .xlsx (Enter — отмена): ").strip().strip("\"'")
                except EOFError:
                    return []
                if not raw:
                    continue
                path = Path(raw).expanduser()
                if path.is_dir():
                    directory = path
                    self.browse_directory = directory
                    continue
                if path.is_file() and path.suffix.lower() == ".xlsx" and not path.name.startswith("~$"):
                    if exclude and path.resolve() == Path(exclude).resolve():
                        print("Эталон нельзя выбрать как заполняемую спецификацию.")
                        pause()
                        continue
                    self.browse_directory = path.parent
                    return [path]
                print("Путь не найден или выбран файл не в формате .xlsx.")
                pause()
                continue
            if not files:
                print("В этой папке нет файлов .xlsx. Выберите другую папку или укажите путь.")
                pause()
                continue
            if multiple:
                indices = choose([path.name for path in files], title=f"\n  {title}\n  {directory}\n", multiple=True)
                return [files[index] for index in indices] if indices is not None else []
            entries = [f"[{index}] {path.name}" for index, path in enumerate(files, 1)] + ["[0] Назад"]
            index = choose(entries, title=f"\n  {title}\n  {directory}\n")
            return [files[index]] if index is not None and index < len(files) else []

    def _run_processor(self, name, **kwargs):
        processor = self.processors[name]
        files = self._browse_files(title=processor.get_name(), multiple=True)
        if files:
            processor.process_batch(files, **kwargs)
            pause()

    def _compression_menu(self):
        selected = choose(("[1] Простое сжатие", "[2] Сжатие с расчётом долей в коробках", "[0] Назад"),
                          title="\n  СЖАТИЕ СПЕЦИФИКАЦИЙ\n  Выберите режим, затем один или несколько файлов.\n",
                          cursor_index=int(self.settings.compression_mode == "with_percentages"))
        if selected in (0, 1):
            self._run_processor("compressor", calculate_percentages=selected == 1)

    def _run_diesel(self):
        files = self._browse_files(title="DIESEL · шаг 1 из 2 — эталон сезонного заказа")
        if not files:
            return
        reference_file = files[0]
        processor = self.processors["diesel"]
        with Halo(text="Читаю эталон DIESEL…", spinner="dots", color="cyan", enabled=sys.stdout.isatty()):
            reference = processor.load_reference(reference_file)
        print(f"✔ Эталон загружен: {len(reference)} артикулов.\n  {reference_file}")
        files = self._browse_files(title="DIESEL · шаг 2 из 2 — спецификации для заполнения",
                                   multiple=True, exclude=reference_file)
        if files:
            processor.process_batch(files, reference_data=reference, reference_file=reference_file)
            pause()

    def _settings_menu(self):
        while True:
            mode = "с долями" if self.settings.compression_mode == "with_percentages" else "простое"
            entries = (f"[1] Режим сжатия по умолчанию: {mode}",
                       f"[2] Архивная копия исходников: {'да' if self.settings.auto_archive else 'нет'}",
                       f"[3] Запись сжатых таблиц: {self.settings.excel_engine}", "[0] Назад")
            selected = choose(entries, title="\n  НАСТРОЙКИ\n  Действуют до выхода из программы. Исходники остаются на месте.\n")
            if selected is None or selected == 3:
                return
            if selected == 0:
                self.settings.compression_mode = "simple" if mode == "с долями" else "with_percentages"
            elif selected == 1:
                self.settings.auto_archive = not self.settings.auto_archive
            elif selected == 2:
                self.settings.excel_engine = "xlsxwriter" if self.settings.excel_engine == "openpyxl" else "openpyxl"

    def _show_statistics(self):
        print("\n  СТАТИСТИКА TDD-Suite\n")
        for name, data in self.file_manager.get_statistics().items():
            print(f"  {name}: {data['count']} файл(ов)")
            for filename in reversed(data["files"]):
                print(f"    {filename}")
        pause()
