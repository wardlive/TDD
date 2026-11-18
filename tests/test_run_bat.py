import unittest
from pathlib import Path


class RunBatTests(unittest.TestCase):
    def test_run_bat_is_ascii_with_crlf(self):
        # cmd.exe misparses non-ASCII text in .bat files, see run.bat
        data = (Path(__file__).resolve().parent.parent / 'run.bat').read_bytes()
        data.decode('ascii')
        self.assertNotIn(b'\n', data.replace(b'\r\n', b''))
