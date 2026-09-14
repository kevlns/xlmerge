import os
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "xlsx_resolver"))
from xlsx_conflict import ConflictError, _replace_prepared_outputs


class WritebackRollbackTests(unittest.TestCase):
    def test_later_replace_failure_restores_earlier_files(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            paths = [root / name for name in ('first.xlsx', 'second.xlsx')]
            temps = [root / name for name in ('first.tmp', 'second.tmp')]
            for index, (path, temporary) in enumerate(zip(paths, temps)):
                path.write_bytes(f'original-{index}'.encode())
                temporary.write_bytes(f'new-{index}'.encode())
            replace = os.replace
            def fail_second(source, target):
                if Path(source) == temps[1]:
                    raise PermissionError('Simulated locked workbook')
                return replace(source, target)
            with patch('xlsx_conflict.os.replace', side_effect=fail_second):
                with self.assertRaisesRegex(ConflictError, 'rolled back'):
                    _replace_prepared_outputs(list(zip(paths, temps)))
            for index, path in enumerate(paths):
                self.assertEqual(path.read_bytes(), f'original-{index}'.encode())
            self.assertFalse(list(root.glob('*backup*')))

    def test_failed_rollback_retains_recovery_copy(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            paths = [root / name for name in ('first.xlsx', 'second.xlsx')]
            temps = [root / name for name in ('first.tmp', 'second.tmp')]
            for path, temporary in zip(paths, temps):
                path.write_bytes(b'original')
                temporary.write_bytes(b'new')
            replace = os.replace
            def fail_after_first(source, target):
                if Path(source) != temps[0]:
                    raise PermissionError('Simulated filesystem failure')
                return replace(source, target)
            with patch('xlsx_conflict.os.replace', side_effect=fail_after_first):
                with self.assertRaisesRegex(ConflictError, 'recovery copies retained'):
                    _replace_prepared_outputs(list(zip(paths, temps)))
            copies = list(root.glob('*backup*'))
            self.assertEqual(len(copies), 1)
            self.assertEqual(copies[0].read_bytes(), b'original')
