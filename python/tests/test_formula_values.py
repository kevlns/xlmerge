import os
import tempfile
import unittest
import zipfile
from datetime import datetime
from pathlib import Path
from unittest import mock

from openpyxl import Workbook, load_workbook
from openpyxl.workbook.defined_name import DefinedName

RESOLVER_DIR = Path(__file__).resolve().parents[1] / "xlsx_resolver"
import sys
sys.path.insert(0, str(RESOLVER_DIR))

import formula_values as fv


class FormulaValuesTests(unittest.TestCase):
    def test_no_formula_returns_source_without_writing_output(self):
        with tempfile.TemporaryDirectory() as d:
            source = Path(d) / "source.xlsx"
            output = Path(d) / "output.xlsx"
            wb = Workbook()
            wb.active["A1"] = "=literal"
            wb.active["A1"].data_type = "s"
            wb.save(source)
            wb.close()
            # openpyxl intentionally stores this as a literal string.
            self.assertEqual(fv.materialize_values(source, output), source.resolve())
            self.assertFalse(output.exists())

    def test_literals_preserve_types_and_equals_prefix(self):
        with tempfile.TemporaryDirectory() as d:
            source = Path(d) / "source.xlsx"
            calculated = Path(d) / "calculated.xlsx"
            output = Path(d) / "output.xlsx"
            wb = Workbook()
            ws = wb.active
            ws["A1"] = "=1+2"
            ws["A2"] = "=1=1"
            ws["A3"] = '= "=literal"'
            ws["A4"] = "=DATE(2020,1,2)"
            ws["A4"].number_format = "yyyy-mm-dd"
            wb.save(source)
            wb.close()
            # This stands in for Excel's recalculated cache and exercises the
            # final, safety-critical typed replacement independently of COM.
            calculated_wb = Workbook()
            calc = calculated_wb.active
            calc["A1"] = 3
            calc["A2"] = True
            calc["A3"] = "=literal"
            calc["A3"].data_type = "s"
            calc["A4"] = datetime(2020, 1, 2)
            calc["A4"].number_format = "yyyy-mm-dd"
            calculated_wb.save(calculated)
            calculated_wb.close()
            fv._materialize_from_result(source, calculated, output,
                                         {("Sheet", "A1"), ("Sheet", "A2"), ("Sheet", "A3"), ("Sheet", "A4")})
            result = load_workbook(output, data_only=False)
            ws = result.active
            self.assertEqual(ws["A1"].value, 3)
            self.assertIs(ws["A2"].value, True)
            self.assertEqual(ws["A3"].value, "=literal")
            self.assertEqual(ws["A3"].data_type, "s")
            self.assertEqual(ws["A4"].value, datetime(2020, 1, 2))
            self.assertEqual(ws["A4"].number_format, "yyyy-mm-dd")
            self.assertFalse(any(c.data_type == "f" for row in ws.iter_rows() for c in row))
            result.close()

    def test_unknown_udf_and_volatile_formula_fail_before_excel(self):
        wb = Workbook()
        ws = wb.active
        ws["A1"] = "=MyCompanyFunction(1)"
        with self.assertRaises(fv.FormulaValueError):
            fv._scan_formulas(wb)
        ws["A1"] = "=NOW()"
        with self.assertRaises(fv.FormulaValueError):
            fv._scan_formulas(wb)
        wb.close()

    def test_defined_name_formula_is_scanned_for_volatile_and_udf(self):
        wb = Workbook()
        wb.defined_names.add(DefinedName("UnsafeVolatile", attr_text="=NOW()"))
        with self.assertRaises(fv.FormulaValueError):
            fv._scan_defined_names(wb)
        wb.defined_names.clear()
        wb.defined_names.add(DefinedName("UnsafeUdf", attr_text="=MyUdf()"))
        with self.assertRaises(fv.FormulaValueError):
            fv._scan_defined_names(wb)
        wb.close()

    def test_formula_dependency_cycle_fails_closed(self):
        wb = Workbook()
        ws = wb.active
        ws["A1"] = "=B1+1"
        ws["B1"] = "=A1+1"
        with self.assertRaises(fv.FormulaValueError):
            fv._scan_formulas(wb)
        wb.close()

    def test_shared_array_formula_xml_is_rejected(self):
        with tempfile.TemporaryDirectory() as d:
            source = Path(d) / "source.xlsx"
            wb = Workbook()
            wb.active["A1"] = "=1+2"
            wb.save(source)
            wb.close()
            rewritten = Path(d) / "rewritten.xlsx"
            with zipfile.ZipFile(source) as src, zipfile.ZipFile(rewritten, "w") as dst:
                for item in src.infolist():
                    content = src.read(item.filename)
                    if item.filename == "xl/worksheets/sheet1.xml":
                        content = content.replace(b"<f>1+2</f>", b'<f t="shared" ref="A1:A2">1+2</f>')
                    dst.writestr(item, content)
            rewritten.replace(source)
            with self.assertRaises(fv.FormulaValueError):
                fv._scan_package(source)


@unittest.skipUnless(os.environ.get("XLSX_EXCEL_E2E") == "1", "set XLSX_EXCEL_E2E=1")
class ExcelIntegrationTests(unittest.TestCase):
    def test_real_excel_snapshot_types(self):
        if not Path(r"C:\Program Files\Microsoft Office\Root\Office16\EXCEL.EXE").is_file():
            self.skipTest("Office16 Excel is not installed")
        with tempfile.TemporaryDirectory() as d:
            source = Path(d) / "source.xlsx"
            output = Path(d) / "snapshot.xlsx"
            wb = Workbook()
            ws = wb.active
            ws.title = "Parameters"
            ws["B1"] = 2
            ws["D5"] = 3
            ws["E5"] = 4
            ws["F1"] = "=D5*100+E5"
            ws["F2"] = "=Parameters!$B$1*D5"
            ws["F3"] = '= "=literal"'
            ws["F4"] = "=1=1"
            ws["F5"] = "=DATE(2020,1,2)"
            ws["F5"].number_format = "yyyy-mm-dd"
            ws["F6"] = '= ""'
            wb.save(source)
            wb.close()
            self.assertEqual(fv.materialize_values(source, output), output.resolve())
            result = load_workbook(output, data_only=False)
            ws = result["Parameters"]
            self.assertEqual(ws["F1"].value, 304)
            self.assertEqual(ws["F2"].value, 6)
            self.assertEqual(ws["F3"].value, "=literal")
            self.assertIs(ws["F4"].value, True)
            self.assertEqual(ws["F5"].value, datetime(2020, 1, 2))
            self.assertEqual(ws["F6"].value, "")
            self.assertFalse(any(c.data_type == "f" for row in ws.iter_rows() for c in row))
            result.close()


if __name__ == "__main__":
    unittest.main()
