import tempfile
import unittest
from pathlib import Path

import pymupdf as fitz
from openpyxl import Workbook

from bomcheck.extract import (
    _mk_word,
    load_config,
    parse_table,
    read_bom,
    validate_table,
)
from bomcheck.match import compare, parse_qty


ROOT = Path(__file__).resolve().parents[1]


def grid_words(description_in_qty=False, first_id="1"):
    edges = [0, 20, 150, 220, 300, 330, 400]
    words = [
        _mk_word(4, 5, 18, 10, "ITEM"),
        _mk_word(158, 5, 188, 10, "MATERIAL"),
        _mk_word(228, 5, 248, 10, "SIZE"),
        _mk_word(334, 5, 348, 10, "PART"),
        _mk_word(350, 5, 360, 10, "NO"),
    ]
    for row_index, (item_id, quantity) in enumerate(
        ((first_id, "VIBRATION MOUNT" if description_in_qty else "4 nos"),
         ("2", "VIBRATION MOUNT" if description_in_qty else "2"),
         ("3", "VIBRATION MOUNT" if description_in_qty else "12"),
         ("4", "VIBRATION MOUNT" if description_in_qty else "1.5"))
    ):
        y = 17 + row_index * 12
        words.extend([
            _mk_word(4, y, 18, y + 5, item_id),
            _mk_word(160, y, 180, y + 5, "STEEL"),
            _mk_word(230, y, 242, y + 5, "180"),
            _mk_word(334, y, 358, y + 5, f"VM-00{row_index + 1}"),
        ])
        if not description_in_qty:
            for index, token in enumerate(quantity.split()):
                x = 304 + index * 12
                words.append(_mk_word(x, y, x + 8, y + 5, token))
        else:
            words.append(_mk_word(304, y, 327, y + 5, quantity))
        if description_in_qty:
            continue
        words.extend([
            _mk_word(25, y, 65, y + 5, "VIBRATION"),
            _mk_word(68, y, 108, y + 5, "ISOLATION"),
            _mk_word(111, y, 137, y + 5, "MOUNT"),
        ])
    return words, edges


def comparison_fixture():
    table_rows = [
        dict(id="1", description="PROCESS PUMP", material="ASTM A36", size="DN50",
             qty="2", partno="PU-01", remarks="", bbox=(1, 1, 100, 12), conf=100),
        dict(id="2", description="GATE VALVE", material="ASTM A105", size="DN25",
             qty="1", partno="XV-01", remarks="", bbox=(1, 13, 100, 24), conf=100),
    ]
    drawing = dict(
        pages=[dict(page=0, table=dict(columns=["item", "description", "material", "size", "qty", "partno"],
                                       rows=table_rows), balloons=[])],
        scanned=False, revision="A", text="", extraction_errors=[], extraction_warnings=[],
    )
    bom_rows = [
        dict(row=2, id="1", description="PROCESS PUMP", material="ASTM A216 WCB",
             size="DN65", qty=3, partno="PU-01", remarks=""),
        dict(row=3, id="3", description="SPARE GASKET", material="ASTM A36",
             size="DN10", qty=1, partno="GS-99", remarks=""),
    ]
    bom = dict(rows=bom_rows, colidx={1: "item", 2: "description", 3: "material", 4: "size", 5: "qty"},
               revision="A")
    return drawing, bom


class ExtractionValidationTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.cfg = load_config(ROOT / "bomcheck" / "config.json")

    def test_scanned_grid_infers_unlabelled_description_and_quantity_columns(self):
        words, edges = grid_words()
        table = parse_table(words, self.cfg, fuzzy=True, grid_x=edges)
        self.assertIsNotNone(table)
        self.assertEqual(table["columns"], ["item", "description", "material", "size", "qty", "partno"])
        self.assertEqual(len(table["rows"]), 4)
        row = table["rows"][0]
        self.assertEqual(row["description"], "VIBRATION ISOLATION MOUNT")
        self.assertEqual(row["qty"], "4 nos")
        self.assertEqual([parse_qty(row["qty"]) for row in table["rows"]], [4, 2, 12, 1.5])
        self.assertEqual(validate_table(table)[0], [])

    def test_vector_table_with_header_below_rows_uses_native_word_ids(self):
        words = []
        header_y = 1127
        for x0, x1, text in (
            (2363, 2388, "SR"), (2400, 2437, "NO."),
            (2651, 2705, "PART"), (2720, 2773, "NAME"),
            (2955, 2979, "QTY."),
        ):
            words.append(_mk_word(x0, header_y, x1, header_y + 8, text))

        descriptions = [
            "PART ALPHA", "PART BRAVO", "PART CHARLIE", "PART DELTA",
            "PART ECHO", "PART FOXTROT", "PART GOLF", "PART HOTEL",
            "PART INDIA", "PART JULIET", "PART KILO", "PART LIMA",
            "PART MIKE", "PART NOVEMBER", "PART OSCAR", "PART PAPA",
            "PART QUEBEC",
        ]
        for item_id, description in zip(range(17, 0, -1), descriptions):
            y = 615 + (17 - item_id) * 30
            words.extend([
                _mk_word(2400, y, 2418, y + 8, str(item_id)),
                _mk_word(2461, y, 2530, y + 8, description.split()[0]),
                _mk_word(2540, y, 2610, y + 8, description.split()[1]),
                _mk_word(2960, y, 2973, y + 8, "1"),
            ])

        table = parse_table(words, self.cfg)
        self.assertIsNotNone(table)
        self.assertEqual(table["columns"], ["item", "description", "qty"])
        self.assertEqual(len(table["rows"]), 17)
        self.assertEqual([row["id"] for row in table["rows"]], [str(i) for i in range(1, 18)])
        self.assertEqual([row["description"] for row in table["rows"]], list(reversed(descriptions)))
        self.assertTrue(all(parse_qty(row["qty"]) == 1 for row in table["rows"]))

    def test_quantity_parser_rejects_description_text(self):
        self.assertEqual([parse_qty(value) for value in ("2", "12", "1.5", "4 nos")],
                         [2, 12, 1.5, 4])
        self.assertIsNone(parse_qty("VIBRATION ISOLATION MOUNT"))

    def test_description_in_quantity_cell_is_rejected(self):
        words, edges = grid_words(description_in_qty=True)
        table = parse_table(words, self.cfg, fuzzy=True, grid_x=edges)
        self.assertIsNotNone(table)
        self.assertNotIn("qty", table["columns"])
        errors, _ = validate_table(table)
        self.assertTrue(any("description column" in error for error in errors))
        malformed = dict(columns=["item", "description", "qty"], rows=[
            dict(id="1", description="VIBRATION ISOLATION MOUNT",
                 qty="VIBRATION ISOLATION MOUNT")
        ])
        errors, _ = validate_table(malformed)
        self.assertTrue(any("invalid or unreadable quantity" in error for error in errors))

    def test_unreadable_item_id_is_not_invented(self):
        words, edges = grid_words(first_id="Peal")
        table = parse_table(words, self.cfg, fuzzy=True, grid_x=edges)
        self.assertEqual(table["rows"][0]["id"], "")
        self.assertFalse(table["rows"][0]["id_inferred"])
        errors, warnings = validate_table(table)
        self.assertEqual(errors, [])
        self.assertTrue(any("identifiers are unreadable" in warning for warning in warnings))

    def test_ocr_garbled_header_is_recognized(self):
        doc = fitz.open()
        page = doc.new_page(width=900, height=300)
        headers = [
            (30, "Item No."), (130, "DESCRIPTlON"), (360, "Material"),
            (500, "Specification / Size"), (690, "QTY"), (770, "Part No."),
        ]
        for x, text in headers:
            page.insert_text((x, 60), text, fontsize=9)
        values = [
            (30, "1"), (130, "VIBRATION ISOLATION MOUNT"), (360, "ASTM A36"),
            (500, "180 x 150"), (690, "4 nos"), (770, "VM-001"),
        ]
        for x, text in values:
            page.insert_text((x, 85), text, fontsize=9)
        words = [
            _mk_word(*w[:4], w[4]) for w in page.get_text("words")
        ]
        table = parse_table(words, self.cfg, fuzzy=True)
        doc.close()
        self.assertIsNotNone(table)
        self.assertIn("description", table["columns"])
        self.assertEqual(table["rows"][0]["description"], "VIBRATION ISOLATION MOUNT")
        self.assertEqual(table["rows"][0]["qty"], "4 nos")

    def test_no_detectable_header_does_not_make_comparison_findings(self):
        drawing, bom = comparison_fixture()
        drawing["pages"][0]["table"] = None
        drawing["pages"][0]["extraction_errors"] = [
            "Page 1: no recognizable parts-table header."
        ]
        drawing["extraction_errors"] = drawing["pages"][0]["extraction_errors"]
        result = compare(drawing, bom, self.cfg)
        self.assertEqual(result["reliability"], "LOW")
        self.assertTrue(result["comparison_skipped"])
        self.assertEqual(result["issues"], [])
        self.assertTrue(any("no recognizable parts-table header" in reason
                            for reason in result["reliability_reasons"]))

    def test_bom_with_common_headers_is_read_with_size_column(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "bom.xlsx"
            workbook = Workbook()
            sheet = workbook.active
            sheet.append(["Item No.", "Part No.", "Description", "Material",
                          "Specification / Size", "Qty", "Remarks"])
            sheet.append([1, "VM-001", "VIBRATION ISOLATION MOUNT", "ASTM A36",
                          "180 x 150", "4 nos", "Steel"])
            workbook.save(path)
            bom = read_bom(str(path), self.cfg)
        self.assertEqual(set(bom["colidx"].values()),
                         {"item", "partno", "description", "material", "size", "qty", "remarks"})
        self.assertEqual(bom["rows"][0]["size"], "180 x 150")
        self.assertEqual(bom["rows"][0]["qty"], "4 nos")

    def test_matching_keeps_quantity_spec_omission_and_extra_checks(self):
        drawing, bom = comparison_fixture()
        result = compare(drawing, bom, self.cfg)
        issue_types = {issue["type"] for issue in result["issues"]}
        self.assertEqual(issue_types, {"quantity", "spec", "omission", "extra"})
        self.assertFalse(result["comparison_skipped"])
        self.assertTrue(any(match["status"] == "QTY+SPEC" for match in result["matches"]))


if __name__ == "__main__":
    unittest.main()
