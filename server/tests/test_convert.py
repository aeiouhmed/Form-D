import io
import tempfile
import unittest
import zipfile
from pathlib import Path
from typing import List

import pandas as pd
from unittest.mock import patch

from server.app import convert as convert_module
from server.app.convert import convert_to_k1

NAMESPACE = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


def _column_ref_to_index(cell_ref: str) -> int:
    column_part = "".join(filter(str.isalpha, cell_ref))
    index = 0
    for char in column_part:
        index = index * 26 + (ord(char.upper()) - ord("A") + 1)
    return index - 1


def _extract_rows(xlsx_bytes: bytes, expected_columns: int) -> List[List[str]]:
    import xml.etree.ElementTree as ET

    rows: List[List[str]] = []
    with zipfile.ZipFile(io.BytesIO(xlsx_bytes)) as archive:
        shared_strings: List[str] = []
        if "xl/sharedStrings.xml" in archive.namelist():
            shared_root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
            for si in shared_root.findall("s:si", NAMESPACE):
                text_fragments = [node.text or "" for node in si.findall(".//s:t", NAMESPACE)]
                shared_strings.append("".join(text_fragments))

        sheet_root = ET.fromstring(archive.read("xl/worksheets/sheet1.xml"))
        for row in sheet_root.findall("s:sheetData/s:row", NAMESPACE):
            values = [None] * expected_columns
            for cell in row.findall("s:c", NAMESPACE):
                ref = cell.get("r")
                if not ref:
                    continue
                index = _column_ref_to_index(ref)
                if not (0 <= index < expected_columns):
                    continue
                value = ""
                cell_type = cell.get("t")
                value_node = cell.find("s:v", NAMESPACE)
                if value_node is not None:
                    text = value_node.text or ""
                    if cell_type == "s":
                        value = shared_strings[int(text)] if text else ""
                    else:
                        value = text
                values[index] = value
            rows.append(values)
    return rows


class ConvertToK1Tests(unittest.TestCase):
    def test_method_columns_and_excise_vehicle_fields(self) -> None:
        template_columns = [
            "Country of Origin",
            "HSCode",
            "StatisticalUOM",
            "DeclaredUOM",
            "StatisticalQty",
            "DeclaredQty",
            "ItemAmount",
            "ItemDescription",
            "ItemDescription2",
            "ItemDescription3",
            "ImportDutyMethod",
            "Method",
            "ImportDutyRateExemptedPercentage",
            "ImportDutyRateExemptedSpecific",
            "SSTMethod",
            "Method",
            "SSTRateExemptedPercentage",
            "SSTRateExemptedSpecific",
            "ExciseDutyMethod",
            "Method",
            "ExciseDutyRateExemptedPercentage",
            "ExciseDutyRateExemptedSpecific",
            "VehicleType",
            "VehicleModel",
            "Brand",
            "Engine",
            "Chassis",
            "CC",
            "Year",
        ]

        with tempfile.TemporaryDirectory() as tmpdir:
            template_path = Path(tmpdir) / "template.xlsx"
            pd.DataFrame(columns=template_columns).to_excel(template_path, index=False)

            source_df = pd.DataFrame(
                {
                    "Form Flag": ["Form D", "Other"],
                    "Hs Code": ["1234", "9999"],
                    "Quantity": [10, 20],
                    "Net Weight(Kg)": [5.5, 12.0],
                    "Amount(USD)": [1000, 2000],
                    "Parts Name": ["Widget", "Gadget"],
                }
            )
            buffer = io.BytesIO()
            source_df.to_excel(buffer, index=False)

            with patch.dict(
                convert_module.HS_CODE_TO_UNIT,
                {"123400": "UNT"},
                clear=True,
            ):
                result_bytes = convert_to_k1(
                    buffer.getvalue(),
                    country="MY",
                    template_path=str(template_path),
                )

            rows = _extract_rows(result_bytes, len(template_columns))
            headers = rows[0]
            self.assertEqual(len(rows), 2)
            values = rows[1]

            self.assertEqual(headers, template_columns)

            method_indices = [idx for idx, name in enumerate(headers) if name == "Method"]
            self.assertEqual(len(method_indices), 3)
            method_values = [values[index] for index in method_indices]
            self.assertEqual(method_values[:2], ["E", "E"])
            self.assertIn(method_values[2], ("", None))

            hs_value = values[headers.index("HSCode")]
            self.assertTrue(str(hs_value).endswith("00"))

            uom_value = values[headers.index("StatisticalUOM")]
            declared_uom_value = values[headers.index("DeclaredUOM")]
            self.assertEqual(uom_value, "UNT")
            self.assertEqual(declared_uom_value, uom_value)

            stat_qty_value = float(values[headers.index("StatisticalQty")])
            decl_qty_value = float(values[headers.index("DeclaredQty")])
            self.assertEqual(stat_qty_value, 10.0)
            self.assertEqual(decl_qty_value, 10.0)

            origin_value = values[headers.index("Country of Origin")]
            self.assertEqual(origin_value, "MY")

            for column_name in [
                "ExciseDutyMethod",
                "ExciseDutyRateExemptedPercentage",
                "ExciseDutyRateExemptedSpecific",
                "VehicleType",
                "VehicleModel",
                "Brand",
                "Engine",
                "Chassis",
                "CC",
                "Year",
            ]:
                self.assertIn(column_name, headers)
                column_index = headers.index(column_name)
                self.assertIn(values[column_index], ("", None))

    def test_unit_mapping_drives_quantities(self) -> None:
        template_columns = [
            "Country of Origin",
            "HSCode",
            "StatisticalUOM",
            "DeclaredUOM",
            "StatisticalQty",
            "DeclaredQty",
            "ItemAmount",
            "ItemDescription",
            "ItemDescription2",
            "ImportDutyMethod",
            "Method",
            "ImportDutyRateExemptedPercentage",
            "SSTMethod",
            "Method",
            "SSTRateExemptedPercentage",
            "ExciseDutyMethod",
            "Method",
        ]

        with tempfile.TemporaryDirectory() as tmpdir:
            template_path = Path(tmpdir) / "template.xlsx"
            pd.DataFrame(columns=template_columns).to_excel(template_path, index=False)

            source_df = pd.DataFrame(
                {
                    "Form Flag": ["Form D", "Form D", "Form D", "Form D"],
                    "Hs Code": ["1234", "2345", "3456", "4567"],
                    "Quantity": [1, 2, 3, 4],
                    "Net Weight(Kg)": [10, 20, 30, 40],
                    "Amount(USD)": [100, 200, 300, 400],
                    "Parts Name": ["A", "B", "C", "D"],
                }
            )
            buffer = io.BytesIO()
            source_df.to_excel(buffer, index=False)

            with patch.dict(
                convert_module.HS_CODE_TO_UNIT,
                {"123400": "UNT", "234500": "KGM", "345600": "UNT"},
                clear=True,
            ):
                result_bytes = convert_to_k1(
                    buffer.getvalue(),
                    country="SG",
                    template_path=str(template_path),
                )

            rows = _extract_rows(result_bytes, len(template_columns))
            headers = rows[0]
            data_rows = rows[1:]
            self.assertEqual(len(data_rows), len(source_df))

            uom_index = headers.index("StatisticalUOM")
            declared_uom_index = headers.index("DeclaredUOM")
            stat_qty_index = headers.index("StatisticalQty")
            declared_index = headers.index("DeclaredQty")
            country_index = headers.index("Country of Origin")
            hs_index = headers.index("HSCode")

            for idx, row in enumerate(data_rows):
                uom_value = row[uom_index]
                declared_uom_value = row[declared_uom_index]
                self.assertEqual(declared_uom_value, uom_value)
                self.assertEqual(row[country_index], "SG")
                self.assertTrue(str(row[hs_index]).endswith("00"))

                if uom_value == "UNT":
                    stat_qty_value = float(row[stat_qty_index])
                    declared_qty_value = float(row[declared_index])
                    self.assertEqual(stat_qty_value, float(source_df.loc[idx, "Quantity"]))
                    self.assertEqual(declared_qty_value, float(source_df.loc[idx, "Quantity"]))
                elif uom_value == "KGM":
                    stat_qty_value = float(row[stat_qty_index])
                    declared_qty_value = float(row[declared_index])
                    self.assertEqual(
                        stat_qty_value, float(source_df.loc[idx, "Net Weight(Kg)"])
                    )
                    self.assertEqual(
                        declared_qty_value, float(source_df.loc[idx, "Net Weight(Kg)"])
                    )
                else:
                    self.assertEqual(uom_value, "N/A")
                    self.assertIn(row[stat_qty_index], ("", None))
                    self.assertIn(row[declared_index], ("", None))
