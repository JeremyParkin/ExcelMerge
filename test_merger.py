import io
import unittest

import pandas as pd

from merger import (
    column_matches_query,
    get_suggested_column_group,
    merge_dataframes,
    parse_uploaded_file,
    worksheet_matches_query,
)


class MergerTests(unittest.TestCase):
    def test_suggested_column_group_handles_accents_and_bilingual_labels(self):
        self.assertEqual(get_suggested_column_group("Catégorie"), "Category")
        self.assertEqual(get_suggested_column_group("Mots clés"), "Keywords")
        self.assertEqual(get_suggested_column_group("Custom Field"), "Custom Field")

    def test_worksheet_query_matches_sheet_or_file_name(self):
        self.assertTrue(worksheet_matches_query("January.xlsx", "Résumé", "resume"))
        self.assertTrue(worksheet_matches_query("January.xlsx", "Details", "jan detail"))
        self.assertFalse(worksheet_matches_query("January.xlsx", "Details", "february"))

    def test_column_query_matches_original_or_output_name(self):
        self.assertTrue(column_matches_query("Mots clés", "Keywords", "keywords"))
        self.assertTrue(column_matches_query("Publication Date", "Date", "pub date"))
        self.assertFalse(column_matches_query("Author", "Writer", "category"))

    def test_merge_dataframes_coalesces_columns_that_map_to_same_name(self):
        df = pd.DataFrame(
            {
                "Title": ["English", None],
                "Titre": [None, "Francais"],
                "Source_File": ["a.xlsx", "a.xlsx"],
            }
        )

        merged = merge_dataframes([df], {"Titre": "Title"})

        self.assertEqual(list(merged.columns), ["Title", "Source_File"])
        self.assertEqual(merged["Title"].tolist(), ["English", "Francais"])

    def test_merge_dataframes_omits_excluded_columns(self):
        df = pd.DataFrame(
            {
                "Title": ["Example"],
                "Ignore Me": ["Nope"],
                "Source_File": ["a.xlsx"],
            }
        )

        merged = merge_dataframes([df], {"Title": "Headline"}, {"Title"})

        self.assertEqual(list(merged.columns), ["Headline", "Source_File"])
        self.assertEqual(merged["Headline"].tolist(), ["Example"])

    def test_parse_uploaded_file_reports_invalid_excel_file(self):
        parsed_sheets, error = parse_uploaded_file("not-excel.xlsx", b"not really a workbook")

        self.assertEqual(parsed_sheets, {})
        self.assertIsNotNone(error)
        self.assertIn("not-excel.xlsx", error)

    def test_parse_uploaded_file_returns_excel_sheets_with_source_file_column(self):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            pd.DataFrame({"Title": ["Example"]}).to_excel(
                writer, index=False, sheet_name="Articles"
            )

        parsed_sheets, error = parse_uploaded_file("sample.xlsx", output.getvalue())

        self.assertIsNone(error)
        self.assertEqual(list(parsed_sheets), ["Articles"])
        self.assertEqual(parsed_sheets["Articles"]["Source_File"].tolist(), ["sample.xlsx"])

    def test_parse_uploaded_file_returns_csv_as_single_source(self):
        parsed_sheets, error = parse_uploaded_file(
            "sample.csv", b"Title,Amount\nJanuary,10\nFebruary,20\n"
        )

        self.assertIsNone(error)
        self.assertEqual(list(parsed_sheets), ["sample.csv"])
        self.assertEqual(parsed_sheets["sample.csv"]["Title"].tolist(), ["January", "February"])
        self.assertEqual(
            parsed_sheets["sample.csv"]["Source_File"].tolist(),
            ["sample.csv", "sample.csv"],
        )


if __name__ == "__main__":
    unittest.main()
