import io
import unittest

import pandas as pd

from merger import (
    get_suggested_column_group,
    get_suggested_sheet_groups,
    merge_dataframes,
    parse_workbook,
    validate_compiled_sheet_names,
    worksheet_matches_query,
)


class MergerTests(unittest.TestCase):
    def test_suggested_column_group_handles_accents_and_bilingual_labels(self):
        self.assertEqual(get_suggested_column_group("Catégorie"), "Category")
        self.assertEqual(get_suggested_column_group("Mots clés"), "Keywords")
        self.assertEqual(get_suggested_column_group("Custom Field"), "Custom Field")

    def test_suggested_sheet_groups_match_normalized_names(self):
        suggestions = get_suggested_sheet_groups(["Résumé", "Resume", "Details"])

        self.assertEqual(suggestions["Details"], "Details")
        self.assertEqual(suggestions["Resume"], "Resume")
        self.assertEqual(suggestions["Résumé"], "Resume")

    def test_worksheet_query_matches_sheet_or_file_name(self):
        self.assertTrue(worksheet_matches_query("January.xlsx", "Résumé", "resume"))
        self.assertTrue(worksheet_matches_query("January.xlsx", "Details", "jan detail"))
        self.assertFalse(worksheet_matches_query("January.xlsx", "Details", "february"))

    def test_validate_sheet_names_allows_grouping_duplicates(self):
        self.assertEqual(validate_compiled_sheet_names(["Articles", "Articles"]), [])

    def test_validate_sheet_names_rejects_blank_invalid_and_truncation_collisions(self):
        errors = validate_compiled_sheet_names(
            [
                "Valid",
                "",
                None,
                pd.NA,
                "Bad/Name",
                "This sheet name has a shared prefix one",
                "This sheet name has a shared prefix two",
            ]
        )

        error_text = "\n".join(errors)
        self.assertEqual(len(errors), 3)
        self.assertIn("non-empty", error_text)
        self.assertIn("cannot contain", error_text)
        self.assertIn("31-character", error_text)

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

    def test_parse_workbook_reports_invalid_workbook(self):
        parsed_sheets, error = parse_workbook("not-excel.xlsx", b"not really a workbook")

        self.assertEqual(parsed_sheets, {})
        self.assertIsNotNone(error)
        self.assertIn("not-excel.xlsx", error)

    def test_parse_workbook_returns_sheets_with_source_file_column(self):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            pd.DataFrame({"Title": ["Example"]}).to_excel(
                writer, index=False, sheet_name="Articles"
            )

        parsed_sheets, error = parse_workbook("sample.xlsx", output.getvalue())

        self.assertIsNone(error)
        self.assertEqual(list(parsed_sheets), ["Articles"])
        self.assertEqual(parsed_sheets["Articles"]["Source_File"].tolist(), ["sample.xlsx"])


if __name__ == "__main__":
    unittest.main()
