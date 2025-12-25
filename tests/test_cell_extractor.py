"""セル値抽出機能のテスト"""
import pytest
from pathlib import Path
from src.cell_extractor import CellExtractor


class TestCellExtractor:
    """CellExtractorクラスのテスト"""

    def test_Excelファイルxlsxからセル値を抽出できる(self):
        """Excelファイル(.xlsx)から指定セルの値を抽出できることを確認"""
        file_path = Path(__file__).parent / "fixtures" / "test_data.xlsx"
        extractor = CellExtractor(file_path)

        values = extractor.extract_cells(["A1", "B1", "C1"])

        assert len(values) == 3
        assert values[0] == "2023/12/25"
        assert values[1] == 33000.5
        assert values[2] == "確定"

    def test_Excelファイルxlsからセル値を抽出できる(self):
        """Excelファイル(.xls)から指定セルの値を抽出できることを確認"""
        file_path = Path(__file__).parent / "fixtures" / "test_data.xls"
        extractor = CellExtractor(file_path)

        values = extractor.extract_cells(["A1", "B1", "C1"])

        assert len(values) == 3
        assert values[0] == "2023/12/25"
        assert values[1] == 33000.5
        assert values[2] == "確定"

    def test_CSVファイルからセル値を抽出できる(self):
        """CSVファイルから指定セルの値を抽出できることを確認"""
        file_path = Path(__file__).parent / "fixtures" / "test_data.csv"
        extractor = CellExtractor(file_path)

        values = extractor.extract_cells(["A1", "B1", "C1"])

        assert len(values) == 3
        assert values[0] == "2023/12/26"
        # CSVから読み込んだ数値は文字列になる可能性がある
        assert str(values[1]) == "33100.2"
        assert values[2] == "暫定"

    def test_複数セルを指定して抽出できる(self):
        """複数のセルを指定して一度に抽出できることを確認"""
        file_path = Path(__file__).parent / "fixtures" / "test_data.xlsx"
        extractor = CellExtractor(file_path)

        values = extractor.extract_cells(["A1", "B2", "C3"])

        assert len(values) == 3
        assert values[0] == "2023/12/25"  # A1
        assert values[1] == "B2_value"    # B2
        assert values[2] == "C3_value"    # C3

    def test_存在しないセルは空文字列を返す(self):
        """範囲外のセルを指定した場合、空文字列を返すことを確認"""
        file_path = Path(__file__).parent / "fixtures" / "test_data.xlsx"
        extractor = CellExtractor(file_path)

        values = extractor.extract_cells(["A1", "Z99", "C1"])

        assert len(values) == 3
        assert values[0] == "2023/12/25"
        assert values[1] == ""  # 範囲外のセル
        assert values[2] == "確定"

    def test_空のセルは空文字列を返す(self):
        """空のセルは空文字列を返すことを確認"""
        file_path = Path(__file__).parent / "fixtures" / "empty.xlsx"
        extractor = CellExtractor(file_path)

        values = extractor.extract_cells(["A1", "B1"])

        assert len(values) == 2
        assert values[0] == ""
        assert values[1] == ""

    def test_存在しないファイルの場合はエラー(self):
        """存在しないファイルを指定した場合、エラーが発生することを確認"""
        file_path = Path(__file__).parent / "fixtures" / "non_existent.xlsx"

        with pytest.raises(FileNotFoundError) as exc_info:
            CellExtractor(file_path)

        assert "ファイルが見つかりません" in str(exc_info.value)

    def test_サポートされていない拡張子の場合はエラー(self):
        """サポートされていない拡張子のファイルを指定した場合、エラーが発生することを確認"""
        file_path = Path(__file__).parent / "fixtures" / "test_dir" / "readme.txt"

        with pytest.raises(ValueError) as exc_info:
            CellExtractor(file_path)

        assert "サポートされていないファイル形式" in str(exc_info.value)

    def test_セル座標のパース(self):
        """セル座標（例: A1, B2）が正しくパースされることを確認"""
        file_path = Path(__file__).parent / "fixtures" / "test_data.xlsx"
        extractor = CellExtractor(file_path)

        # A1 = (0, 0), B2 = (1, 1), C3 = (2, 2)
        values = extractor.extract_cells(["A1", "B2", "C3"])

        assert values[0] == "2023/12/25"
        assert values[1] == "B2_value"
        assert values[2] == "C3_value"

    def test_戻り値がリスト型である(self):
        """抽出結果がリスト型であることを確認"""
        file_path = Path(__file__).parent / "fixtures" / "test_data.xlsx"
        extractor = CellExtractor(file_path)

        values = extractor.extract_cells(["A1", "B1"])

        assert isinstance(values, list)
        assert len(values) == 2
