"""画像判定機能のテスト"""
import pytest
from pathlib import Path
from src.image_checker import ImageChecker


class TestImageChecker:
    """ImageCheckerクラスのテスト"""

    def test_画像があるセルを正しく判定できる(self):
        """画像が配置されているセルを正しく判定できることを確認"""
        file_path = Path(__file__).parent / "fixtures" / "excel_with_images.xlsx"
        checker = ImageChecker(file_path)

        results = checker.check_images(["D1", "E1"])

        assert len(results) == 2
        assert results[0] == "○"  # D1に画像あり
        assert results[1] == "○"  # E1に画像あり

    def test_画像がないセルを正しく判定できる(self):
        """画像が配置されていないセルを正しく判定できることを確認"""
        file_path = Path(__file__).parent / "fixtures" / "excel_no_images.xlsx"
        checker = ImageChecker(file_path)

        results = checker.check_images(["D1", "E1"])

        assert len(results) == 2
        assert results[0] == "×"  # D1に画像なし
        assert results[1] == "×"  # E1に画像なし

    def test_一部のセルにのみ画像がある場合(self):
        """一部のセルにのみ画像がある場合、正しく判定できることを確認"""
        file_path = Path(__file__).parent / "fixtures" / "excel_partial_images.xlsx"
        checker = ImageChecker(file_path)

        results = checker.check_images(["D1", "E1"])

        assert len(results) == 2
        assert results[0] == "○"  # D1に画像あり
        assert results[1] == "×"  # E1に画像なし

    def test_複数セルの画像判定(self):
        """複数のセルを指定して画像判定できることを確認"""
        file_path = Path(__file__).parent / "fixtures" / "excel_with_images.xlsx"
        checker = ImageChecker(file_path)

        results = checker.check_images(["A1", "D1", "E1", "F1"])

        assert len(results) == 4
        assert results[0] == "×"  # A1に画像なし
        assert results[1] == "○"  # D1に画像あり
        assert results[2] == "○"  # E1に画像あり
        assert results[3] == "×"  # F1に画像なし

    def test_CSVファイルの場合は非対応マークを返す(self):
        """CSVファイルの場合、非対応を示す"-"を返すことを確認"""
        file_path = Path(__file__).parent / "fixtures" / "test_data.csv"
        checker = ImageChecker(file_path)

        results = checker.check_images(["D1", "E1"])

        assert len(results) == 2
        assert results[0] == "-"  # CSVは画像判定非対応
        assert results[1] == "-"  # CSVは画像判定非対応

    def test_xlsファイルの場合は非対応マークを返す(self):
        """古い形式(.xls)のExcelファイルの場合、非対応を示す"-"を返すことを確認"""
        file_path = Path(__file__).parent / "fixtures" / "test_data.xls"
        checker = ImageChecker(file_path)

        results = checker.check_images(["D1", "E1"])

        assert len(results) == 2
        assert results[0] == "-"  # .xlsは画像判定非対応
        assert results[1] == "-"  # .xlsは画像判定非対応

    def test_存在しないファイルの場合はエラー(self):
        """存在しないファイルを指定した場合、エラーが発生することを確認"""
        file_path = Path(__file__).parent / "fixtures" / "non_existent.xlsx"

        with pytest.raises(FileNotFoundError) as exc_info:
            ImageChecker(file_path)

        assert "ファイルが見つかりません" in str(exc_info.value)

    def test_戻り値がリスト型である(self):
        """判定結果がリスト型であることを確認"""
        file_path = Path(__file__).parent / "fixtures" / "excel_with_images.xlsx"
        checker = ImageChecker(file_path)

        results = checker.check_images(["D1", "E1"])

        assert isinstance(results, list)
        assert len(results) == 2

    def test_空のセルリストの場合は空リストを返す(self):
        """空のセルリストを渡した場合、空リストを返すことを確認"""
        file_path = Path(__file__).parent / "fixtures" / "excel_with_images.xlsx"
        checker = ImageChecker(file_path)

        results = checker.check_images([])

        assert results == []

    def test_シート名を指定して画像判定できる(self):
        """シート名を指定して特定シートの画像判定ができることを確認"""
        file_path = Path(__file__).parent / "fixtures" / "multi_sheet_with_images.xlsx"
        checker = ImageChecker(file_path)

        # Sheet1は画像なし
        results_sheet1 = checker.check_images(["D1"], sheet_name="Sheet1")
        assert results_sheet1[0] == "×"

        # Sheet2は画像あり
        results_sheet2 = checker.check_images(["D1"], sheet_name="Sheet2")
        assert results_sheet2[0] == "○"

    def test_シート名を指定しない場合は最初のシートを使用(self):
        """シート名を指定しない場合、最初のシート（インデックス0）を使用することを確認"""
        file_path = Path(__file__).parent / "fixtures" / "multi_sheet_with_images.xlsx"
        checker = ImageChecker(file_path)

        # シート名を指定しない（最初のシートSheet1が使用される）
        results = checker.check_images(["D1"])

        # Sheet1には画像がないので×
        assert results[0] == "×"
