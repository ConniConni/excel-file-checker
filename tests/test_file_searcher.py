"""ファイル探索機能のテスト"""
import pytest
from pathlib import Path
from src.file_searcher import FileSearcher


class TestFileSearcher:
    """FileSearcherクラスのテスト"""

    def test_再帰的にファイルを探索できる(self):
        """指定ディレクトリ配下を再帰的に探索し、マッチするファイルを取得できることを確認"""
        test_dir = Path(__file__).parent / "fixtures" / "test_dir"
        searcher = FileSearcher(test_dir, "日経平均")

        found_files = searcher.search()

        # 3つのファイルが見つかることを確認
        assert len(found_files) == 3

        # ファイル名に「日経平均」が含まれることを確認
        for file_path in found_files:
            assert "日経平均" in file_path.name

    def test_キーワードにマッチしないファイルは除外される(self):
        """キーワードにマッチしないファイルは結果に含まれないことを確認"""
        test_dir = Path(__file__).parent / "fixtures" / "test_dir"
        searcher = FileSearcher(test_dir, "日経平均")

        found_files = searcher.search()

        # other_file.xlsx と readme.txt は含まれないことを確認
        file_names = [f.name for f in found_files]
        assert "other_file.xlsx" not in file_names
        assert "readme.txt" not in file_names

    def test_Excel_CSV_ファイルのみを対象とする(self):
        """Excel(.xlsx, .xls)とCSVファイルのみが対象となることを確認"""
        test_dir = Path(__file__).parent / "fixtures" / "test_dir"
        searcher = FileSearcher(test_dir, "日経平均")

        found_files = searcher.search()

        # 全てのファイルが .xlsx, .xls, .csv のいずれかであることを確認
        for file_path in found_files:
            assert file_path.suffix in ['.xlsx', '.xls', '.csv']

    def test_サブディレクトリも探索される(self):
        """サブディレクトリ内のファイルも探索されることを確認"""
        test_dir = Path(__file__).parent / "fixtures" / "test_dir"
        searcher = FileSearcher(test_dir, "日経平均")

        found_files = searcher.search()
        file_names = [f.name for f in found_files]

        # ルート、sub1、sub1/sub2 のファイルが全て含まれることを確認
        assert "日経平均_2024.xlsx" in file_names  # ルート
        assert "日経平均_data.csv" in file_names    # sub1
        assert "日経平均_report.xlsx" in file_names # sub1/sub2

    def test_存在しないディレクトリの場合はエラー(self):
        """存在しないディレクトリを指定した場合、エラーが発生することを確認"""
        non_existent_dir = Path("/non/existent/directory")

        with pytest.raises(FileNotFoundError) as exc_info:
            FileSearcher(non_existent_dir, "keyword")

        assert "ディレクトリが見つかりません" in str(exc_info.value)

    def test_空のディレクトリの場合は空リストを返す(self):
        """マッチするファイルがない場合、空リストを返すことを確認"""
        test_dir = Path(__file__).parent / "fixtures" / "test_dir"
        searcher = FileSearcher(test_dir, "存在しないキーワード")

        found_files = searcher.search()

        assert found_files == []

    def test_結果がPath型のリストである(self):
        """検索結果がPathオブジェクトのリストであることを確認"""
        test_dir = Path(__file__).parent / "fixtures" / "test_dir"
        searcher = FileSearcher(test_dir, "日経平均")

        found_files = searcher.search()

        assert isinstance(found_files, list)
        for file_path in found_files:
            assert isinstance(file_path, Path)
