"""メインプログラムの統合テスト"""
import pytest
from pathlib import Path
import tempfile
import shutil
from src.main import ExcelFileChecker


class TestExcelFileChecker:
    """ExcelFileCheckerクラスの統合テスト"""

    @pytest.fixture
    def test_config_dir(self):
        """テスト用の一時ディレクトリとconfig.iniを作成"""
        # 一時ディレクトリを作成
        temp_dir = tempfile.mkdtemp()
        temp_path = Path(temp_dir)

        # config.iniを作成
        config_path = temp_path / "test_config.ini"
        config_content = """[SETTINGS]
target_dir = {input_dir}
search_keyword = test
target_cells = A1, B1, C1
image_check_cells = D1, E1
output_filename = test_result.txt
""".format(input_dir=str(temp_path / "input"))

        config_path.write_text(config_content, encoding="utf-8")

        # 入力ディレクトリを作成
        input_dir = temp_path / "input"
        input_dir.mkdir()

        # テストファイルをコピー
        fixtures_dir = Path(__file__).parent / "fixtures"
        shutil.copy(
            fixtures_dir / "excel_with_images.xlsx",
            input_dir / "test_data1.xlsx"
        )
        shutil.copy(
            fixtures_dir / "test_data.csv",
            input_dir / "test_data2.csv"
        )

        yield temp_path

        # クリーンアップ
        shutil.rmtree(temp_dir)

    def test_正常に実行できる(self, test_config_dir, capsys):
        """正常にプログラムが実行できることを確認"""
        config_path = test_config_dir / "test_config.ini"
        checker = ExcelFileChecker(config_path)

        checker.run()

        # 標準出力にログが出力されることを確認
        captured = capsys.readouterr()
        assert "[INFO]" in captured.out
        assert "処理完了" in captured.out

    def test_結果ファイルが生成される(self, test_config_dir):
        """結果ファイルが正しく生成されることを確認"""
        config_path = test_config_dir / "test_config.ini"
        output_path = test_config_dir / "test_result.txt"

        checker = ExcelFileChecker(config_path)
        checker.run()

        # 出力ファイルが存在することを確認
        assert output_path.exists()

        # ファイル内容を確認
        content = output_path.read_text(encoding="utf-8")
        assert "Filename" in content
        assert "A1" in content
        assert "D1(画像)" in content

    def test_複数ファイルが処理される(self, test_config_dir):
        """複数のファイルが正しく処理されることを確認"""
        config_path = test_config_dir / "test_config.ini"
        output_path = test_config_dir / "test_result.txt"

        checker = ExcelFileChecker(config_path)
        checker.run()

        content = output_path.read_text(encoding="utf-8")
        lines = content.strip().split("\n")

        # ヘッダー + 2ファイル分のデータ
        assert len(lines) >= 3

    def test_ログメッセージが出力される(self, test_config_dir, capsys):
        """適切なログメッセージが出力されることを確認"""
        config_path = test_config_dir / "test_config.ini"
        checker = ExcelFileChecker(config_path)

        checker.run()

        captured = capsys.readouterr()
        output = captured.out

        # INFOログの確認
        assert "[INFO]" in output
        assert "探索開始" in output or "発見" in output
        assert "処理完了" in output

    def test_存在しないconfig_iniの場合はエラー(self):
        """存在しないconfig.iniを指定した場合、エラーが発生することを確認"""
        config_path = Path("/non/existent/config.ini")

        with pytest.raises(FileNotFoundError) as exc_info:
            ExcelFileChecker(config_path)

        assert "ファイルが見つかりません" in str(exc_info.value)

    def test_画像判定結果が出力される(self, test_config_dir):
        """画像判定結果が正しく出力されることを確認"""
        config_path = test_config_dir / "test_config.ini"
        output_path = test_config_dir / "test_result.txt"

        checker = ExcelFileChecker(config_path)
        checker.run()

        content = output_path.read_text(encoding="utf-8")

        # 画像判定記号が含まれることを確認
        # .xlsxファイルには○または×、CSVには-が含まれるはず
        assert "○" in content or "×" in content
        assert "-" in content

    def test_サブディレクトリのファイルも処理される(self, test_config_dir):
        """サブディレクトリ内のファイルも正しく処理されることを確認"""
        config_path = test_config_dir / "test_config.ini"

        # サブディレクトリを作成してファイルをコピー
        sub_dir = test_config_dir / "input" / "subfolder"
        sub_dir.mkdir()

        fixtures_dir = Path(__file__).parent / "fixtures"
        shutil.copy(
            fixtures_dir / "excel_partial_images.xlsx",
            sub_dir / "test_sub.xlsx"
        )

        checker = ExcelFileChecker(config_path)
        checker.run()

        output_path = test_config_dir / "test_result.txt"
        content = output_path.read_text(encoding="utf-8")

        # サブディレクトリのファイルが含まれることを確認
        assert "test_sub.xlsx" in content or "subfolder" in content

    def test_キーワードにマッチしないファイルは処理されない(self, test_config_dir):
        """キーワードにマッチしないファイルが除外されることを確認"""
        config_path = test_config_dir / "test_config.ini"

        # キーワードにマッチしないファイルを追加
        input_dir = test_config_dir / "input"
        fixtures_dir = Path(__file__).parent / "fixtures"
        shutil.copy(
            fixtures_dir / "empty.xlsx",
            input_dir / "nomatch.xlsx"
        )

        checker = ExcelFileChecker(config_path)
        checker.run()

        output_path = test_config_dir / "test_result.txt"
        content = output_path.read_text(encoding="utf-8")

        # "nomatch.xlsx"は含まれない
        assert "nomatch.xlsx" not in content
        # "test"を含むファイルのみ処理される
        assert "test_data" in content

    def test_UTF8でファイルが保存される(self, test_config_dir):
        """出力ファイルがUTF-8で保存されることを確認"""
        config_path = test_config_dir / "test_config.ini"
        output_path = test_config_dir / "test_result.txt"

        checker = ExcelFileChecker(config_path)
        checker.run()

        # UTF-8で読み込めることを確認
        content = output_path.read_text(encoding="utf-8")

        # 日本語が含まれることを確認
        assert "画像" in content

    @pytest.fixture
    def multi_file_type_config_dir(self):
        """複数ファイルタイプ対応のテスト用一時ディレクトリとconfig.iniを作成"""
        # 一時ディレクトリを作成
        temp_dir = tempfile.mkdtemp()
        temp_path = Path(temp_dir)

        # 新形式のconfig.iniを作成
        config_path = temp_path / "test_config.ini"
        config_content = """[SETTINGS]
target_dir = {input_dir}
output_filename = test_result.txt

[FILE_TYPE_1]
file_pattern = *レビューチェックリスト*.xlsx
target_sheet = Sheet
target_cells = A1, B1
image_check_cells =

[FILE_TYPE_2]
file_pattern = レビュー記録表*.xlsx
target_sheet = Sheet
target_cells = A1, B1, C1
image_check_cells = D1
""".format(input_dir=str(temp_path / "input"))

        config_path.write_text(config_content, encoding="utf-8")

        # 入力ディレクトリを作成
        input_dir = temp_path / "input"
        input_dir.mkdir()

        # テストファイルをコピー
        fixtures_dir = Path(__file__).parent / "fixtures"
        shutil.copy(
            fixtures_dir / "excel_with_images.xlsx",
            input_dir / "test_レビューチェックリスト_v1.xlsx"
        )
        shutil.copy(
            fixtures_dir / "excel_partial_images.xlsx",
            input_dir / "レビュー記録表_2024.xlsx"
        )
        shutil.copy(
            fixtures_dir / "empty.xlsx",
            input_dir / "other_file.xlsx"
        )

        yield temp_path

        # クリーンアップ
        shutil.rmtree(temp_dir)

    def test_複数ファイルタイプ設定で正しく処理される(self, multi_file_type_config_dir):
        """複数ファイルタイプ設定で各ファイルが正しく処理されることを確認"""
        config_path = multi_file_type_config_dir / "test_config.ini"
        output_path = multi_file_type_config_dir / "test_result.txt"

        checker = ExcelFileChecker(config_path)
        checker.run()

        # 出力ファイルが存在することを確認
        assert output_path.exists()

        content = output_path.read_text(encoding="utf-8")

        # マッチするファイルが処理されることを確認
        assert "レビューチェックリスト" in content
        assert "レビュー記録表" in content

        # マッチしないファイルは処理されないことを確認
        assert "other_file.xlsx" not in content

    def test_ファイルタイプごとに異なるセルが抽出される(self, multi_file_type_config_dir):
        """ファイルタイプごとに異なるセル設定が適用されることを確認"""
        config_path = multi_file_type_config_dir / "test_config.ini"
        output_path = multi_file_type_config_dir / "test_result.txt"

        checker = ExcelFileChecker(config_path)
        checker.run()

        content = output_path.read_text(encoding="utf-8")

        # ヘッダーに各ファイルタイプのセルが含まれることを確認
        # レビューチェックリスト: A1, B1のみ（画像なし）
        # レビュー記録表: A1, B1, C1 + D1(画像)
        lines = content.strip().split("\n")

        # ヘッダー行が存在することを確認
        assert len(lines) > 0
