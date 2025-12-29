"""設定ファイル読み込み機能のテスト"""
import pytest
from pathlib import Path
from src.config_loader import ConfigLoader


class TestConfigLoader:
    """ConfigLoaderクラスのテスト"""

    def test_正常な設定ファイルの読み込み(self):
        """正常な設定ファイルが正しく読み込めることを確認"""
        config_path = Path(__file__).parent / "fixtures" / "test_config.ini"
        loader = ConfigLoader(config_path)

        assert loader.target_dir == "./test_input"
        assert loader.search_keyword == "テスト"
        assert loader.target_sheet == "Sheet1"
        assert loader.target_cells == ["A1", "B2", "C3"]
        assert loader.image_check_cells == ["D1", "E2"]
        assert loader.output_filename == "test_output.txt"

    def test_設定ファイルが存在しない場合はエラー(self):
        """存在しない設定ファイルを指定した場合、エラーが発生することを確認"""
        config_path = Path(__file__).parent / "fixtures" / "non_existent.ini"

        with pytest.raises(FileNotFoundError) as exc_info:
            ConfigLoader(config_path)

        assert "config.iniが見つかりません" in str(exc_info.value)

    def test_必須項目が欠けている場合はエラー(self):
        """必須項目が欠けている設定ファイルの場合、エラーが発生することを確認"""
        config_path = Path(__file__).parent / "fixtures" / "invalid_config.ini"

        with pytest.raises(ValueError) as exc_info:
            ConfigLoader(config_path)

        # エラーメッセージに不足している項目が含まれることを確認
        error_msg = str(exc_info.value)
        assert "target_cells" in error_msg or "output_filename" in error_msg

    def test_カンマ区切りリストのパース(self):
        """カンマ区切りのセルリストが正しくリスト型にパースされることを確認"""
        config_path = Path(__file__).parent / "fixtures" / "test_config.ini"
        loader = ConfigLoader(config_path)

        # target_cellsがリスト型であることを確認
        assert isinstance(loader.target_cells, list)
        assert len(loader.target_cells) == 3
        assert loader.target_cells[0] == "A1"
        assert loader.target_cells[1] == "B2"
        assert loader.target_cells[2] == "C3"

        # image_check_cellsがリスト型であることを確認
        assert isinstance(loader.image_check_cells, list)
        assert len(loader.image_check_cells) == 2
        assert loader.image_check_cells[0] == "D1"
        assert loader.image_check_cells[1] == "E2"

    def test_空白文字が適切にトリムされること(self):
        """カンマ区切りリストの前後の空白が除去されることを確認"""
        config_path = Path(__file__).parent / "fixtures" / "test_config.ini"
        loader = ConfigLoader(config_path)

        # 空白が含まれていないことを確認
        for cell in loader.target_cells:
            assert cell == cell.strip()

        for cell in loader.image_check_cells:
            assert cell == cell.strip()

    def test_複数ファイルタイプの設定ファイルの読み込み(self):
        """複数のファイルタイプを持つ設定ファイルが正しく読み込めることを確認"""
        config_path = Path(__file__).parent / "fixtures" / "multi_file_type_config.ini"
        loader = ConfigLoader(config_path)

        # 基本設定の確認
        assert loader.target_dir == "./test_input"
        assert loader.output_filename == "test_output.txt"

        # ファイルタイプの数を確認
        assert len(loader.file_types) == 2

        # ファイルタイプ1の確認
        file_type_1 = loader.file_types[0]
        assert file_type_1["file_pattern"] == "*0_レビューチェックリスト*.xlsx"
        assert file_type_1["target_sheet"] == "sheet1"
        assert file_type_1["target_cells"] == ["E4", "E5", "E6", "N6", "M10", "M11", "M15", "M16"]
        assert file_type_1["image_check_cells"] == []

        # ファイルタイプ2の確認
        file_type_2 = loader.file_types[1]
        assert file_type_2["file_pattern"] == "レビュー記録表(社*.xlsx"
        assert file_type_2["target_sheet"] == "sheet1"
        assert file_type_2["target_cells"] == ["AE2", "AE4", "AE5", "AE6", "AE7", "AE8", "AB17"]
        assert file_type_2["image_check_cells"] == ["BY3"]

    def test_ファイル名パターンでファイルタイプを取得できる(self):
        """ファイル名からマッチするファイルタイプ設定を取得できることを確認"""
        config_path = Path(__file__).parent / "fixtures" / "multi_file_type_config.ini"
        loader = ConfigLoader(config_path)

        # レビューチェックリストに対応するファイル名でテスト
        file_type = loader.get_file_type_config("test_0_レビューチェックリスト_v1.xlsx")
        assert file_type is not None
        assert file_type["file_pattern"] == "*0_レビューチェックリスト*.xlsx"
        assert file_type["target_sheet"] == "sheet1"

        # レビュー記録表に対応するファイル名でテスト
        file_type = loader.get_file_type_config("レビュー記録表(社内).xlsx")
        assert file_type is not None
        assert file_type["file_pattern"] == "レビュー記録表(社*.xlsx"
        assert file_type["target_sheet"] == "sheet1"

        # マッチしないファイル名でテスト
        file_type = loader.get_file_type_config("other_file.xlsx")
        assert file_type is None
