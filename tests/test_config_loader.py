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
