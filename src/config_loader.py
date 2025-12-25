"""設定ファイル読み込みモジュール

config.iniから設定を読み込み、プログラムで使用可能な形式に変換する。
"""
import configparser
from pathlib import Path
from typing import List


class ConfigLoader:
    """設定ファイル読み込みクラス

    Attributes:
        target_dir (str): 探索対象のルートディレクトリ
        search_keyword (str): ファイル名に含まれるキーワード
        target_sheet (str): 対象シート名（Excelファイル用）
        target_cells (List[str]): 抽出対象のセルリスト
        image_check_cells (List[str]): 画像判定対象のセルリスト
        output_filename (str): 出力ファイル名
    """

    def __init__(self, config_path: Path):
        """設定ファイルを読み込む

        Args:
            config_path: 設定ファイル（.ini）のパス

        Raises:
            FileNotFoundError: 設定ファイルが存在しない場合
            ValueError: 必須項目が欠けている場合
        """
        if not config_path.exists():
            raise FileNotFoundError(f"エラー: config.iniが見つかりません: {config_path}")

        config = configparser.ConfigParser()
        config.read(config_path, encoding='utf-8')

        # 必須項目のチェック
        required_fields = [
            'target_dir',
            'search_keyword',
            'target_cells',
            'image_check_cells',
            'output_filename'
        ]

        missing_fields = []
        for field in required_fields:
            if not config.has_option('SETTINGS', field):
                missing_fields.append(field)

        if missing_fields:
            raise ValueError(
                f"エラー: 必須項目が設定ファイルに含まれていません: {', '.join(missing_fields)}"
            )

        # 設定値の取得
        self.target_dir = config.get('SETTINGS', 'target_dir')
        self.search_keyword = config.get('SETTINGS', 'search_keyword')
        self.output_filename = config.get('SETTINGS', 'output_filename')

        # target_sheetはオプション（指定がない場合はNone）
        self.target_sheet = config.get('SETTINGS', 'target_sheet', fallback=None)

        # カンマ区切りのリストをパースして空白を除去
        self.target_cells = self._parse_cell_list(
            config.get('SETTINGS', 'target_cells')
        )
        self.image_check_cells = self._parse_cell_list(
            config.get('SETTINGS', 'image_check_cells')
        )

    def _parse_cell_list(self, cell_string: str) -> List[str]:
        """カンマ区切りのセル文字列をリストにパース

        Args:
            cell_string: カンマ区切りのセル文字列（例: "A1, B2, C3"）

        Returns:
            パースされたセルリスト（例: ["A1", "B2", "C3"]）
        """
        # カンマで分割して前後の空白を除去
        return [cell.strip() for cell in cell_string.split(',')]
