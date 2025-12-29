"""設定ファイル読み込みモジュール

config.iniから設定を読み込み、プログラムで使用可能な形式に変換する。
"""
import configparser
import fnmatch
from pathlib import Path
from typing import List, Dict, Optional


class ConfigLoader:
    """設定ファイル読み込みクラス

    Attributes:
        target_dir (str): 探索対象のルートディレクトリ
        output_filename (str): 出力ファイル名
        file_types (List[Dict]): ファイルタイプごとの設定リスト

        # 後方互換性のための属性（旧形式の設定ファイルに対応）
        search_keyword (str): ファイル名に含まれるキーワード（旧形式のみ）
        target_sheet (str): 対象シート名（旧形式のみ）
        target_cells (List[str]): 抽出対象のセルリスト（旧形式のみ）
        image_check_cells (List[str]): 画像判定対象のセルリスト（旧形式のみ）
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

        # SETTINGS セクションの必須項目チェック
        settings_required_fields = ['target_dir', 'output_filename']
        missing_fields = []
        for field in settings_required_fields:
            if not config.has_option('SETTINGS', field):
                missing_fields.append(field)

        if missing_fields:
            raise ValueError(
                f"エラー: 必須項目が設定ファイルに含まれていません: {', '.join(missing_fields)}"
            )

        # 基本設定の取得
        self.target_dir = config.get('SETTINGS', 'target_dir')
        self.output_filename = config.get('SETTINGS', 'output_filename')

        # 新形式（FILE_TYPE_* セクション）と旧形式を判定
        file_type_sections = [s for s in config.sections() if s.startswith('FILE_TYPE_')]

        if file_type_sections:
            # 新形式: 複数のファイルタイプに対応
            self.file_types = self._load_file_types(config, file_type_sections)
            # 旧形式の属性はNoneに設定
            self.search_keyword = None
            self.target_sheet = None
            self.target_cells = None
            self.image_check_cells = None
        else:
            # 旧形式: 後方互換性のため従来の動作を維持
            self._load_legacy_format(config)
            self.file_types = []

    def _load_file_types(self, config: configparser.ConfigParser, sections: List[str]) -> List[Dict]:
        """FILE_TYPE_* セクションから各ファイルタイプの設定を読み込む

        Args:
            config: ConfigParserインスタンス
            sections: FILE_TYPE_* セクション名のリスト

        Returns:
            ファイルタイプ設定の辞書のリスト
        """
        file_types = []

        for section in sections:
            file_type_config = {}

            # 必須項目のチェック
            required_fields = ['file_pattern', 'target_sheet', 'target_cells']
            missing = [f for f in required_fields if not config.has_option(section, f)]

            if missing:
                raise ValueError(
                    f"エラー: セクション '{section}' に必須項目が含まれていません: {', '.join(missing)}"
                )

            # 設定値の取得
            file_type_config['file_pattern'] = config.get(section, 'file_pattern')
            file_type_config['target_sheet'] = config.get(section, 'target_sheet')
            file_type_config['target_cells'] = self._parse_cell_list(
                config.get(section, 'target_cells')
            )

            # image_check_cells はオプション（空文字列の場合は空リスト）
            image_check_cells_str = config.get(section, 'image_check_cells', fallback='')
            if image_check_cells_str.strip():
                file_type_config['image_check_cells'] = self._parse_cell_list(image_check_cells_str)
            else:
                file_type_config['image_check_cells'] = []

            file_types.append(file_type_config)

        return file_types

    def _load_legacy_format(self, config: configparser.ConfigParser):
        """旧形式の設定ファイルを読み込む（後方互換性）

        Args:
            config: ConfigParserインスタンス

        Raises:
            ValueError: 必須項目が欠けている場合
        """
        # 旧形式の必須項目のチェック
        required_fields = ['search_keyword', 'target_cells', 'image_check_cells']
        missing_fields = []
        for field in required_fields:
            if not config.has_option('SETTINGS', field):
                missing_fields.append(field)

        if missing_fields:
            raise ValueError(
                f"エラー: 必須項目が設定ファイルに含まれていません: {', '.join(missing_fields)}"
            )

        # 設定値の取得
        self.search_keyword = config.get('SETTINGS', 'search_keyword')
        self.target_sheet = config.get('SETTINGS', 'target_sheet', fallback=None)
        self.target_cells = self._parse_cell_list(
            config.get('SETTINGS', 'target_cells')
        )
        self.image_check_cells = self._parse_cell_list(
            config.get('SETTINGS', 'image_check_cells')
        )

    def get_file_type_config(self, filename: str) -> Optional[Dict]:
        """ファイル名にマッチするファイルタイプ設定を取得

        Args:
            filename: ファイル名（パスではなくファイル名のみ）

        Returns:
            マッチしたファイルタイプ設定の辞書。マッチしない場合はNone
        """
        for file_type in self.file_types:
            if fnmatch.fnmatch(filename, file_type['file_pattern']):
                return file_type
        return None

    def _parse_cell_list(self, cell_string: str) -> List[str]:
        """カンマ区切りのセル文字列をリストにパース

        Args:
            cell_string: カンマ区切りのセル文字列（例: "A1, B2, C3"）

        Returns:
            パースされたセルリスト（例: ["A1", "B2", "C3"]）
        """
        # カンマで分割して前後の空白を除去
        return [cell.strip() for cell in cell_string.split(',') if cell.strip()]
