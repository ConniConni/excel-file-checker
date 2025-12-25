"""ファイル探索モジュール

指定ディレクトリ配下を再帰的に探索し、条件に合致するファイルを検索する。
"""
from pathlib import Path
from typing import List


class FileSearcher:
    """ファイル探索クラス

    Attributes:
        target_dir (Path): 探索対象のルートディレクトリ
        search_keyword (str): ファイル名に含まれるキーワード
    """

    # 対応するファイル拡張子
    SUPPORTED_EXTENSIONS = {'.xlsx', '.xls', '.csv'}

    def __init__(self, target_dir: Path, search_keyword: str):
        """ファイル探索の初期化

        Args:
            target_dir: 探索対象のルートディレクトリ
            search_keyword: ファイル名に含まれるキーワード

        Raises:
            FileNotFoundError: 指定されたディレクトリが存在しない場合
        """
        if not target_dir.exists():
            raise FileNotFoundError(
                f"エラー: ディレクトリが見つかりません: {target_dir}"
            )

        if not target_dir.is_dir():
            raise NotADirectoryError(
                f"エラー: 指定されたパスはディレクトリではありません: {target_dir}"
            )

        self.target_dir = target_dir
        self.search_keyword = search_keyword

    def search(self) -> List[Path]:
        """ファイルを再帰的に探索

        指定ディレクトリ配下を再帰的に探索し、以下の条件に合致するファイルを取得する：
        1. ファイル名にキーワードが含まれる
        2. 拡張子が .xlsx, .xls, .csv のいずれか

        Returns:
            マッチしたファイルのPathリスト（空の場合は空リスト）
        """
        matched_files = []

        # rglob で再帰的に全ファイルを探索
        for file_path in self.target_dir.rglob('*'):
            # ディレクトリは除外
            if not file_path.is_file():
                continue

            # 拡張子チェック
            if file_path.suffix not in self.SUPPORTED_EXTENSIONS:
                continue

            # キーワードマッチング（ファイル名に含まれるか）
            if self.search_keyword in file_path.name:
                matched_files.append(file_path)

        return matched_files
