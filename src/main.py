"""メインプログラム

Excelファイルチェッカーのメイン実行プログラム。
設定ファイルを読み込み、ファイル探索、セル値抽出、画像判定、結果出力を統合的に実行する。
"""
import argparse
from pathlib import Path
from typing import List, Dict, Any
from src.config_loader import ConfigLoader
from src.file_searcher import FileSearcher
from src.cell_extractor import CellExtractor
from src.image_checker import ImageChecker
from src.output_formatter import OutputFormatter


class ExcelFileChecker:
    """Excelファイルチェッカークラス

    設定ファイルに基づいてファイル探索、セル値抽出、画像判定を実行し、
    結果を整形して出力する。

    Attributes:
        config_path (Path): 設定ファイルのパス
        config (ConfigLoader): 設定ローダー
    """

    def __init__(self, config_path: Path):
        """メインプログラムの初期化

        Args:
            config_path: 設定ファイル（config.ini）のパス

        Raises:
            FileNotFoundError: 設定ファイルが存在しない場合
        """
        if not config_path.exists():
            raise FileNotFoundError(
                f"エラー: ファイルが見つかりません: {config_path}"
            )

        self.config_path = config_path
        self.config = ConfigLoader(config_path)

    def run(self):
        """メイン処理を実行

        以下の処理を順次実行する:
        1. ファイル探索
        2. 各ファイルのセル値抽出
        3. 各ファイルの画像判定
        4. 結果の整形と出力
        """
        self._log_info(f"探索開始: {self.config.target_dir}")

        # ファイル探索
        searcher = FileSearcher(
            target_dir=Path(self.config.target_dir),
            search_keyword=self.config.search_keyword
        )
        matched_files = searcher.search()

        self._log_info(f"発見: {len(matched_files)}件のファイル")

        # 各ファイルを処理
        results = []
        for file_path in matched_files:
            try:
                result = self._process_file(file_path)
                results.append(result)
                self._log_info(f"処理: {file_path.name}")
            except Exception as e:
                self._log_error(f"ファイル処理失敗: {file_path.name} - {str(e)}")
                continue

        # 結果を整形
        formatter = OutputFormatter(
            target_cells=self.config.target_cells,
            image_check_cells=self.config.image_check_cells
        )
        formatted_output = formatter.format_results(results)

        # 結果をファイルに保存
        self._save_results(formatted_output)

        self._log_info(f"処理完了: {len(results)}ファイル処理")

    def _process_file(self, file_path: Path) -> Dict[str, Any]:
        """単一ファイルを処理

        Args:
            file_path: 処理対象ファイルのパス

        Returns:
            処理結果の辞書
            形式: {"filename": str, "cell_values": List[Any], "image_results": List[str]}
        """
        # セル値を抽出（シート名を指定）
        extractor = CellExtractor(file_path)
        cell_values = extractor.extract_cells(
            self.config.target_cells,
            sheet_name=self.config.target_sheet
        )

        # 画像判定を実行（シート名を指定）
        checker = ImageChecker(file_path)
        image_results = checker.check_images(
            self.config.image_check_cells,
            sheet_name=self.config.target_sheet
        )

        return {
            "filename": str(file_path.name),
            "cell_values": cell_values,
            "image_results": image_results
        }

    def _save_results(self, content: str):
        """結果をファイルに保存

        Args:
            content: 保存する内容
        """
        # 設定ファイルと同じディレクトリに出力
        output_path = self.config_path.parent / self.config.output_filename

        try:
            output_path.write_text(content, encoding="utf-8")
            self._log_info(f"結果を保存: {output_path}")
        except Exception as e:
            self._log_error(f"ファイル保存失敗: {output_path} - {str(e)}")
            raise

    def _log_info(self, message: str):
        """INFOレベルのログを出力

        Args:
            message: ログメッセージ
        """
        print(f"[INFO] {message}")

    def _log_warning(self, message: str):
        """WARNINGレベルのログを出力

        Args:
            message: ログメッセージ
        """
        print(f"[WARNING] {message}")

    def _log_error(self, message: str):
        """ERRORレベルのログを出力

        Args:
            message: ログメッセージ
        """
        print(f"[ERROR] {message}")


def main():
    """エントリーポイント

    コマンドライン引数で指定されたconfig.iniを読み込んで実行する。
    -i オプションが指定されない場合は、カレントディレクトリのconfig.iniを使用する。
    """
    # コマンドライン引数のパース
    parser = argparse.ArgumentParser(
        description="Excelファイルチェッカー - 指定ディレクトリからExcel/CSVファイルを探索し、セル値と画像を抽出します"
    )
    parser.add_argument(
        "-i",
        "--ini",
        type=str,
        default="config.ini",
        help="設定ファイル（.ini）のパス（デフォルト: config.ini）"
    )

    args = parser.parse_args()
    config_path = Path(args.ini)

    if not config_path.exists():
        print(f"[ERROR] 設定ファイルが見つかりません: {config_path}")
        print("使い方: python run.py -i <設定ファイルのパス>")
        return

    try:
        checker = ExcelFileChecker(config_path)
        checker.run()
    except Exception as e:
        print(f"[ERROR] 実行エラー: {str(e)}")
        raise


if __name__ == "__main__":
    main()
