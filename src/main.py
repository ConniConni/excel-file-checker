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
from src.review_validator import ReviewValidator


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

        # 新形式（複数ファイルタイプ）と旧形式で処理を分岐
        if self.config.file_types:
            # 新形式: 全ての.xlsxファイルを探索し、各ファイルタイプの設定を適用
            all_files = self._search_all_excel_files()
            matched_files = []
            results = []
            validator = ReviewValidator()

            for file_path in all_files:
                # ファイル名にマッチする設定を取得
                file_type_config = self.config.get_file_type_config(file_path.name)

                if file_type_config:
                    try:
                        result = self._process_file_with_config(file_path, file_type_config)
                        result["file_path"] = file_path
                        results.append(result)
                        matched_files.append(file_path)

                        # ファイルタイプを判定してバリデーターに追加
                        file_type = self._determine_file_type(file_path.name)
                        if file_type:
                            validator.add_file(result, file_type)

                        self._log_info(f"処理: {file_path.name}")
                    except Exception as e:
                        self._log_error(f"ファイル処理失敗: {file_path.name} - {str(e)}")
                        continue

            self._log_info(f"発見: {len(matched_files)}件のファイル")

            # ペアリングと検証を実行
            validation_results = validator.validate_all()

            # 結果を整形（新形式用・検証結果付き）
            formatted_output = self._format_validation_results(
                validation_results,
                root_dir=Path(self.config.target_dir)
            )
        else:
            # 旧形式: search_keywordでファイルを探索
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

            # 結果を整形（旧形式用）
            formatter = OutputFormatter(
                target_cells=self.config.target_cells,
                image_check_cells=self.config.image_check_cells
            )
            formatted_output = formatter.format_results(
                results,
                root_dir=Path(self.config.target_dir),
                file_paths=matched_files
            )

        # 結果をファイルに保存
        self._save_results(formatted_output)

        self._log_info(f"処理完了: {len(results)}ファイル処理")

    def _search_all_excel_files(self) -> List[Path]:
        """対象ディレクトリ配下の全てのExcelファイルを探索

        Returns:
            見つかったExcelファイルのパスリスト
        """
        target_dir = Path(self.config.target_dir)
        excel_files = []

        # .xlsxファイルを再帰的に探索
        for file_path in target_dir.rglob("*.xlsx"):
            if file_path.is_file():
                excel_files.append(file_path)

        return excel_files

    def _process_file(self, file_path: Path) -> Dict[str, Any]:
        """単一ファイルを処理（旧形式用）

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

    def _process_file_with_config(self, file_path: Path, file_type_config: Dict) -> Dict[str, Any]:
        """単一ファイルを指定された設定で処理（新形式用）

        Args:
            file_path: 処理対象ファイルのパス
            file_type_config: ファイルタイプ設定の辞書

        Returns:
            処理結果の辞書
            形式: {"filename": str, "cell_values": List[Any], "image_results": List[str],
                   "target_cells": List[str], "image_check_cells": List[str]}
        """
        # セル値を抽出
        extractor = CellExtractor(file_path)
        cell_values = extractor.extract_cells(
            file_type_config['target_cells'],
            sheet_name=file_type_config['target_sheet']
        )

        # 画像判定を実行
        checker = ImageChecker(file_path)
        image_results = checker.check_images(
            file_type_config['image_check_cells'],
            sheet_name=file_type_config['target_sheet']
        )

        return {
            "filename": str(file_path.name),
            "cell_values": cell_values,
            "image_results": image_results,
            "target_cells": file_type_config['target_cells'],
            "image_check_cells": file_type_config['image_check_cells'],
            "cell_labels": file_type_config['cell_labels']
        }

    def _format_multi_file_type_results(
        self,
        results: List[Dict[str, Any]],
        root_dir: Path,
        file_paths: List[Path]
    ) -> str:
        """複数ファイルタイプの結果を整形（新形式用）

        Args:
            results: 処理結果のリスト
            root_dir: ルートディレクトリ
            file_paths: 処理したファイルのパスリスト

        Returns:
            整形された結果文字列
        """
        if not results:
            return "処理対象のファイルが見つかりませんでした。"

        # 各ファイルごとに出力を生成
        output_lines = []

        for i, result in enumerate(results):
            file_path = file_paths[i]

            # 相対パスを取得
            try:
                relative_path = file_path.relative_to(root_dir)
            except ValueError:
                relative_path = file_path

            # ヘッダー行を生成
            header_cols = ["Filename", "Path"] + result["target_cells"]
            if result["image_check_cells"]:
                header_cols += [f"{cell}(画像)" for cell in result["image_check_cells"]]

            if not output_lines:
                # 最初のファイルの場合のみヘッダーを出力
                output_lines.append("\t".join(header_cols))

            # データ行を生成
            data_cols = [
                result["filename"],
                str(relative_path.parent)
            ] + [str(val) if val is not None else "" for val in result["cell_values"]]

            # 画像判定結果を追加
            if result["image_check_cells"]:
                data_cols += result["image_results"]

            output_lines.append("\t".join(data_cols))

        return "\n".join(output_lines)

    def _determine_file_type(self, filename: str) -> str:
        """ファイル名からファイルタイプを判定

        Args:
            filename: ファイル名

        Returns:
            "checklist", "record", または None
        """
        if "レビューチェックリスト" in filename or "checklist" in filename.lower():
            return "checklist"
        elif "レビュー記録表" in filename or "記録表" in filename:
            return "record"
        return None

    def _format_validation_results(
        self,
        validation_results: List[Dict[str, Any]],
        root_dir: Path
    ) -> str:
        """検証結果を詳細検証表形式で整形

        Args:
            validation_results: 検証結果のリスト
            root_dir: ルートディレクトリ

        Returns:
            整形された結果文字列
        """
        if not validation_results:
            return "処理対象のファイルが見つかりませんでした。"

        output_lines = []

        # ヘッダー行を生成
        headers = [
            "Filename", "Type", "Path", "プロジェクト名", "日付",
            "担当者", "承認者", "捺印", "ペア", "一致状況"
        ]
        output_lines.append("\t".join(headers))

        # 各ペアの結果を出力
        for result in validation_results:
            project_name = result["project_name"]
            checklist = result["checklist"]
            record = result["record"]
            validation = result["validation"]

            # チェックリスト行
            if checklist:
                checklist_row = self._format_file_row(
                    checklist, "チェックリスト", root_dir,
                    validation, project_name
                )
                output_lines.append(checklist_row)

            # 記録表行
            if record:
                record_row = self._format_file_row(
                    record, "記録表", root_dir,
                    validation, project_name
                )
                output_lines.append(record_row)

        # サマリを追加
        summary = self._generate_summary(validation_results)
        output_lines.append("")
        output_lines.append("【サマリ】")
        output_lines.extend(summary)

        return "\n".join(output_lines)

    def _format_file_row(
        self,
        file_data: Dict[str, Any],
        file_type: str,
        root_dir: Path,
        validation: Dict[str, Any],
        project_name: str
    ) -> str:
        """ファイルデータを行形式に整形

        Args:
            file_data: ファイルデータ
            file_type: ファイルタイプ
            root_dir: ルートディレクトリ
            validation: 検証結果
            project_name: プロジェクト名

        Returns:
            タブ区切りの行文字列
        """
        filename = file_data.get("filename", "")
        file_path = file_data.get("file_path", Path(""))
        cell_labels = file_data.get("cell_labels", [])
        cell_values = file_data.get("cell_values", [])

        # 相対パスを取得
        try:
            relative_path = file_path.relative_to(root_dir) if file_path else Path("")
            path_str = str(relative_path.parent) if relative_path.parent != Path(".") else "."
        except (ValueError, AttributeError):
            path_str = "."

        # ラベルから値を取得するヘルパー関数
        def get_value(label):
            if label in cell_labels:
                index = cell_labels.index(label)
                if index < len(cell_values):
                    return str(cell_values[index]) if cell_values[index] is not None else ""
            return ""

        # 各列の値を取得
        date_val = get_value("日付") if file_type == "チェックリスト" else get_value("承認日")
        reviewer_val = get_value("担当者") if file_type == "チェックリスト" else get_value("レビュアー")
        approver_val = get_value("承認者") if file_type == "チェックリスト" else ""

        # 捺印の状態
        has_stamp = validation.get("has_stamp")
        if file_type == "チェックリスト":
            stamp_str = "-"
        else:
            stamp_str = "○" if has_stamp else "×" if has_stamp is False else "-"

        # ペアの状態
        has_pair = validation.get("has_pair", False)
        pair_str = "あり" if has_pair else "なし"

        # 一致状況
        status = validation.get("status", "-")

        columns = [
            filename,
            file_type,
            path_str,
            project_name,
            date_val,
            reviewer_val,
            approver_val,
            stamp_str,
            pair_str,
            status
        ]

        return "\t".join(columns)

    def _generate_summary(self, validation_results: List[Dict[str, Any]]) -> List[str]:
        """サマリを生成

        Args:
            validation_results: 検証結果のリスト

        Returns:
            サマリ行のリスト
        """
        total_pairs = len(validation_results)
        complete_pairs = 0
        checklist_only = 0
        record_only = 0
        no_stamp = 0

        for result in validation_results:
            validation = result["validation"]
            checklist = result["checklist"]
            record = result["record"]

            if validation.get("has_pair"):
                complete_pairs += 1
            elif checklist and not record:
                checklist_only += 1
            elif record and not checklist:
                record_only += 1

            if validation.get("has_stamp") is False:
                no_stamp += 1

        summary = [
            f"- プロジェクト総数: {total_pairs}件",
            f"- 完全なペア: {complete_pairs}件",
            f"- チェックリストのみ: {checklist_only}件",
            f"- 記録表のみ: {record_only}件",
            f"- 捺印なし: {no_stamp}件"
        ]

        return summary

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
