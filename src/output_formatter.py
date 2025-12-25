"""出力整形モジュール

ファイル抽出結果を視認性の高いフォーマットで整形する。
"""
from typing import List, Dict, Any
import unicodedata


class OutputFormatter:
    """出力整形クラス

    Attributes:
        target_cells (List[str]): 抽出対象セルのリスト
        image_check_cells (List[str]): 画像判定対象セルのリスト
    """

    def __init__(self, target_cells: List[str], image_check_cells: List[str]):
        """出力整形の初期化

        Args:
            target_cells: 抽出対象セルのリスト（例: ["A1", "B1", "C1"]）
            image_check_cells: 画像判定対象セルのリスト（例: ["D1", "E1"]）
        """
        self.target_cells = target_cells
        self.image_check_cells = image_check_cells

    def format_results(self, results: List[Dict[str, Any]]) -> str:
        """抽出結果を整形して文字列で返す

        Args:
            results: 各ファイルの抽出結果のリスト
                形式: [{"filename": str, "cell_values": List[Any], "image_results": List[str]}, ...]

        Returns:
            整形された結果文字列（カンマ位置が垂直に揃う）
        """
        # 空の場合はヘッダーのみ返す
        if not results:
            return self._generate_header() + "\n"

        # 各列の最大幅を計算
        column_widths = self._calculate_column_widths(results)

        # ヘッダー行を生成
        output_lines = [self._generate_header_with_padding(column_widths)]

        # 各結果行を整形
        for result in results:
            formatted_row = self._format_row_with_padding(result, column_widths)
            output_lines.append(formatted_row)

        return "\n".join(output_lines) + "\n"

    def _generate_header(self) -> str:
        """ヘッダー行を生成（パディングなし）

        Returns:
            ヘッダー文字列（例: "Filename, A1, B1, C1, D1(画像), E1(画像)"）
        """
        columns = ["Filename"]

        # target_cellsを追加
        columns.extend(self.target_cells)

        # image_check_cellsを追加（"セル名(画像)"形式）
        for cell in self.image_check_cells:
            columns.append(f"{cell}(画像)")

        return ", ".join(columns)

    def _generate_header_with_padding(self, column_widths: List[int]) -> str:
        """ヘッダー行をパディング付きで生成

        Args:
            column_widths: 各列の最大幅のリスト

        Returns:
            パディング済みヘッダー文字列
        """
        columns = ["Filename"]
        columns.extend(self.target_cells)

        for cell in self.image_check_cells:
            columns.append(f"{cell}(画像)")

        # 各列をパディング
        padded_columns = []
        for i, col in enumerate(columns):
            width = column_widths[i]
            padded = self._pad_string(col, width)
            padded_columns.append(padded)

        return ", ".join(padded_columns)

    def _format_row(self, result: Dict[str, Any]) -> str:
        """単一の結果行を整形（パディングなし）

        Args:
            result: ファイルの抽出結果
                形式: {"filename": str, "cell_values": List[Any], "image_results": List[str]}

        Returns:
            整形された行文字列
        """
        columns = [result["filename"]]

        # セル値を追加（文字列に変換）
        for value in result["cell_values"]:
            columns.append(str(value))

        # 画像判定結果を追加
        columns.extend(result["image_results"])

        return ", ".join(columns)

    def _format_row_with_padding(
        self, result: Dict[str, Any], column_widths: List[int]
    ) -> str:
        """単一の結果行をパディング付きで整形

        Args:
            result: ファイルの抽出結果
            column_widths: 各列の最大幅のリスト

        Returns:
            パディング済み行文字列
        """
        columns = [result["filename"]]

        # セル値を追加（文字列に変換）
        for value in result["cell_values"]:
            columns.append(str(value))

        # 画像判定結果を追加
        columns.extend(result["image_results"])

        # 各列をパディング
        padded_columns = []
        for i, col in enumerate(columns):
            width = column_widths[i]
            padded = self._pad_string(col, width)
            padded_columns.append(padded)

        return ", ".join(padded_columns)

    def _calculate_column_widths(self, results: List[Dict[str, Any]]) -> List[int]:
        """各列の最大幅を計算

        Args:
            results: 各ファイルの抽出結果のリスト

        Returns:
            各列の最大幅のリスト（日本語文字も考慮）
        """
        # ヘッダー行の列名を取得
        header_columns = ["Filename"]
        header_columns.extend(self.target_cells)
        for cell in self.image_check_cells:
            header_columns.append(f"{cell}(画像)")

        # 各列の最大幅を初期化（ヘッダーの幅）
        column_widths = [self._display_width(col) for col in header_columns]

        # 各結果行の幅をチェック
        for result in results:
            # ファイル名の幅
            filename_width = self._display_width(result["filename"])
            column_widths[0] = max(column_widths[0], filename_width)

            # セル値の幅
            for i, value in enumerate(result["cell_values"]):
                value_width = self._display_width(str(value))
                column_widths[i + 1] = max(column_widths[i + 1], value_width)

            # 画像判定結果の幅
            offset = 1 + len(result["cell_values"])
            for i, image_result in enumerate(result["image_results"]):
                result_width = self._display_width(image_result)
                column_widths[offset + i] = max(
                    column_widths[offset + i], result_width
                )

        return column_widths

    def _display_width(self, text: str) -> int:
        """文字列の表示幅を計算（日本語文字は幅2、半角文字は幅1）

        Args:
            text: 計算対象の文字列

        Returns:
            表示幅（半角文字換算）
        """
        width = 0
        for char in text:
            # East Asian Width プロパティを取得
            east_asian_width = unicodedata.east_asian_width(char)

            # 全角文字（F, W）は幅2、それ以外は幅1
            if east_asian_width in ("F", "W"):
                width += 2
            else:
                width += 1

        return width

    def _pad_string(self, text: str, target_width: int) -> str:
        """文字列を指定幅までスペースでパディング

        Args:
            text: パディング対象の文字列
            target_width: 目標幅（半角文字換算）

        Returns:
            パディング済み文字列
        """
        current_width = self._display_width(text)
        padding_needed = target_width - current_width

        # パディングが負の場合は0にする
        if padding_needed < 0:
            padding_needed = 0

        return text + " " * padding_needed
