"""セル値抽出モジュール

ExcelファイルおよびCSVファイルから指定されたセルの値を抽出する。
"""
import pandas as pd
import openpyxl
from pathlib import Path
from typing import List, Any
import re


class CellExtractor:
    """セル値抽出クラス

    Attributes:
        file_path (Path): 対象ファイルのパス
        file_extension (str): ファイルの拡張子
    """

    # サポートする拡張子
    SUPPORTED_EXTENSIONS = {'.xlsx', '.xls', '.csv'}

    def __init__(self, file_path: Path):
        """セル値抽出の初期化

        Args:
            file_path: 対象ファイルのパス

        Raises:
            FileNotFoundError: ファイルが存在しない場合
            ValueError: サポートされていないファイル形式の場合
        """
        if not file_path.exists():
            raise FileNotFoundError(
                f"エラー: ファイルが見つかりません: {file_path}"
            )

        self.file_path = file_path
        self.file_extension = file_path.suffix.lower()

        if self.file_extension not in self.SUPPORTED_EXTENSIONS:
            raise ValueError(
                f"エラー: サポートされていないファイル形式です: {self.file_extension}"
            )

    def extract_cells(self, cell_list: List[str]) -> List[Any]:
        """指定されたセルの値を抽出

        Args:
            cell_list: セル座標のリスト（例: ["A1", "B2", "C3"]）

        Returns:
            抽出されたセル値のリスト
        """
        if self.file_extension == '.csv':
            return self._extract_from_csv(cell_list)
        else:
            return self._extract_from_excel(cell_list)

    def _extract_from_excel(self, cell_list: List[str]) -> List[Any]:
        """Excelファイルからセル値を抽出

        Args:
            cell_list: セル座標のリスト

        Returns:
            抽出されたセル値のリスト
        """
        # .xls形式の場合はpandasで読み込む（openpyxlは.xlsをサポートしない）
        if self.file_extension == '.xls':
            return self._extract_from_excel_with_pandas(cell_list)

        try:
            # openpyxlでExcelファイルを開く（.xlsx形式）
            workbook = openpyxl.load_workbook(self.file_path, data_only=True)
            sheet = workbook.active

            values = []
            for cell_address in cell_list:
                try:
                    cell_value = sheet[cell_address].value
                    # Noneの場合は空文字列に変換
                    values.append(cell_value if cell_value is not None else "")
                except Exception:
                    # セル範囲外などのエラーの場合は空文字列
                    values.append("")

            workbook.close()
            return values

        except Exception as e:
            raise RuntimeError(
                f"エラー: Excelファイルの読み込みに失敗しました: {self.file_path} - {str(e)}"
            )

    def _extract_from_excel_with_pandas(self, cell_list: List[str]) -> List[Any]:
        """Excelファイル(.xls)からpandasでセル値を抽出

        Args:
            cell_list: セル座標のリスト

        Returns:
            抽出されたセル値のリスト
        """
        try:
            # pandasでExcelファイルを読み込む（ヘッダーなし）
            df = pd.read_excel(self.file_path, header=None)

            values = []
            for cell_address in cell_list:
                row, col = self._parse_cell_address(cell_address)
                try:
                    # pandasのDataFrameから値を取得（0-indexed）
                    cell_value = df.iloc[row, col]
                    # NaNの場合は空文字列に変換
                    if pd.isna(cell_value):
                        values.append("")
                    else:
                        values.append(cell_value)
                except (IndexError, KeyError):
                    # 範囲外の場合は空文字列
                    values.append("")

            return values

        except Exception as e:
            raise RuntimeError(
                f"エラー: Excelファイルの読み込みに失敗しました: {self.file_path} - {str(e)}"
            )

    def _extract_from_csv(self, cell_list: List[str]) -> List[Any]:
        """CSVファイルからセル値を抽出

        Args:
            cell_list: セル座標のリスト

        Returns:
            抽出されたセル値のリスト
        """
        try:
            # pandasでCSVファイルを読み込む（ヘッダーなし）
            df = pd.read_csv(self.file_path, header=None)

            values = []
            for cell_address in cell_list:
                row, col = self._parse_cell_address(cell_address)
                try:
                    # pandasのDataFrameから値を取得（0-indexed）
                    cell_value = df.iloc[row, col]
                    # NaNの場合は空文字列に変換
                    if pd.isna(cell_value):
                        values.append("")
                    else:
                        values.append(cell_value)
                except (IndexError, KeyError):
                    # 範囲外の場合は空文字列
                    values.append("")

            return values

        except Exception as e:
            raise RuntimeError(
                f"エラー: CSVファイルの読み込みに失敗しました: {self.file_path} - {str(e)}"
            )

    def _parse_cell_address(self, cell_address: str) -> tuple:
        """セル座標（例: A1, B2）を行列インデックスに変換

        Args:
            cell_address: セル座標（例: "A1", "B2"）

        Returns:
            (row_index, col_index) のタプル（0-indexed）
        """
        # セル座標を正規表現で分解（例: "A1" -> "A", "1"）
        match = re.match(r'([A-Z]+)(\d+)', cell_address.upper())
        if not match:
            raise ValueError(f"不正なセル座標です: {cell_address}")

        col_letters, row_num = match.groups()

        # 列文字を数値に変換（A=0, B=1, ..., Z=25, AA=26, ...）
        col_index = 0
        for char in col_letters:
            col_index = col_index * 26 + (ord(char) - ord('A') + 1)
        col_index -= 1  # 0-indexedに変換

        # 行番号を0-indexedに変換
        row_index = int(row_num) - 1

        return (row_index, col_index)
