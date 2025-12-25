"""画像判定モジュール

Excelファイルの指定セルに画像が存在するかを判定する。
"""
import openpyxl
from pathlib import Path
from typing import List
import re


class ImageChecker:
    """画像判定クラス

    Attributes:
        file_path (Path): 対象ファイルのパス
        file_extension (str): ファイルの拡張子
    """

    # 画像判定をサポートする拡張子（.xlsxのみ）
    SUPPORTED_EXTENSIONS = {'.xlsx'}

    def __init__(self, file_path: Path):
        """画像判定の初期化

        Args:
            file_path: 対象ファイルのパス

        Raises:
            FileNotFoundError: ファイルが存在しない場合
        """
        if not file_path.exists():
            raise FileNotFoundError(
                f"エラー: ファイルが見つかりません: {file_path}"
            )

        self.file_path = file_path
        self.file_extension = file_path.suffix.lower()

    def check_images(self, cell_list: List[str], sheet_name: str = None) -> List[str]:
        """指定されたセルに画像が存在するかを判定

        Args:
            cell_list: セル座標のリスト（例: ["D1", "E1"]）
            sheet_name: シート名（Excelファイルのみ有効、指定しない場合は最初のシート）

        Returns:
            判定結果のリスト
            - "○": 画像あり
            - "×": 画像なし
            - "-": 画像判定非対応（CSV、.xls形式）
        """
        # 空のセルリストの場合は空リストを返す
        if not cell_list:
            return []

        # .xlsx形式以外は画像判定非対応
        if self.file_extension not in self.SUPPORTED_EXTENSIONS:
            return ["-"] * len(cell_list)

        return self._check_images_in_excel(cell_list, sheet_name)

    def _check_images_in_excel(self, cell_list: List[str], sheet_name: str = None) -> List[str]:
        """Excelファイル内の画像を判定

        Args:
            cell_list: セル座標のリスト
            sheet_name: シート名（指定しない場合は最初のシート）

        Returns:
            判定結果のリスト（"○" or "×"）
        """
        try:
            # openpyxlでExcelファイルを開く
            workbook = openpyxl.load_workbook(self.file_path)

            # シート名が指定されている場合はそのシートを、指定がない場合は最初のシートを使用
            if sheet_name:
                sheet = workbook[sheet_name]
            else:
                sheet = workbook.worksheets[0]

            # シート内の全画像を取得
            images = sheet._images

            # 各セルについて画像の有無を判定
            results = []
            for cell_address in cell_list:
                has_image = self._has_image_at_cell(images, cell_address)
                results.append("○" if has_image else "×")

            workbook.close()
            return results

        except Exception as e:
            # エラーが発生した場合は全て"×"を返す
            return ["×"] * len(cell_list)

    def _has_image_at_cell(self, images, cell_address: str) -> bool:
        """指定セルに画像が存在するかをチェック

        Args:
            images: シート内の画像リスト
            cell_address: セル座標（例: "D1"）

        Returns:
            画像が存在する場合True、存在しない場合False
        """
        # 画像がない場合はFalse
        if not images:
            return False

        # セル座標を行列番号に変換
        row, col = self._parse_cell_address(cell_address)

        # 各画像のアンカーポイントをチェック
        for image in images:
            # 画像のアンカー位置を取得
            anchor = image.anchor

            # アンカーがOneCellAnchorまたはTwoCellAnchorの場合
            if hasattr(anchor, '_from'):
                # アンカーの開始セル位置を取得
                anchor_col = anchor._from.col
                anchor_row = anchor._from.row

                # セル位置が一致するかチェック
                if anchor_col == col and anchor_row == row:
                    return True

        return False

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
