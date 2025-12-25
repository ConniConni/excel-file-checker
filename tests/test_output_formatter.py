"""出力整形機能のテスト"""
import pytest
from pathlib import Path
from src.output_formatter import OutputFormatter


class TestOutputFormatter:
    """OutputFormatterクラスのテスト"""

    def test_ヘッダー行を生成できる(self):
        """ヘッダー行（列名）を正しく生成できることを確認"""
        formatter = OutputFormatter(
            target_cells=["A1", "B1", "C1"],
            image_check_cells=["D1", "E1"]
        )

        header = formatter._generate_header()

        assert "Filename" in header
        assert "A1" in header
        assert "B1" in header
        assert "C1" in header
        assert "D1(画像)" in header
        assert "E1(画像)" in header

    def test_単一行のデータを整形できる(self):
        """単一のファイル結果を整形できることを確認"""
        formatter = OutputFormatter(
            target_cells=["A1", "B1"],
            image_check_cells=["D1"]
        )

        result = {
            "filename": "test.xlsx",
            "cell_values": ["2023/12/25", "33000.5"],
            "image_results": ["○"]
        }

        formatted = formatter._format_row(result)

        assert "test.xlsx" in formatted
        assert "2023/12/25" in formatted
        assert "33000.5" in formatted
        assert "○" in formatted
        assert "," in formatted

    def test_複数行のデータを整形できる(self):
        """複数ファイルの結果を整形できることを確認"""
        formatter = OutputFormatter(
            target_cells=["A1", "B1"],
            image_check_cells=["D1"]
        )

        results = [
            {
                "filename": "file1.xlsx",
                "cell_values": ["2023/12/25", "33000.5"],
                "image_results": ["○"]
            },
            {
                "filename": "file2.csv",
                "cell_values": ["2023/12/26", "33100.2"],
                "image_results": ["-"]
            }
        ]

        formatted = formatter.format_results(results)

        assert "file1.xlsx" in formatted
        assert "file2.csv" in formatted
        assert "2023/12/25" in formatted
        assert "2023/12/26" in formatted

    def test_カンマが垂直に揃う(self):
        """各行のカンマ位置が垂直に揃うことを確認"""
        formatter = OutputFormatter(
            target_cells=["A1", "B1"],
            image_check_cells=["D1"]
        )

        results = [
            {
                "filename": "short.xlsx",
                "cell_values": ["A", "B"],
                "image_results": ["○"]
            },
            {
                "filename": "very_long_filename.csv",
                "cell_values": ["ABC", "DEF"],
                "image_results": ["-"]
            }
        ]

        formatted = formatter.format_results(results)
        lines = formatted.strip().split("\n")

        # 各行の最初のカンマの位置を取得
        comma_positions = [line.index(",") for line in lines if "," in line]

        # 全ての行で最初のカンマの位置が同じであることを確認
        assert len(set(comma_positions)) == 1

    def test_日本語文字を含むデータを正しく整形できる(self):
        """日本語文字を含むデータを正しくパディングできることを確認"""
        formatter = OutputFormatter(
            target_cells=["A1", "B1"],
            image_check_cells=["D1"]
        )

        results = [
            {
                "filename": "test.xlsx",
                "cell_values": ["確定", "暫定"],
                "image_results": ["○"]
            }
        ]

        formatted = formatter.format_results(results)

        assert "確定" in formatted
        assert "暫定" in formatted
        # カンマが含まれることを確認
        assert "," in formatted

    def test_空のリストの場合はヘッダーのみ返す(self):
        """空の結果リストの場合、ヘッダー行のみを返すことを確認"""
        formatter = OutputFormatter(
            target_cells=["A1", "B1"],
            image_check_cells=["D1"]
        )

        formatted = formatter.format_results([])
        lines = formatted.strip().split("\n")

        # ヘッダー行のみが存在
        assert len(lines) == 1
        assert "Filename" in lines[0]

    def test_画像チェックセルが空の場合も動作する(self):
        """画像チェックセルが指定されていない場合も正しく動作することを確認"""
        formatter = OutputFormatter(
            target_cells=["A1", "B1"],
            image_check_cells=[]
        )

        results = [
            {
                "filename": "test.xlsx",
                "cell_values": ["2023/12/25", "33000.5"],
                "image_results": []
            }
        ]

        formatted = formatter.format_results(results)

        assert "test.xlsx" in formatted
        assert "2023/12/25" in formatted
        assert "D1(画像)" not in formatted

    def test_出力がUTF8文字列である(self):
        """出力がUTF-8文字列であることを確認"""
        formatter = OutputFormatter(
            target_cells=["A1"],
            image_check_cells=["D1"]
        )

        results = [
            {
                "filename": "test.xlsx",
                "cell_values": ["確定"],
                "image_results": ["○"]
            }
        ]

        formatted = formatter.format_results(results)

        # strであることを確認
        assert isinstance(formatted, str)
        # 日本語が含まれることを確認
        assert "確定" in formatted
        assert "○" in formatted

    def test_全列のパディングが正しく計算される(self):
        """全ての列で正しくパディングが計算されることを確認"""
        formatter = OutputFormatter(
            target_cells=["A1", "B1"],
            image_check_cells=["D1"]
        )

        results = [
            {
                "filename": "short.xlsx",
                "cell_values": ["A", "B"],
                "image_results": ["○"]
            },
            {
                "filename": "very_long_name.csv",
                "cell_values": ["ABC", "DEF"],
                "image_results": ["-"]
            }
        ]

        formatted = formatter.format_results(results)
        lines = formatted.strip().split("\n")

        # 全ての行が同じ構造を持つことを確認（カンマの数が同じ）
        comma_counts = [line.count(",") for line in lines]
        assert len(set(comma_counts)) == 1

    def test_サマリ部分が正しく生成される(self):
        """サマリ部分が正しいフォーマットで生成されることを確認"""
        formatter = OutputFormatter(
            target_cells=["A1", "B1"],
            image_check_cells=["D1"]
        )

        results = [
            {
                "filename": "test1.xlsx",
                "cell_values": ["A", "B"],
                "image_results": ["○"]
            },
            {
                "filename": "test2.xlsx",
                "cell_values": ["C", "D"],
                "image_results": ["×"]
            }
        ]

        root_dir = Path("/data")
        file_paths = [
            Path("/data/test1.xlsx"),
            Path("/data/test2.xlsx")
        ]

        formatted = formatter.format_results(results, root_dir=root_dir, file_paths=file_paths)

        # サマリセクションが含まれることを確認
        assert "=== サマリ ===" in formatted
        assert "=============" in formatted
        assert "出力対象ファイル件数: 2件" in formatted

    def test_サマリのツリー構造が正しく生成される(self):
        """サマリのツリー構造が正しく生成されることを確認"""
        formatter = OutputFormatter(
            target_cells=["A1"],
            image_check_cells=[]
        )

        results = [
            {"filename": "file1.xlsx", "cell_values": ["A"], "image_results": []},
            {"filename": "file2.xlsx", "cell_values": ["B"], "image_results": []},
            {"filename": "file3.xlsx", "cell_values": ["C"], "image_results": []}
        ]

        root_dir = Path("/data")
        file_paths = [
            Path("/data/file1.xlsx"),
            Path("/data/subfolder/file2.xlsx"),
            Path("/data/file3.xlsx")
        ]

        formatted = formatter.format_results(results, root_dir=root_dir, file_paths=file_paths)

        # ツリー構造の要素が含まれることを確認
        assert "/data/" in formatted or "/data" in formatted
        assert "file1.xlsx" in formatted
        assert "file2.xlsx" in formatted
        assert "file3.xlsx" in formatted
        assert "subfolder" in formatted
        # ツリー記号が含まれることを確認
        assert "├──" in formatted or "└──" in formatted

    def test_サマリのファイルがアルファベット順にソートされる(self):
        """サマリのファイルがアルファベット順にソートされることを確認"""
        formatter = OutputFormatter(
            target_cells=["A1"],
            image_check_cells=[]
        )

        results = [
            {"filename": "c.xlsx", "cell_values": ["A"], "image_results": []},
            {"filename": "a.xlsx", "cell_values": ["B"], "image_results": []},
            {"filename": "b.xlsx", "cell_values": ["C"], "image_results": []}
        ]

        root_dir = Path("/data")
        file_paths = [
            Path("/data/c.xlsx"),
            Path("/data/a.xlsx"),
            Path("/data/b.xlsx")
        ]

        formatted = formatter.format_results(results, root_dir=root_dir, file_paths=file_paths)

        # サマリ部分を抽出
        summary_start = formatted.index("=== サマリ ===")
        summary_section = formatted[summary_start:]

        # a.xlsx, b.xlsx, c.xlsx の順に出現することを確認
        a_pos = summary_section.index("a.xlsx")
        b_pos = summary_section.index("b.xlsx")
        c_pos = summary_section.index("c.xlsx")

        assert a_pos < b_pos < c_pos

    def test_サマリなしでも既存機能が動作する(self):
        """root_dirとfile_pathsを指定しない場合、サマリなしで既存機能が動作することを確認"""
        formatter = OutputFormatter(
            target_cells=["A1", "B1"],
            image_check_cells=["D1"]
        )

        results = [
            {
                "filename": "test.xlsx",
                "cell_values": ["A", "B"],
                "image_results": ["○"]
            }
        ]

        # サマリ情報を渡さない
        formatted = formatter.format_results(results)

        # 通常の出力は含まれる
        assert "Filename" in formatted
        assert "test.xlsx" in formatted

        # サマリは含まれない
        assert "=== サマリ ===" not in formatted
