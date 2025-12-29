"""レビュー文書のペアリングと検証を行うサービス"""
from typing import List, Dict, Any, Optional
from pathlib import Path


class ReviewPair:
    """レビューチェックリストとレビュー記録表のペア"""

    def __init__(self, project_name: str):
        """
        Args:
            project_name: プロジェクト名
        """
        self.project_name = project_name
        self.checklist: Optional[Dict[str, Any]] = None
        self.record: Optional[Dict[str, Any]] = None

    def set_checklist(self, data: Dict[str, Any]):
        """チェックリストを設定"""
        self.checklist = data

    def set_record(self, data: Dict[str, Any]):
        """記録表を設定"""
        self.record = data

    def has_both(self) -> bool:
        """両方のファイルが揃っているか"""
        return self.checklist is not None and self.record is not None

    def has_checklist_only(self) -> bool:
        """チェックリストのみか"""
        return self.checklist is not None and self.record is None

    def has_record_only(self) -> bool:
        """記録表のみか"""
        return self.checklist is None and self.record is not None

    def validate(self) -> Dict[str, Any]:
        """ペアの検証を実行

        Returns:
            検証結果の辞書
            {
                "has_pair": bool,
                "date_match": bool or None,
                "reviewer_match": bool or None,
                "has_stamp": bool or None,
                "status": str,
                "warnings": List[str]
            }
        """
        warnings = []

        if not self.has_both():
            if self.has_checklist_only():
                return {
                    "has_pair": False,
                    "date_match": None,
                    "reviewer_match": None,
                    "has_stamp": None,
                    "status": "⚠️ 記録表なし",
                    "warnings": ["レビュー記録表が見つかりません"]
                }
            elif self.has_record_only():
                # 記録表の捺印チェック
                has_stamp = self._check_stamp(self.record)
                status = "⚠️ チェックリストなし"
                if not has_stamp:
                    warnings.append("捺印がありません")
                    status = "⚠️ チェックリストなし・捺印なし"

                return {
                    "has_pair": False,
                    "date_match": None,
                    "reviewer_match": None,
                    "has_stamp": has_stamp,
                    "status": status,
                    "warnings": ["レビューチェックリストが見つかりません"] + warnings
                }

        # 両方ある場合の検証
        date_match = self._validate_date()
        reviewer_match = self._validate_reviewer()
        has_stamp = self._check_stamp(self.record)

        # ステータス判定
        if date_match and reviewer_match and has_stamp:
            status = "✓ OK"
        else:
            status = "⚠️ 不一致あり"
            if not date_match:
                warnings.append("日付が一致しません")
            if not reviewer_match:
                warnings.append("担当者/レビュアーが一致しません")
            if not has_stamp:
                warnings.append("捺印がありません")

        return {
            "has_pair": True,
            "date_match": date_match,
            "reviewer_match": reviewer_match,
            "has_stamp": has_stamp,
            "status": status,
            "warnings": warnings
        }

    def _validate_date(self) -> bool:
        """日付の一致確認

        チェックリストの「日付」と記録表の「承認日」を比較
        """
        if not self.has_both():
            return False

        # チェックリストから日付を取得（cell_labelsの「日付」に対応する値）
        checklist_date = self._get_value_by_label(self.checklist, "日付")
        # 記録表から承認日を取得（cell_labelsの「承認日」に対応する値）
        record_date = self._get_value_by_label(self.record, "承認日")

        if checklist_date is None or record_date is None:
            return False

        # 承認日から日付部分を抽出（"承認日: 2025-01-15" → "2025-01-15"）
        record_date_str = str(record_date)
        if ":" in record_date_str:
            record_date_str = record_date_str.split(":", 1)[1].strip()

        return str(checklist_date) == record_date_str

    def _validate_reviewer(self) -> bool:
        """担当者/レビュアーの一致確認

        チェックリストの「担当者」と記録表の「レビュアー」を比較
        """
        if not self.has_both():
            return False

        checklist_reviewer = self._get_value_by_label(self.checklist, "担当者")
        record_reviewer = self._get_value_by_label(self.record, "レビュアー")

        if checklist_reviewer is None or record_reviewer is None:
            return False

        # レビュアーから名前部分を抽出（"レビュアー: 山田太郎" → "山田太郎"）
        record_reviewer_str = str(record_reviewer)
        if ":" in record_reviewer_str:
            record_reviewer_str = record_reviewer_str.split(":", 1)[1].strip()

        return str(checklist_reviewer) == record_reviewer_str

    def _check_stamp(self, record: Optional[Dict[str, Any]]) -> bool:
        """捺印の有無確認"""
        if record is None:
            return False

        image_results = record.get("image_results", [])
        if not image_results:
            return False

        # 画像結果に「○」が含まれているかチェック
        return "○" in image_results

    def _get_value_by_label(self, data: Dict[str, Any], label: str) -> Any:
        """ラベル名から値を取得

        Args:
            data: ファイルデータ
            label: ラベル名（例: "日付", "プロジェクト名"）

        Returns:
            対応する値、見つからない場合はNone
        """
        cell_labels = data.get("cell_labels", [])
        cell_values = data.get("cell_values", [])

        if label not in cell_labels:
            return None

        index = cell_labels.index(label)
        if index >= len(cell_values):
            return None

        return cell_values[index]


class ReviewValidator:
    """レビュー文書の検証サービス"""

    def __init__(self):
        """初期化"""
        self.pairs: Dict[str, ReviewPair] = {}

    def add_file(self, file_data: Dict[str, Any], file_type: str):
        """ファイルを追加してペアリング

        Args:
            file_data: ファイルデータ
                {
                    "filename": str,
                    "cell_values": List[Any],
                    "cell_labels": List[str],
                    "image_results": List[str],
                    "file_path": Path,
                    "file_type": str
                }
            file_type: ファイルタイプ（"checklist" or "record"）
        """
        # プロジェクト名を取得
        project_name = self._extract_project_name(file_data)

        # ペアを取得または作成
        if project_name not in self.pairs:
            self.pairs[project_name] = ReviewPair(project_name)

        pair = self.pairs[project_name]

        # ファイルタイプに応じて設定
        if file_type == "checklist":
            pair.set_checklist(file_data)
        elif file_type == "record":
            pair.set_record(file_data)

    def validate_all(self) -> List[Dict[str, Any]]:
        """全ペアの検証を実行

        Returns:
            検証結果のリスト
        """
        results = []

        for project_name in sorted(self.pairs.keys()):
            pair = self.pairs[project_name]
            validation = pair.validate()

            result = {
                "project_name": project_name,
                "checklist": pair.checklist,
                "record": pair.record,
                "validation": validation
            }
            results.append(result)

        return results

    def _extract_project_name(self, file_data: Dict[str, Any]) -> str:
        """ファイルデータからプロジェクト名を抽出

        Args:
            file_data: ファイルデータ

        Returns:
            プロジェクト名
        """
        cell_labels = file_data.get("cell_labels", [])
        cell_values = file_data.get("cell_values", [])

        # "プロジェクト名"ラベルに対応する値を取得
        if "プロジェクト名" in cell_labels:
            index = cell_labels.index("プロジェクト名")
            if index < len(cell_values):
                return str(cell_values[index])

        # 見つからない場合はファイル名から推測
        filename = file_data.get("filename", "Unknown")
        return f"Unknown_{filename}"
