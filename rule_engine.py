"""Lightweight data-quality rule primitives for the Excel companion.

The rule layer is intentionally UI-independent. It detects issues and proposes
explicit fixes, while the Tkinter app remains responsible for presentation and
user confirmation.
"""

from dataclasses import dataclass
import re


WHITESPACE_RE = re.compile(r"[ \t\r\n\f\v]+")


@dataclass(frozen=True)
class RuleResult:
    """A detected data-quality issue in a workbook cell."""

    rule_id: str
    category: str
    severity: str
    sheet_name: str
    row_index: int
    column_index: int
    column_name: str
    value: object
    message: str
    fixable: bool = False


@dataclass(frozen=True)
class CellChange:
    """A before/after change that can be applied to a workbook cell."""

    sheet_name: str
    row_index: int
    column_index: int
    column_name: str
    before: str
    after: str


@dataclass(frozen=True)
class FixProposal:
    """A user-confirmed, column-scoped fix proposal."""

    rule_id: str
    label: str
    sheet_name: str
    column_index: int
    column_name: str
    changes: tuple[CellChange, ...]

    @property
    def affected_count(self):
        return len(self.changes)

    def preview_lines(self, limit=10):
        """Return representative before/after preview lines."""
        lines = []
        for change in self.changes[:limit]:
            lines.append(
                f'Row {change.row_index}: {change.before!r} → {change.after!r}'
            )
        if len(self.changes) > limit:
            lines.append(f"Showing {limit} of {len(self.changes)} changes")
        return lines

    def apply(self, workbook):
        """Apply the proposal to an in-memory openpyxl workbook."""
        sheet = workbook[self.sheet_name]
        for change in self.changes:
            sheet.cell(row=change.row_index, column=change.column_index).value = change.after


class WhitespaceCleansingRule:
    """Detect fixable whitespace that can break lookups and consolidation."""

    rule_id = "whitespace_cleanup"
    label = "Whitespace Cleanup"
    category = "cleansing"
    severity = "warning"

    def normalize(self, value):
        normalized = str(value).replace("\u00a0", " ")
        normalized = WHITESPACE_RE.sub(" ", normalized)
        return normalized.strip()

    def evaluate_cell(self, sheet_name, row_index, column_index, column_name, value):
        if not isinstance(value, str):
            return None

        cleaned = self.normalize(value)
        if cleaned == value:
            return None

        return RuleResult(
            rule_id=self.rule_id,
            category=self.category,
            severity=self.severity,
            sheet_name=sheet_name,
            row_index=row_index,
            column_index=column_index,
            column_name=column_name,
            value=value,
            message=(
                f"{column_name}: hidden or repeated whitespace may affect Excel "
                "lookups and consolidation."
            ),
            fixable=True,
        )

    def build_fix_proposals(self, results):
        grouped = {}
        for result in results:
            key = (result.sheet_name, result.column_index, result.column_name)
            grouped.setdefault(key, []).append(result)

        proposals = []
        for (sheet_name, column_index, column_name), grouped_results in grouped.items():
            changes = []
            for result in grouped_results:
                before = str(result.value)
                changes.append(
                    CellChange(
                        sheet_name=sheet_name,
                        row_index=result.row_index,
                        column_index=column_index,
                        column_name=column_name,
                        before=before,
                        after=self.normalize(before),
                    )
                )
            proposals.append(
                FixProposal(
                    rule_id=self.rule_id,
                    label=self.label,
                    sheet_name=sheet_name,
                    column_index=column_index,
                    column_name=column_name,
                    changes=tuple(changes),
                )
            )

        return proposals


class RuleRegistry:
    """Small registry for workbook-scoped data-quality rules."""

    def __init__(self, rules=None):
        self.rules = list(rules or [WhitespaceCleansingRule()])

    def scan_sheet(self, sheet, sheet_name, headers, start_row):
        results = []
        for row in sheet.iter_rows(min_row=start_row, max_row=sheet.max_row, max_col=len(headers)):
            for column_index, cell in enumerate(row, start=1):
                column_name = headers[column_index - 1]
                for rule in self.rules:
                    result = rule.evaluate_cell(
                        sheet_name=sheet_name,
                        row_index=cell.row,
                        column_index=column_index,
                        column_name=column_name,
                        value=cell.value,
                    )
                    if result:
                        results.append(result)
        return results

    def build_fix_proposals(self, results):
        proposals = []
        for rule in self.rules:
            rule_results = [result for result in results if result.rule_id == rule.rule_id and result.fixable]
            proposals.extend(rule.build_fix_proposals(rule_results))
        return proposals
