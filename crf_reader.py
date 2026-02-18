from __future__ import annotations

import json

from openpyxl.worksheet.worksheet import Worksheet

from models import Crf, IdConfig, IdConfigField


class CrfReader:
    @staticmethod
    def read_crfs_worksheet(worksheet: Worksheet) -> list[Crf]:
        crfs: list[Crf] = []
        for row_idx in range(2, worksheet.max_row + 1):
            crf = Crf(
                display_order=_nullable_int(_cell_trim(worksheet, row_idx, 1)),
                tablename=_null_if_empty(_cell_trim(worksheet, row_idx, 2)),
                displayname=_null_if_empty(_cell_trim(worksheet, row_idx, 3)),
                primarykey=_null_if_empty(_cell_trim(worksheet, row_idx, 4)),
                isbase=_nullable_int(_cell_trim(worksheet, row_idx, 6)),
                linkingfield=_null_if_empty(_cell_trim(worksheet, row_idx, 7)),
                parenttable=_null_if_empty(_cell_trim(worksheet, row_idx, 8)),
                incrementfield=_null_if_empty(_cell_trim(worksheet, row_idx, 9)),
                requireslink=_nullable_int(_cell_trim(worksheet, row_idx, 10)),
                repeat_count_field=_null_if_empty(_cell_trim(worksheet, row_idx, 11)),
                auto_start_repeat=_nullable_int(_cell_trim(worksheet, row_idx, 12)),
                repeat_enforce_count=_nullable_int(_cell_trim(worksheet, row_idx, 13)),
                display_fields=_null_if_empty(_cell_trim(worksheet, row_idx, 14)),
                entry_condition=_null_if_empty(_cell_trim(worksheet, row_idx, 15)),
            )
            idconfig_json = _cell_trim(worksheet, row_idx, 5)
            if idconfig_json:
                try:
                    raw = json.loads(idconfig_json)
                    fields = raw.get("fields")
                    parsed_fields = None
                    if isinstance(fields, list):
                        parsed_fields = [
                            IdConfigField(name=f.get("name", ""), length=int(f.get("length", 0))) for f in fields
                        ]
                    crf.idconfig = IdConfig(
                        prefix=raw.get("prefix"),
                        fields=parsed_fields,
                        incrementLength=raw.get("incrementLength"),
                    )
                except Exception:
                    crf.idconfig = None
            crfs.append(crf)
        return crfs


def _cell_trim(worksheet: Worksheet, row: int, col: int) -> str:
    value = worksheet.cell(row=row, column=col).value
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def _nullable_int(value: str) -> int | None:
    try:
        return int(value)
    except Exception:
        return None


def _null_if_empty(value: str) -> str | None:
    return value if value else None
