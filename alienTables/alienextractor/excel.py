"""
alienextractor.excel
~~~~~~~~~~~~~~~~~~

Excel output generation with user-specified column filtering and ordering.
"""

from pathlib import Path
from typing import List, Dict, Optional

import pandas as pd

from .utils import resolve_column


def write_output(
    all_orders: List[Dict],
    output_path: Path,
    requested_columns: List[str],
) -> Path:
    """
    Write extracted orders to an Excel file.

    Only the columns requested by the user are included, in the exact
    order they were entered. If a requested column doesn't map to any
    internal extraction field, it appears with blank values.

    Args:
        all_orders: List of order dicts (keyed by internal COLUMNS names).
        output_path: Path for the output .xlsx file.
        requested_columns: List of column names as entered by the user.

    Returns:
        The path to the written Excel file.
    """
    if not all_orders:
        # Create an empty Excel with just the headers
        df = pd.DataFrame(columns=requested_columns)
        df.to_excel(str(output_path), index=False)
        return output_path

    # Build output rows with only the requested columns
    output_rows = []
    for order in all_orders:
        row = {}
        for user_col in requested_columns:
            internal_col = resolve_column(user_col)
            if internal_col and internal_col in order:
                row[user_col] = order[internal_col]
            else:
                # No matching rule — leave blank
                row[user_col] = ""
        output_rows.append(row)

    df = pd.DataFrame(output_rows, columns=requested_columns)
    df.to_excel(str(output_path), index=False)
    return output_path
