"""
merge_lca.py
------------
Reads two LCA (Life Cycle Assessment) Excel files and merges them into a
single combined Excel output file.

Usage
-----
    python merge_lca.py [file1] [file2] [output]

Arguments (all optional, defaults shown below):
    file1   – path to the first LCA Excel file   (default: lca_file1.xlsx)
    file2   – path to the second LCA Excel file  (default: lca_file2.xlsx)
    output  – path for the merged output file    (default: lca_merged.xlsx)

Examples
--------
    # Use default file names
    python merge_lca.py

    # Specify custom paths
    python merge_lca.py data/lca_a.xlsx data/lca_b.xlsx results/merged.xlsx
"""

import sys
import pandas as pd


DEFAULT_FILE1 = "lca_file1.xlsx"
DEFAULT_FILE2 = "lca_file2.xlsx"
DEFAULT_OUTPUT = "lca_merged.xlsx"


def read_lca_file(path: str) -> pd.DataFrame:
    """Read a single LCA Excel file and return it as a DataFrame."""
    print(f"Reading: {path}")
    try:
        df = pd.read_excel(path)
    except FileNotFoundError:
        raise SystemExit(f"Error: file not found – {path}")
    except Exception as exc:
        raise SystemExit(f"Error: could not read '{path}': {exc}") from exc
    print(f"  -> {len(df)} rows, {len(df.columns)} columns")
    return df


def _validate_columns(df1: pd.DataFrame, df2: pd.DataFrame) -> None:
    """Warn when the two DataFrames have different column sets."""
    cols1 = set(df1.columns)
    cols2 = set(df2.columns)
    only_in_1 = cols1 - cols2
    only_in_2 = cols2 - cols1
    if only_in_1:
        print(f"Warning: columns only in file 1 (will be NaN for file 2 rows): {sorted(only_in_1)}")
    if only_in_2:
        print(f"Warning: columns only in file 2 (will be NaN for file 1 rows): {sorted(only_in_2)}")


def merge_lca_files(
    file1: str = DEFAULT_FILE1,
    file2: str = DEFAULT_FILE2,
    output: str = DEFAULT_OUTPUT,
) -> pd.DataFrame:
    """
    Read two LCA Excel files, concatenate them row-wise, and write the
    combined result to a new Excel file.

    If the two files have different column sets, columns present in only one
    file will be included in the output with NaN values for rows from the
    other file.  A warning is printed for each such mismatch.

    Parameters
    ----------
    file1 : str
        Path to the first input Excel file.
    file2 : str
        Path to the second input Excel file.
    output : str
        Path for the merged output Excel file.

    Returns
    -------
    pd.DataFrame
        The merged DataFrame.
    """
    df1 = read_lca_file(file1)
    df2 = read_lca_file(file2)

    _validate_columns(df1, df2)

    merged = pd.concat([df1, df2], ignore_index=True)

    print(f"\nMerged dataset: {len(merged)} rows, {len(merged.columns)} columns")

    merged.to_excel(output, index=False)
    print(f"Output written to: {output}")

    return merged


def main() -> None:
    args = sys.argv[1:]
    file1 = args[0] if len(args) > 0 else DEFAULT_FILE1
    file2 = args[1] if len(args) > 1 else DEFAULT_FILE2
    output = args[2] if len(args) > 2 else DEFAULT_OUTPUT

    merge_lca_files(file1, file2, output)


if __name__ == "__main__":
    main()
