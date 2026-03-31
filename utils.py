"""Shared utility functions for the procurement categorization pipeline."""
import pandas as pd


def safe_num_series(s: pd.Series) -> pd.Series:
    """Convert a currency-formatted Series to numeric, coercing errors to 0."""
    return (
        s.astype(str)
        .str.replace("$", "", regex=False)
        .str.replace(",", "", regex=False)
        .str.strip()
        .pipe(pd.to_numeric, errors="coerce")
        .fillna(0.0)
    )


def read_csv_robust(path: str) -> pd.DataFrame:
    """Read a CSV trying multiple encodings, skipping bad lines."""
    for enc in ("utf-8-sig", "utf-8", "latin-1"):
        try:
            return pd.read_csv(path, low_memory=False, encoding=enc)
        except UnicodeDecodeError:
            continue
        except Exception:
            break
    return pd.read_csv(path, encoding="latin-1",
                       on_bad_lines="skip", engine="python")
