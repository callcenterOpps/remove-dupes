import os
import re
import sqlite3
import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st

CHUNK_SIZE = 100_000


def get_columns(uploaded_file):
    """Read only the header/sample so the app can offer a column selector."""
    uploaded_file.seek(0)
    sample = pd.read_csv(uploaded_file, nrows=5, dtype=str, keep_default_na=False)
    uploaded_file.seek(0)
    sample.columns = sample.columns.astype(str).str.strip()
    return list(sample.columns)


def remove_duplicates_chunked(uploaded_file, selected_column, output_path, chunk_size=CHUNK_SIZE):
    """
    Deduplicate a CSV in chunks, preserving the first occurrence of each selected-column
    value across the entire file. SQLite tracks values already seen, avoiding a huge
    Python set for multi-million-row files.
    """
    uploaded_file.seek(0)
    conn = sqlite3.connect(":memory:")
    conn.execute("CREATE TABLE seen (value TEXT PRIMARY KEY)")
    conn.execute("CREATE TEMP TABLE candidates (value TEXT PRIMARY KEY, row_num INTEGER)")
    conn.commit()

    rows_processed = 0
    rows_kept = 0
    duplicates_removed = 0
    chunk_count = 0
    header_written = False

    try:
        reader = pd.read_csv(
            uploaded_file,
            dtype=str,
            keep_default_na=False,
            chunksize=chunk_size,
            encoding_errors="replace",
        )

        for chunk_count, chunk in enumerate(reader, start=1):
            chunk.columns = chunk.columns.astype(str).str.strip()
            if selected_column not in chunk.columns:
                raise ValueError(
                    f"Column '{selected_column}' was not found. Check whether the CSV header changed."
                )

            original_count = len(chunk)
            # Keep the first occurrence inside this chunk, preserving original row order.
            chunk = chunk.drop_duplicates(subset=[selected_column], keep="first").copy()

            conn.execute("DELETE FROM candidates")
            values = chunk[selected_column].astype(str).tolist()
            conn.executemany(
                "INSERT OR IGNORE INTO candidates(value, row_num) VALUES (?, ?)",
                [(value, i) for i, value in enumerate(values)],
            )

            # Candidate values not seen in earlier chunks; row_num keeps their order.
            new_row_numbers = [
                row[0]
                for row in conn.execute(
                    """
                    SELECT c.row_num
                    FROM candidates c
                    LEFT JOIN seen s ON s.value = c.value
                    WHERE s.value IS NULL
                    ORDER BY c.row_num
                    """
                )
            ]

            # Mark all candidate values as seen, including values already seen previously.
            conn.execute("INSERT OR IGNORE INTO seen(value) SELECT value FROM candidates")
            conn.commit()

            kept_chunk = chunk.iloc[new_row_numbers]
            duplicates_removed += original_count - len(kept_chunk)
            rows_processed += original_count
            rows_kept += len(kept_chunk)

            kept_chunk.to_csv(
                output_path,
                mode="a",
                index=False,
                header=not header_written,
                encoding="utf-8-sig" if not header_written else "utf-8",
            )
            header_written = True

            yield {
                "chunk": chunk_count,
                "processed": rows_processed,
                "kept": rows_kept,
                "duplicates": duplicates_removed,
            }

        if chunk_count == 0:
            raise ValueError("The uploaded CSV is empty or could not be read.")

        if not header_written:
            # Preserve headers when the file has headers but no data rows.
            uploaded_file.seek(0)
            empty_df = pd.read_csv(uploaded_file, nrows=0, dtype=str, keep_default_na=False)
            empty_df.columns = empty_df.columns.astype(str).str.strip()
            empty_df.to_csv(output_path, index=False, encoding="utf-8-sig")

    finally:
        conn.close()


st.set_page_config(page_title="CSV Duplicate Remover", page_icon="🧹", layout="wide")
st.title("🧹 CSV Duplicate Remover")
st.write(
    "Upload a CSV, choose the column used to identify duplicates, and download a CSV "
    "that keeps the first occurrence of each value."
)
st.info(
    "Large-file mode processes 100,000 rows per chunk and tracks duplicates in SQLite "
    "instead of a large Python set. Actual capacity depends on the host's RAM, disk, "
    "execution limits, and upload limits."
)
st.warning(
    "On Streamlit Cloud, uploaded files are transmitted to Streamlit's servers. "
    "The app does not upload them to GitHub or a permanent database, but the input is "
    "held in the active app session and the generated output uses temporary server storage."
)

uploaded_file = st.file_uploader("Select CSV file", type=["csv"])

if uploaded_file is not None:
    st.caption(f"File: {uploaded_file.name} · Size: {uploaded_file.size / (1024 * 1024):,.2f} MB")
    try:
        columns = get_columns(uploaded_file)
        if not columns:
            st.error("No columns were found in this CSV.")
        else:
            selected_column = st.selectbox(
                "Select column to remove duplicates by",
                options=columns,
                help="The first row for each distinct value in this column is kept.",
            )
            chunk_size = st.select_slider(
                "Rows processed per chunk",
                options=[25_000, 50_000, 100_000, 200_000],
                value=100_000,
                help="If the app runs out of memory, choose a smaller chunk size.",
            )

            if st.button("Remove Duplicates", type="primary"):
                output_name = f"deduped_{Path(uploaded_file.name).name}"
                if not output_name.lower().endswith(".csv"):
                    output_name += ".csv"
                output_path = os.path.join(
                    tempfile.mkdtemp(prefix="csv_dedupe_"),
                    output_name,
                )
                progress = st.progress(0, text="Starting...")
                status = st.empty()
                metrics = st.empty()
                last_stats = {"processed": 0, "kept": 0, "duplicates": 0, "chunk": 0}

                try:
                    # Total rows are unknown in advance; report chunk progress and counts.
                    for stats in remove_duplicates_chunked(
                        uploaded_file, selected_column, output_path, chunk_size
                    ):
                        last_stats = stats
                        status.write(
                            f"Processed **{stats['processed']:,}** rows · "
                            f"Duplicates removed: **{stats['duplicates']:,}**"
                        )
                        metrics.write(
                            f"Rows kept: **{stats['kept']:,}** · "
                            f"Chunks completed: **{stats['chunk']:,}**"
                        )
                        # Indeterminate total-row count: progress reflects completed work, not % of file.
                        progress.progress(min(0.95, stats["chunk"] / (stats["chunk"] + 2)))
                    progress.progress(1.0, text="Completed")
                    st.session_state["dedupe_output_path"] = output_path
                    st.session_state["dedupe_output_name"] = output_name
                    st.session_state["dedupe_stats"] = last_stats
                    st.success(
                        f"Completed. Removed {last_stats['duplicates']:,} duplicate rows; "
                        f"kept {last_stats['kept']:,} rows."
                    )
                except Exception as exc:
                    st.error(f"Could not process this CSV: {exc}")
                    if os.path.exists(output_path):
                        os.remove(output_path)
    except Exception as exc:
        st.error(f"Could not read the CSV header: {exc}")

if st.session_state.get("dedupe_output_path") and os.path.isfile(st.session_state["dedupe_output_path"]):
    output_path = st.session_state["dedupe_output_path"]
    output_name = st.session_state.get("dedupe_output_name", "deduped_output.csv")
    st.subheader("Download Result")
    st.caption(f"Output size: {os.path.getsize(output_path) / (1024 * 1024):,.2f} MB")
    # Streamlit accepts a binary file-like object; avoid reading the entire output into RAM.
    with open(output_path, "rb") as output_handle:
        st.download_button(
            "Download Deduplicated CSV",
            data=output_handle,
            file_name=output_name,
            mime="text/csv",
            use_container_width=True,
        )
