
from pathlib import Path
import tempfile
import zipfile

import pandas as pd
import streamlit as st

from docx_engine import (
    create_zip,
    find_placeholders,
    generate_batch,
)
from io import BytesIO

st.set_page_config(
    page_title="Bulk DOCX to PDF",
    page_icon="📄",
    layout="wide",
)

st.title("Bulk DOCX → PDF Generator")
st.caption("Generate personalized PDF documents locally from a Word template and Excel workbook.")

template_file = st.file_uploader(
    "1. Upload Word template (.docx)",
    type=["docx"],
)
excel_file = st.file_uploader(
    "2. Upload Excel workbook (.xlsx)",
    type=["xlsx"],
)

if template_file and excel_file:
    # Use a temporary workspace; never trust uploaded filenames as paths.
    with tempfile.TemporaryDirectory(prefix="bulk_pdf_") as tmp:
        workspace = Path(tmp)
        template_path = workspace / "template.docx"
        template_path.write_bytes(template_file.getvalue())

        excel_stream = BytesIO(excel_file.getvalue())

        try:
            workbook = pd.ExcelFile(excel_stream, engine="openpyxl")
        except Exception as exc:
            st.error(f"Could not read the Excel workbook: {exc}")
            st.stop()

        sheet_name = st.selectbox(
            "3. Choose worksheet",
            workbook.sheet_names,
        )

        try:
            df = pd.read_excel(
                excel_stream,
                sheet_name=sheet_name,
                engine="openpyxl",
                dtype=object,
            )
        except Exception as exc:
            st.error(f"Could not read the worksheet: {exc}")
            st.stop()

        # Normalize header names and remove completely blank rows.
        df.columns = [str(c).strip() for c in df.columns]
        df = df.dropna(how="all").reset_index(drop=True)

        if df.empty:
            st.warning("The selected worksheet has no data rows.")
            st.stop()

        if df.columns.duplicated().any():
            st.error("Excel contains duplicate column names. Rename them and retry.")
            st.stop()

        st.subheader("4. Review spreadsheet")
        st.write(f"Detected **{len(df)} data rows**.")
        st.dataframe(df.head(10), use_container_width=True)

        try:
            placeholders = find_placeholders(template_path)
        except Exception as exc:
            st.error(f"Could not read the DOCX template: {exc}")
            st.stop()

        st.subheader("5. Map your Excel columns")

        if not placeholders:
            st.warning(
                "No placeholders were detected. Add tokens such as "
                "{{NAME}} and {{ID}} to the Word template."
            )
            st.stop()

        st.write("Detected template placeholders:")
        st.code(", ".join("{{" + p + "}}" for p in sorted(placeholders)))

        columns = list(df.columns)

        def suggest_column(placeholder: str) -> str:
            normalized = placeholder.strip().casefold()
            aliases = {
                "name": ["name", "full name", "employee name", "person name"],
                "id": ["id", "employee id", "employee_id", "emp id"],
            }
            candidates = aliases.get(normalized, [normalized])
            for col in columns:
                if col.casefold() in candidates:
                    return col
            return columns[0]

        mapping = {}
        for placeholder in sorted(placeholders):
            default_col = suggest_column(placeholder)
            default_index = columns.index(default_col)
            mapping[placeholder] = st.selectbox(
                f"{{{{{placeholder}}}}} maps to",
                options=columns,
                index=default_index,
                key=f"map_{placeholder}",
            )

        name_candidates = [
            c for c in columns
            if c.casefold() in {"name", "full name", "employee name", "person name"}
        ]
        id_candidates = [
            c for c in columns
            if c.casefold() in {"id", "employee id", "employee_id", "emp id"}
        ]

        name_default = name_candidates[0] if name_candidates else columns[0]
        id_default = id_candidates[0] if id_candidates else columns[-1]

        left, right = st.columns(2)
        with left:
            name_column = st.selectbox(
                "Name column for output filenames",
                columns,
                index=columns.index(name_default),
            )
        with right:
            id_column = st.selectbox(
                "ID column for output filenames",
                columns,
                index=columns.index(id_default),
            )

        # Map spreadsheet headers to the actual placeholder names.
        # Multiple placeholders may intentionally use the same column.
        missing_columns = set(mapping.values()) - set(df.columns)
        if missing_columns:
            st.error(f"Missing Excel columns: {missing_columns}")
            st.stop()

        st.info(
            "Review your mapping carefully. The first 10 spreadsheet rows "
            "are shown above; generation uses all non-empty rows."
        )

        if st.button("Generate PDFs", type="primary", use_container_width=True):
            # Construct a data frame with placeholder names as its headers.
            mapped_df = pd.DataFrame({
                placeholder: df[column]
                for placeholder, column in mapping.items()
            })

            # Keep filename columns too, even if they aren't placeholders.
            mapped_df["__OUTPUT_NAME"] = df[name_column]
            mapped_df["__OUTPUT_ID"] = df[id_column]

            # The engine receives placeholder names as keys, so provide
            # a dedicated name/ID column for filenames.
            mapped_df["__FILENAME_NAME"] = mapped_df["__OUTPUT_NAME"]
            mapped_df["__FILENAME_ID"] = mapped_df["__OUTPUT_ID"]

            # Save a temporary workbook-like dataset using the mapped
            # placeholder columns and explicit filename fields.
            # Pass actual filename-column keys to the generator.
            output_dir = workspace / "pdfs"

            progress = st.progress(0.0)
            status = st.empty()

            def update_progress(done, total):
                progress.progress(done / total if total else 1.0)
                status.write(f"Processing row {done} of {total}...")

            # Generate filenames using explicit internal columns.
            generated, errors = generate_batch(
                template_path=template_path,
                dataframe=mapped_df.rename(columns={
                    "__FILENAME_NAME": "__GEN_NAME",
                    "__FILENAME_ID": "__GEN_ID",
                }),
                output_dir=output_dir,
                name_column="__GEN_NAME",
                id_column="__GEN_ID",
                progress_callback=update_progress,
            )

            status.write("Generation complete.")
            st.success(f"Generated {len(generated)} PDF(s).")

            if errors:
                st.warning(f"{len(errors)} row(s) failed or were skipped.")
                st.dataframe(pd.DataFrame(errors), use_container_width=True)
                st.download_button(
                    "Download error report (CSV)",
                    data=pd.DataFrame(errors).to_csv(index=False).encode("utf-8"),
                    file_name="generation_errors.csv",
                    mime="text/csv",
                )

            if generated:
                zip_path = workspace / "generated_pdfs.zip"
                create_zip(generated, zip_path)
                st.download_button(
                    "Download all PDFs (ZIP)",
                    data=zip_path.read_bytes(),
                    file_name="generated_pdfs.zip",
                    mime="application/zip",
                    type="primary",
                    use_container_width=True,
                )
