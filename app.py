
import re
import os
from io import BytesIO

import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


st.set_page_config(
    page_title="Huawei MML Dump Parser",
    page_icon="📊",
    layout="wide"
)

st.title("Huawei MML Dump Parser")
st.write("Upload file MML TXT untuk dikonversi menjadi Excel.")


def parse_mml_to_excel(file_bytes):
    text = file_bytes.decode("utf-8", errors="ignore")
    lines = text.splitlines()

    wb = Workbook()

    if wb.active is not None:
        wb.remove(wb.active)

    summary_ws = wb.create_sheet("SUMMARY")
    sheet_dict = {}
    summary_count = {}

    RE_CMD = re.compile(r"^[A-Z0-9_ ]+:\s*;$")
    RE_SEP = re.compile(r"^-{10,}\s*$")
    RE_RESULT = re.compile(r"^\(Number of results")

    i = 0

    while i < len(lines):
        line = lines[i].strip()

        if not RE_CMD.match(line):
            i += 1
            continue

        neinfo = line
        command_name = re.sub(r"[:;\s]+", "", line).strip()

        sheet_name = re.sub(
            r"[:\\/?*\[\]]", "_", command_name
        )[:31]

        if not sheet_name:
            i += 1
            continue

        nename = ""

        for x in range(i + 1, min(i + 10, len(lines))):
            temp = lines[x].strip()

            if "#MC" in temp and not temp.startswith("+++"):
                nename = temp
                break

        j = i + 1
        sep_index = -1

        while j < len(lines):
            if RE_SEP.match(lines[j].strip()):
                sep_index = j
                break

            if RE_CMD.match(lines[j].strip()):
                break

            j += 1

        if sep_index == -1:
            i += 1
            continue

        table_type = ""

        for t in range(sep_index - 1, i, -1):
            if lines[t].strip():
                table_type = lines[t].strip()
                break

        if sep_index + 1 >= len(lines):
            i = sep_index + 1
            continue

        header_line = lines[sep_index + 1].strip()

        if not header_line:
            i = sep_index + 1
            continue

        headers = re.split(r"\s{2,}", header_line)
        headers = ["NEName", "NEInfo", "TableType"] + headers

        if sheet_name not in sheet_dict:
            ws = wb.create_sheet(sheet_name)
            sheet_dict[sheet_name] = ws

            for col_num, header in enumerate(headers, 1):
                ws.cell(
                    row=1,
                    column=col_num,
                    value=header
                ).font = Font(bold=True)

        else:
            ws = sheet_dict[sheet_name]

        start_row = (ws.max_row or 0) + 1
        k = sep_index + 2
        data_count = 0

        while k < len(lines):
            row_strip = lines[k].strip()

            if not row_strip:
                k += 1
                continue

            if RE_RESULT.match(row_strip):
                k += 1
                continue

            if RE_SEP.match(row_strip) or row_strip.startswith("---"):
                break

            if RE_CMD.match(row_strip):
                break

            cols = re.split(r"\s{2,}", row_strip)

            row_data = [
                nename,
                neinfo,
                table_type
            ] + cols

            for col_num, value in enumerate(row_data, 1):
                ws.cell(
                    row=start_row,
                    column=col_num,
                    value=value
                )

            start_row += 1
            data_count += 1
            k += 1

        summary_count[sheet_name] = (
            summary_count.get(sheet_name, 0) + data_count
        )

        i = k

    # SUMMARY
    summary_ws.append(["Sheet Name", "Rows"])

    for cell in summary_ws[1]:
        cell.font = Font(bold=True)

    for sheet_name, rows in sorted(summary_count.items()):
        row_num = (summary_ws.max_row or 0) + 1
        safe_name = sheet_name.replace("'", "''")

        cell = summary_ws.cell(
            row=row_num,
            column=1,
            value=sheet_name
        )

        cell.hyperlink = f"#'{safe_name}'!A1"
        cell.style = "Hyperlink"

        summary_ws.cell(
            row=row_num,
            column=2,
            value=rows
        )

    # Freeze panes and column widths
    for ws in wb.worksheets:
        ws.freeze_panes = "A2"

        for col_cells in ws.columns:
            if not col_cells:
                continue

            max_len = 0
            col_idx = col_cells[0].column or 1
            letter = get_column_letter(col_idx)

            for cell in col_cells[:200]:
                if cell.value is not None:
                    max_len = max(
                        max_len,
                        len(str(cell.value))
                    )

            ws.column_dimensions[letter].width = min(
                max_len + 2, 50
            )

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    return (
        output,
        len(sheet_dict),
        sum(summary_count.values())
    )


# UPLOAD FILE
uploaded_file = st.file_uploader(
    "Pilih File MML TXT",
    type=["txt", "log"],
    help="Upload dump MML Huawei berformat TXT atau LOG."
)

if uploaded_file is not None:
    st.success(f"File dipilih: {uploaded_file.name}")

    if st.button("Proses MML dan Buat Excel", type="primary"):

        try:
            with st.spinner("Sedang memproses file..."):
                excel_output, total_sheets, total_rows = (
                    parse_mml_to_excel(
                        uploaded_file.getvalue()
                    )
                )

            base_name = os.path.splitext(
                uploaded_file.name
            )[0]

            output_name = f"{base_name}.xlsx"

            st.success("Excel berhasil dibuat!")

            col1, col2 = st.columns(2)

            col1.metric(
                "Total Sheet Command",
                total_sheets
            )

            col2.metric(
                "Total Baris Data",
                total_rows
            )

            st.download_button(
                label="Download Hasil Excel",
                data=excel_output.getvalue(),
                file_name=output_name,
                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                ),
                type="primary"
            )

            if total_sheets == 0:
                st.warning(
                    "Tidak ada command yang terdeteksi. "
                    "Periksa format file MML dan separator."
                )

        except Exception as exc:
            st.error(f"Gagal memproses file: {exc}")
            st.exception(exc)

else:
    st.info(
        "Silakan upload file MML TXT terlebih dahulu."
    )
