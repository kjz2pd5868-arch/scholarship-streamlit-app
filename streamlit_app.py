import re
from io import BytesIO
from datetime import datetime

import pandas as pd
import streamlit as st
from pypdf import PdfReader
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


HEADER_FILL_COLOR = "1F4E78"
HEADER_FONT_COLOR = "FFFFFF"


def clean_text(text):
    if not text:
        return ""
    text = text.replace("\r", "\n")
    text = re.sub(r"https?://\S+", "", text)
    text = re.sub(r"\b\d{1,2}/\d{1,2}/\d{2,4},?\s+\d{1,2}:\d{2}\s*[APMapm]{2}\b", "", text)
    text = re.sub(r"\n[ \t]*\n+", "\n\n", text)
    text = re.sub(r"[ \t]+", " ", text)
    return text.strip()


def normalize(text):
    return re.sub(r"\s+", " ", text).strip() if text else ""


def get_texts_from_upload(uploaded_file):
    reader = PdfReader(uploaded_file)

    raw_pages = []
    for page in reader.pages:
        page_text = page.extract_text() or ""
        raw_pages.append(page_text)

    raw_text = "\n".join(raw_pages)
    cleaned_text = clean_text(raw_text)
    return raw_text, cleaned_text


def between(text, start, end):
    match = re.search(start + r"(.*?)" + end, text, re.S | re.I)
    return clean_text(match.group(1)) if match else ""


def single(text, label):
    patterns = [
        rf"{re.escape(label)}\s+(.*?)(?=\n[A-Z][A-Za-z /-]+(?:\n|$))",
        rf"{re.escape(label)}\s+(.*?)(?=\s+[A-Z][A-Za-z /-]+(?:\s|$))"
    ]
    for pattern in patterns:
        match = re.search(pattern, text, re.S)
        if match:
            value = clean_text(match.group(1))
            value = re.split(
                r"\b(?:Visibility|Financial Information|Opportunity-Specific Information|Department|Donor|Fund Code|Auxiliary Fund Code|Project ID|Type|Post-Acceptance|Source|Visible Award Amount)\b",
                value
            )[0].strip()
            return value
    return ""


def number(text, label):
    match = re.search(rf"{re.escape(label)}\s*\$?([\d,]+(?:\.\d+)?)", text)
    return float(match.group(1).replace(",", "")) if match else None


def gpa(text):
    patterns = [
        r"(\d\.\d)\s*GPA",
        r"GPA of at least (\d\.\d)",
        r"cumulative GPA of (\d\.\d)",
        r"maintain a GPA of at least (\d\.\d)"
    ]
    for pattern in patterns:
        match = re.search(pattern, text, re.I)
        if match:
            return float(match.group(1))
    if re.search(r"\bB average\b", text, re.I):
        return 3.0
    return None


def flag(text, pattern):
    return "Yes" if re.search(pattern, text, re.I) else "No"


def class_levels(text):
    levels = []
    ordered = ["Freshman", "Sophomore", "Junior", "Senior", "Graduate", "Undergraduate"]
    for level in ordered:
        if re.search(rf"\b{re.escape(level)}s?\b", text, re.I):
            levels.append(level)
    return "; ".join(levels)


def keywords_list(raw):
    return [normalize(x) for x in raw.split(",") if x.strip()]


def extract_name(raw_text, fallback_name):
    lines = [line.strip() for line in raw_text.splitlines() if line.strip()]

    skip_exact = {
        "Portfolio",
        "Applicant",
        "Basic Information",
        "Financial Information",
        "Opportunity-Specific Information",
        "Award Information",
        "Questions",
    }

    skip_patterns = [
        r"^Fall\s+\d{4}$",
        r"^Spring\s+\d{4}$",
        r"^\|\s*Ended",
        r"^https?://",
        r"^\d+/\d+$",
        r"^\d{1,2}/\d{1,2}/\d{2,4},?",
        r"^Name\s+",
    ]

    for line in lines:
        clean_line = normalize(line)

        if not clean_line:
            continue
        if clean_line in skip_exact:
            continue
        if any(re.search(pattern, clean_line, re.I) for pattern in skip_patterns):
            continue
        if "Eastern Illinois University Scholarships" in clean_line:
            continue

        return clean_line

    return fallback_name.rsplit(".", 1)[0]


def extract(uploaded_file):
    raw_text, text = get_texts_from_upload(uploaded_file)

    scholarship_name = extract_name(raw_text, uploaded_file.name)

    description = between(text, r"Description\s*", r"\s*Full\s+[Dd]escription")
    full_description = between(text, r"Full\s+[Dd]escription:?\s*", r"\s*Keywords:")
    keywords_raw = between(text, r"Keywords:\s*", r"\s*Type\b")

    combined = f"{description} {full_description}"
    keywords = keywords_list(keywords_raw)

    return {
        "Scholarship Name": scholarship_name,
        "Import Date": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "Fund Period Amount": number(text, "Fund Period Amount"),
        "Department": single(raw_text, "Department"),
        "Donor": single(raw_text, "Donor"),
        "Fund Code": single(raw_text, "Fund Code"),
        "Opportunity Type": single(raw_text, "Type"),
        "Post-Acceptance Enabled": single(raw_text, "Post-Acceptance"),
        "Minimum GPA": gpa(combined),
        "Full-Time Required": flag(combined, r"full-time"),
        "Class Level Eligible": class_levels(combined),
        "Geographic Preference": flag(combined, r"Illinois|Bloomington|Normal|McLean County"),
        "Financial Need Considered": flag(combined, r"financial need"),
        "Renewable / Reapply Allowed": flag(combined, r"apply again|eligible to apply again|continues to meet criteria"),
        "Resume Required": flag(combined, r"resume"),
        "Essay Required": flag(combined, r"essay|brief summary|short essay|summary"),
        "Recommendation Required": flag(combined, r"recommendation"),
        "Character / Leadership Mentioned": flag(combined, r"character|leadership|work ethic|personal values|motivation|goals|magnetic personality"),
        "Keyword Count": len(keywords),
        "Notes": "",
        "Keywords Raw": keywords_raw,
        "Source Application Type": single(raw_text, "Source"),
    }


def summary(df):
    rows = [
        ["Total Scholarships", len(df)],
        ["Total Fund Amount", df["Fund Period Amount"].fillna(0).sum()],
        ["", ""],
        ["Source Application Type", "Count"]
    ]

    source_counts = df["Source Application Type"].fillna("Unknown")
    source_counts = source_counts.replace("", "Unknown").value_counts()

    for key, value in source_counts.items():
        rows.append([key, int(value)])

    return pd.DataFrame(rows, columns=["Metric", "Value"])


def instructions_sheet():
    rows = [
        ["1", "Upload all scholarship PDF files in the app uploader."],
        ["2", "Click the process button."],
        ["3", "The Excel file will be generated automatically."],
        ["4", "Download the finished Excel file."],
    ]
    return pd.DataFrame(rows, columns=["Step", "Instructions"])


def style_header_row(ws):
    header_fill = PatternFill(fill_type="solid", fgColor=HEADER_FILL_COLOR)
    header_font = Font(bold=True, color=HEADER_FONT_COLOR)
    thin_border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )

    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border

    ws.row_dimensions[1].height = 24


def format_instructions_sheet(ws):
    style_header_row(ws)
    ws.freeze_panes = "A2"
    ws.column_dimensions["A"].width = 10
    ws.column_dimensions["B"].width = 75


def format_scholarships_sheet(ws, df_columns):
    style_header_row(ws)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    col_index = {name: idx + 1 for idx, name in enumerate(df_columns)}

    preferred_widths = {
        "Scholarship Name": 38,
        "Import Date": 21,
        "Fund Period Amount": 16,
        "Department": 24,
        "Donor": 24,
        "Fund Code": 16,
        "Opportunity Type": 18,
        "Post-Acceptance Enabled": 20,
        "Minimum GPA": 12,
        "Full-Time Required": 14,
        "Class Level Eligible": 24,
        "Geographic Preference": 16,
        "Financial Need Considered": 18,
        "Renewable / Reapply Allowed": 20,
        "Resume Required": 14,
        "Essay Required": 14,
        "Recommendation Required": 18,
        "Character / Leadership Mentioned": 24,
        "Keyword Count": 12,
        "Notes": 28,
    }

    for col_name, width in preferred_widths.items():
        if col_name in col_index:
            letter = get_column_letter(col_index[col_name])
            ws.column_dimensions[letter].width = width


def format_summary_sheet(ws):
    style_header_row(ws)
    ws.freeze_panes = "A2"
    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 18


def build_excel_bytes(df):
    main_df = df.drop(columns=["Keywords Raw", "Source Application Type"])
    summary_df = summary(df)
    instructions_df = instructions_sheet()

    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        instructions_df.to_excel(writer, sheet_name="Instructions", index=False)
        main_df.to_excel(writer, sheet_name="Scholarships", index=False)
        summary_df.to_excel(writer, sheet_name="Summary", index=False)

        wb = writer.book
        ws_instructions = wb["Instructions"]
        ws_scholarships = wb["Scholarships"]
        ws_summary = wb["Summary"]

        format_instructions_sheet(ws_instructions)
        format_scholarships_sheet(ws_scholarships, list(main_df.columns))
        format_summary_sheet(ws_summary)

    output.seek(0)
    return output


st.set_page_config(page_title="Scholarship Import Tool", layout="wide")
st.title("Scholarship Import Tool")
st.write("Upload scholarship PDFs, then download the finished Excel workbook.")

uploaded_files = st.file_uploader(
    "Upload PDF files",
    type=["pdf"],
    accept_multiple_files=True
)

if uploaded_files:
    st.write(f"{len(uploaded_files)} PDF file(s) ready.")

    if st.button("Process PDFs"):
        with st.spinner("Reading PDFs and building workbook..."):
            data = [extract(f) for f in uploaded_files]
            df = pd.DataFrame(data)
            excel_file = build_excel_bytes(df)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"scholarship_database_v2_{timestamp}.xlsx"

        st.success("Done.")
        st.dataframe(df.drop(columns=["Keywords Raw", "Source Application Type"]), use_container_width=True)

        st.download_button(
            label="Download Excel file",
            data=excel_file,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.info("Upload one or more PDF files to begin.")