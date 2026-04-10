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

MAJOR_TERMS = [
    "accounting", "business", "finance", "marketing", "management",
    "economics", "computer information systems", "information systems",
    "computer science", "education", "elementary education",
    "special education", "nursing", "engineering", "biology",
    "chemistry", "physics", "mathematics", "math", "psychology",
    "sociology", "social work", "criminal justice", "history",
    "english", "journalism", "communication", "communications",
    "agriculture", "mba", "master of business administration", "stem"
]

LOCATION_WORDS = [
    "illinois", "indiana", "kentucky", "missouri", "iowa", "chicago",
    "bloomington", "normal", "mattoon", "charleston", "coles county",
    "cumberland county", "clay county", "richland county",
    "mclean county", "cook county", "central illinois",
    "southern illinois", "northern illinois", "united states"
]


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


def unique_keep_order(items):
    seen = set()
    output = []
    for item in items:
        key = item.casefold()
        if key not in seen:
            seen.add(key)
            output.append(item)
    return output


def get_texts_from_upload(uploaded_file):
    uploaded_file.seek(0)
    reader = PdfReader(uploaded_file)
    raw_pages = [(page.extract_text() or "") for page in reader.pages]
    raw_text = "\n".join(raw_pages)
    return raw_text, clean_text(raw_text)


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


def clean_for_requirement_matching(text):
    text = re.sub(r"\bEastern Illinois University\b", "", text, flags=re.I)
    text = re.sub(r"\bEIU\b", "", text, flags=re.I)
    return text


def gpa(text):
    patterns = [
        r"minimum GPA of (\d\.\d+|\d)",
        r"GPA of at least (\d\.\d+|\d)",
        r"cumulative GPA of (\d\.\d+|\d)",
        r"maintain a GPA of at least (\d\.\d+|\d)",
        r"have a GPA of (\d\.\d+|\d) or higher",
        r"(\d\.\d+|\d)\s*GPA"
    ]
    for pattern in patterns:
        match = re.search(pattern, text, re.I)
        if match:
            try:
                return float(match.group(1))
            except ValueError:
                pass
    if re.search(r"\bB average\b", text, re.I):
        return 3.0
    return None


def flag(text, pattern):
    return "Yes" if re.search(pattern, text, re.I) else "No"


def class_levels(text):
    levels = []

    checks = [
        ("Freshman", r"\bfreshm[ae]n?\b|first[- ]year"),
        ("Sophomore", r"\bsophomores?\b|second[- ]year"),
        ("Junior", r"\bjuniors?\b|third[- ]year"),
        ("Senior", r"\bseniors?\b|fourth[- ]year"),
        ("Graduate", r"\bgraduate\b|\bmaster'?s\b|\bMBA\b"),
        ("Undergraduate", r"\bundergraduate\b")
    ]

    for label, pattern in checks:
        if re.search(pattern, text, re.I):
            levels.append(label)

    return "; ".join(levels)


def extract_name(raw_text, fallback_name):
    lines = [line.strip() for line in raw_text.splitlines() if line.strip()]

    skip_exact = {
        "Portfolio", "Applicant", "Basic Information", "Financial Information",
        "Opportunity-Specific Information", "Award Information", "Questions",
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
        if not clean_line or clean_line in skip_exact:
            continue
        if any(re.search(pattern, clean_line, re.I) for pattern in skip_patterns):
            continue
        if "Eastern Illinois University Scholarships" in clean_line:
            continue
        return clean_line

    return fallback_name.rsplit(".", 1)[0]


def geographic_preference(text):
    text = clean_for_requirement_matching(text)

    context_patterns = [
        r"(?:resident|residents|residency|living|live|from|preference(?: given)? to students from|students from|permanent address in|graduates? of high schools? in|applicants? from)\s+([^.:\n;]+)",
        r"(?:must be|shall be)\s+(?:a\s+)?resident of\s+([^.:\n;]+)",
    ]

    found = []

    for pattern in context_patterns:
        for match in re.finditer(pattern, text, re.I):
            chunk = match.group(1)

            for county in re.findall(r"\b[A-Z][a-z]+ County\b", chunk):
                found.append(county)

            for phrase in LOCATION_WORDS:
                if re.search(rf"\b{re.escape(phrase)}\b", chunk, re.I):
                    found.append(phrase.title() if phrase != "mba" else "MBA")

    found = [x.replace("Mclean", "McLean").replace("Coles county", "Coles County")
               .replace("Clay county", "Clay County").replace("Richland county", "Richland County")
               .replace("Cumberland county", "Cumberland County").replace("Cook county", "Cook County")
               .replace("Central illinois", "Central Illinois").replace("Southern illinois", "Southern Illinois")
               .replace("Northern illinois", "Northern Illinois").replace("United states", "United States")
             for x in found]

    return ", ".join(unique_keep_order(found))


def financial_need(text):
    pattern = r"financial need|demonstrated need|FAFSA|Pell(?: Grant)?|need-based|economic need"
    return "Yes" if re.search(pattern, text, re.I) else "No"


def underserved_flag(text):
    pattern = (
        r"low[- ]income|underprivileged|underserved|disadvantaged|"
        r"first[- ]generation|first generation|underrepresented|"
        r"historically marginalized|background of hardship"
    )
    return "Yes" if re.search(pattern, text, re.I) else "No"


def major_field(text):
    text = clean_for_requirement_matching(text)

    context_patterns = [
        r"(?:major(?:ing)? in|majors? in|major field of study in|students? in|enrolled in|pursuing a degree in|degree in)\s+([^.:\n;]+)",
        r"(?:accounting students|business students|education students|MBA students)",
    ]

    found = []

    for term in MAJOR_TERMS:
        # tighter: only count majors when they appear in requirement-like contexts
        for pattern in context_patterns:
            for match in re.finditer(pattern, text, re.I):
                chunk = match.group(0)
                if re.search(rf"\b{re.escape(term)}\b", chunk, re.I):
                    found.append(term)

    # direct scholarship phrases
    direct_map = {
        "accounting students": "Accounting",
        "business students": "Business",
        "education students": "Education",
        "mba students": "MBA",
    }
    for phrase, label in direct_map.items():
        if re.search(rf"\b{re.escape(phrase)}\b", text, re.I):
            found.append(label)

    cleaned = []
    for item in unique_keep_order(found):
        if item.lower() == "mba":
            cleaned.append("MBA")
        elif item.lower() == "stem":
            cleaned.append("STEM")
        else:
            cleaned.append(item.title())

    return ", ".join(cleaned)


def deadline(text):
    patterns = [
        r"deadline[:\s]+([A-Za-z]+\s+\d{1,2},\s+\d{4})",
        r"deadline[:\s]+(\d{1,2}/\d{1,2}/\d{2,4})",
        r"applications? due[:\s]+([A-Za-z]+\s+\d{1,2},\s+\d{4})",
        r"applications? due[:\s]+(\d{1,2}/\d{1,2}/\d{2,4})",
        r"due date[:\s]+([A-Za-z]+\s+\d{1,2},\s+\d{4})",
        r"due date[:\s]+(\d{1,2}/\d{1,2}/\d{2,4})",
    ]
    for pattern in patterns:
        match = re.search(pattern, text, re.I)
        if match:
            return match.group(1).strip()
    return ""


def extract(uploaded_file):
    raw_text, text = get_texts_from_upload(uploaded_file)

    scholarship_name = extract_name(raw_text, uploaded_file.name)
    description = between(text, r"Description\s*", r"\s*Full\s+[Dd]escription")
    full_description = between(text, r"Full\s+[Dd]escription:?\s*", r"\s*Keywords:")
    keywords_raw = between(text, r"Keywords:\s*", r"\s*Type\b")

    combined = clean_text(f"{description} {full_description} {keywords_raw}")

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
        "Full-Time Required": flag(combined, r"full[- ]time"),
        "Class Level Eligible": class_levels(combined),
        "Geographic Preference": geographic_preference(combined),
        "Financial Need Considered": financial_need(combined),
        "Low Income / Underprivileged Background": underserved_flag(combined),
        "Major / Field of Study": major_field(combined),
        "Deadline": deadline(text + "\n" + combined),
        "Renewable / Reapply Allowed": flag(
            combined,
            r"apply again|eligible to apply again|continues to meet criteria|renewable|may reapply"
        ),
        "Resume Required": flag(combined, r"\bresume\b|\bcv\b|curriculum vitae"),
        "Essay Required": flag(combined, r"\bessay\b|brief summary|short essay|personal statement|statement of purpose"),
        "Recommendation Required": flag(combined, r"recommendation|reference letter|letter of recommendation"),
        "Character / Leadership Mentioned": flag(
            combined,
            r"character|leadership|work ethic|personal values|motivation|goals|service|integrity"
        ),
        "Notes": "",
    }


def summary(df):
    rows = [
        ["Total Scholarships", len(df)],
        ["Total Fund Amount", df["Fund Period Amount"].fillna(0).sum()],
        ["", ""],
        ["Opportunity Type", "Count"]
    ]

    counts = df["Opportunity Type"].fillna("Unknown").replace("", "Unknown").value_counts()
    for key, value in counts.items():
        rows.append([key, int(value)])

    return pd.DataFrame(rows, columns=["Metric", "Value"])


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


def format_scholarships_sheet(ws, df_columns):
    style_header_row(ws)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    col_index = {name: idx + 1 for idx, name in enumerate(df_columns)}

    widths = {
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
        "Geographic Preference": 28,
        "Financial Need Considered": 18,
        "Low Income / Underprivileged Background": 26,
        "Major / Field of Study": 24,
        "Deadline": 16,
        "Renewable / Reapply Allowed": 20,
        "Resume Required": 14,
        "Essay Required": 14,
        "Recommendation Required": 18,
        "Character / Leadership Mentioned": 24,
        "Notes": 28,
    }

    for col_name, width in widths.items():
        if col_name in col_index:
            ws.column_dimensions[get_column_letter(col_index[col_name])].width = width


def format_summary_sheet(ws):
    style_header_row(ws)
    ws.freeze_panes = "A2"
    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 18


def build_excel_bytes(df):
    summary_df = summary(df)
    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Scholarships", index=False)
        summary_df.to_excel(writer, sheet_name="Summary", index=False)

        wb = writer.book
        format_scholarships_sheet(wb["Scholarships"], list(df.columns))
        format_summary_sheet(wb["Summary"])

    output.seek(0)
    return output


st.set_page_config(page_title="Scholarship Data Extraction Tool", layout="wide")
st.title("Scholarship Data Extraction Tool")
st.markdown("""
Upload scholarship PDFs, review the extracted fields, and download a formatted Excel workbook.

**How to use**
1. Upload one or more PDF files  
2. Click **Process PDFs**  
3. Download the finished Excel file
""")

uploaded_files = st.file_uploader("Upload PDF files", type=["pdf"], accept_multiple_files=True)

if uploaded_files:
    st.write(f"{len(uploaded_files)} PDF file(s) ready.")

    if st.button("Process PDFs"):
        data = []
        errors = []

        with st.spinner("Reading PDFs and building workbook..."):
            for f in uploaded_files:
                try:
                    data.append(extract(f))
                except Exception as e:
                    errors.append((f.name, str(e)))

        if data:
            df = pd.DataFrame(data)
            excel_file = build_excel_bytes(df)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"scholarship_database_v2_{timestamp}.xlsx"

            st.success("Processing complete.")
            st.dataframe(df, use_container_width=True)

            st.download_button(
                label="Download Excel file",
                data=excel_file,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        if errors:
            st.warning("Some files could not be processed.")
            for name, err in errors:
                st.error(f"{name}: {err}")
else:
    st.info("Upload one or more PDF files to begin.")
