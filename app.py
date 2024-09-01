import streamlit as st
import pdfplumber
import pandas as pd
from docx import Document
from docx.shared import Pt
import re
from io import BytesIO

def main():
    st.title("PDF Table Extraction and DOCX Conversion")

    # File upload
    uploaded_file = st.file_uploader("Upload a PDF file", type="pdf")

    if uploaded_file is not None:
        # Extract tables using pdfplumber
        with pdfplumber.open(uploaded_file) as pdf:
            tables = []
            for page in pdf.pages:
                extracted_tables = page.extract_tables()
                tables.extend(extracted_tables)

        if len(tables) < 2:
            st.error("The uploaded PDF does not contain enough tables.")
            return

        # Select the second table (index 1)
        table_data = tables[1]

        # Handle duplicate and None column names
        cleaned_columns = []
        for i, column_name in enumerate(table_data[0]):
            if column_name is None:
                cleaned_columns.append(f"Unnamed_{i}")
            elif table_data[0].count(column_name) > 1:
                cleaned_columns.append(f"{column_name}_{i}")
            else:
                cleaned_columns.append(column_name)

        # Create DataFrame from table data using cleaned column names
        df = pd.DataFrame(table_data[1:], columns=cleaned_columns)
        st.write("Extracted Table:")
        st.dataframe(df)

        # Select and rename specific columns
        selected_columns = df.iloc[:, [0, 3]]  # Adjust column indices based on table structure
        selected_columns.columns = ["Domain Scores", "Percentile"]

        # Clean and process the Percentile column
        selected_columns['Percentile'] = selected_columns['Percentile'].apply(clean_percentile)
        selected_columns.dropna(subset=['Percentile'], inplace=True)
        selected_columns['Grade'] = selected_columns['Percentile'].apply(grading_system)

        st.write("Processed Data with Grades:")
        st.dataframe(selected_columns)

        # Convert to DOCX and prepare for download
        docx_data = csv_to_docx_with_flagging(selected_columns)
        st.download_button(
            label="Download DOCX",
            data=docx_data,
            file_name="extracted_table_with_grades.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

def clean_percentile(value):
    match = re.search(r'\d+', str(value))
    return int(match.group(0)) if match else None

def grading_system(percentile):
    if percentile > 74:
        return "Above Average"
    elif 25 <= percentile <= 74:
        return "Average"
    elif 9 <= percentile <= 24:
        return "Low Average"
    elif 2 <= percentile <= 8:
        return "Low"
    else:
        return "Very Low"

def csv_to_docx_with_flagging(df):
    doc = Document()

    # Adding title with adjusted font size
    title = doc.add_paragraph("CNSVS Metrics with Percentiles and Grades")
    title_run = title.runs[0]
    title_run.font.size = Pt(16)

    for _, row in df.iterrows():
        text = f"{row['Domain Scores']}: {row['Percentile']}, {row['Grade']}"
        p = doc.add_paragraph(style='ListBullet')

        if row['Grade'] in ['Low Average', 'Low', 'Very Low']:
            run = p.add_run(text + " | ")
            run_bold = p.add_run("FLAG")
            run_bold.bold = True
        else:
            p.add_run(text)

    # Convert DOCX to a BytesIO object for download
    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

if __name__ == "__main__":
    main()
