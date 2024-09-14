import streamlit as st
import pdfplumber
import pandas as pd
from docx import Document
from docx.shared import Pt
import re
from io import BytesIO
from docx.shared import RGBColor


def main():
    st.title("PDF Data Extraction and DOCX Conversion")

    # File upload
    uploaded_file = st.file_uploader("Upload a PDF file", type="pdf")

    if uploaded_file is not None:
        with pdfplumber.open(uploaded_file) as pdf:
            tables = []
            for page in pdf.pages:
                extracted_tables = page.extract_tables()
                tables.extend(extracted_tables)

        # Initialize table-related data variables
        df1 = None
        df2 = None

        # Check if any tables exist
        if tables:
            # Extract and process the upper table (second table in the PDF) if present
            if len(tables) > 1:
                table_data_1 = tables[1]
                cleaned_columns_1 = []
                for i, column_name in enumerate(table_data_1[0]):
                    if column_name is None:
                        cleaned_columns_1.append(f"Unnamed_{i}")
                    elif table_data_1[0].count(column_name) > 1:
                        cleaned_columns_1.append(f"{column_name}_{i}")
                    else:
                        cleaned_columns_1.append(column_name)

                df1 = pd.DataFrame(table_data_1[1:], columns=cleaned_columns_1)
                st.write("Extracted Upper Table:")
                st.dataframe(df1)

                # Process the upper table
                selected_columns_1 = df1.iloc[:, [0, 3]]
                selected_columns_1.columns = ["Domain Scores", "Percentile"]
                selected_columns_1['Percentile'] = selected_columns_1['Percentile'].apply(clean_percentile)
                selected_columns_1.dropna(subset=['Percentile'], inplace=True)
                selected_columns_1['Grade'] = selected_columns_1['Percentile'].apply(grading_system)

                st.write("Processed Data with Grades for Upper Table:")
                st.dataframe(selected_columns_1)

            # Extract and process the lower table containing Domain, Score, and Severity if present
            lower_table_data = None
            for i, table in enumerate(tables):
                # Check each table for the lower table structure based on known headers
                if len(table) > 0 and len(table[0]) >= 3 and table[0][0] == "Domain" and table[0][1] == "Score" and table[0][2] == "Severity":
                    lower_table_data = table
                    break

            if lower_table_data:
                # Handle rows with extra columns
                cleaned_lower_table_data = [row[:3] for row in lower_table_data[1:] if len(row) >= 3]
                df2 = pd.DataFrame(cleaned_lower_table_data, columns=["Domain", "Score", "Severity"])

                st.write("Extracted Lower Table (Domain, Score, Severity):")
                st.dataframe(df2)

                # Clean and process the lower table
                df2['Score'] = df2['Score'].apply(clean_score)
                df2.dropna(subset=['Score'], inplace=True)
                df2['Grade'] = df2['Severity'].apply(grading_system_2)

                st.write("Processed Data with Grades for Lower Table:")
                st.dataframe(df2)

        else:
            st.error("The uploaded PDF does not contain any tables.")
            return

        # Combine both dataframes if available or process separately
        if df1 is not None and df2 is not None:
            combined_df = pd.concat([selected_columns_1, df2], axis=0, ignore_index=True)
        elif df1 is not None:
            combined_df = selected_columns_1
        elif df2 is not None:
            combined_df = df2
        else:
            combined_df = pd.DataFrame()
            st.error("No table data was extracted.")

        # Extract PHQ-9 and PCL-5 data if available
        phq9_data, pcl5_data = extract_phq9_and_pcl5(uploaded_file)

        # Display PHQ-9 Score with Interpretation
        if phq9_data:
            phq9_score = phq9_data['PHQ-9 Score']
            interpretation = interpret_phq9(phq9_score)
            st.markdown("**Patient Health Questionnaire (PHQ-9)**")
            st.write(f"• PHQ-9 Score: {phq9_score} ({interpretation})")

        # Display PCL-5 Scores if available
        if pcl5_data:
            st.markdown("**PTSD Checklist (PCL-5) SF-20**")
            st.write(f"• Intrusion (Items 1 - 5): {pcl5_data['Intrusion']}")
            st.write(f"• Persistent Avoidance (Items 6 - 7): {pcl5_data['Persistent Avoidance']}")
            st.write(f"• Negative Alterations in Cognitions and Mood (Items 8 - 14): {pcl5_data['Negative Alterations in Cognitions and Mood']}")
            st.write(f"• Alterations in Arousal and Reactivity (Items 15 - 20): {pcl5_data['Alterations in Arousal and Reactivity']}")
            st.write(f"• Total Score (Items 1 - 20): {pcl5_data['Total Score']}")

        # Convert to DOCX and prepare for download
        if not combined_df.empty or phq9_data or pcl5_data:
            docx_data = csv_to_docx_with_flagging(combined_df, phq9_data, pcl5_data)
            st.download_button(
                label="Download DOCX",
                data=docx_data,
                file_name="extracted_data.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
        else:
            st.error("No data to convert to DOCX.")

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

def clean_score(value):
    match = re.search(r'\d+', str(value))
    return int(match.group(0)) if match else None

def grading_system_2(severity):
    if severity in ["Mild", "Moderate", "Severe"]:
        return "FLAG"
    else:
        return None

def csv_to_docx_with_flagging(df, phq9_data, pcl5_data):
    doc = Document()

    # Adding title with adjusted font size
    title = doc.add_paragraph("CNSVS Metrics with Percentiles, Scores, and Flags")
    title_run = title.runs[0]
    title_run.font.size = Pt(16)
    title_run.bold = True  # Make the text bold
    title_run.font.color.rgb = RGBColor(0, 0, 0)  # Set font color to black

    # Flag to track when to insert the new heading
    npq_heading_added = False

    for _, row in df.iterrows():
        if 'Domain Scores' in row and pd.notna(row['Domain Scores']):
            text = f"{row['Domain Scores']}: {row['Percentile']}, {row['Grade']}"
        elif 'Domain' in row and pd.notna(row['Domain']):
            text = f"{row['Domain']}: {row['Score']}, {row['Severity']}"
            # Check if the current row is part of the NeuroPsych Questionnaire (NPQ)
            if row['Domain'] == "Attention" and not npq_heading_added:
                # Add the new heading for the NPQ
                heading = doc.add_paragraph()
                heading_run = heading.add_run("NeuroPsych Questionnaire (NPQ) SF-45")  # Create the run and add the text
                heading_run.bold = True  # Make the text bold
                heading_run.font.size = Pt(14)  # Adjust the size if needed
                heading_run.font.color.rgb = RGBColor(0, 0, 0)  # Set heading color to black
                npq_heading_added = True

        # Create a new paragraph
        p = doc.add_paragraph(style='ListBullet')

        # Add " | FLAG" if the grade is "Low Average," "Low," or "Very Low" for both tables
        if 'Grade' in row and row['Grade'] in ['Low Average', 'Low', 'Very Low']:
            run = p.add_run(text + " | ")
            run_bold = p.add_run("FLAG")
            run_bold.bold = True
        elif 'Grade' in row and row['Grade'] == "FLAG":
            run = p.add_run(text + " | ")
            run_bold = p.add_run("FLAG")
            run_bold.bold = True
        else:
            p.add_run(text)

    # Add PHQ-9 and PCL-5 data to the DOCX document
    if phq9_data:
        phq9_paragraph = doc.add_paragraph()
        phq9_run = phq9_paragraph.add_run("Patient Health Questionnaire (PHQ-9)")
        phq9_run.bold = True  # Bold heading for PHQ-9
        phq9_paragraph.add_run(f"\n• PHQ-9 Score: {phq9_data['PHQ-9 Score']} ({interpret_phq9(phq9_data['PHQ-9 Score'])})")

    if pcl5_data:
        pcl5_paragraph = doc.add_paragraph()
        pcl5_run = pcl5_paragraph.add_run("PTSD Checklist (PCL-5) SF-20")
        pcl5_run.bold = True  # Bold heading for PCL-5
        pcl5_paragraph.add_run(f"\n• Intrusion (Items 1 - 5): {pcl5_data['Intrusion']}")
        pcl5_paragraph.add_run(f"\n• Persistent Avoidance (Items 6 - 7): {pcl5_data['Persistent Avoidance']}")
        pcl5_paragraph.add_run(f"\n• Negative Alterations in Cognitions and Mood (Items 8 - 14): {pcl5_data['Negative Alterations in Cognitions and Mood']}")
        pcl5_paragraph.add_run(f"\n• Alterations in Arousal and Reactivity (Items 15 - 20): {pcl5_data['Alterations in Arousal and Reactivity']}")
        pcl5_paragraph.add_run(f"\n• Total Score (Items 1 - 20): {pcl5_data['Total Score']}")

    # Convert DOCX to a BytesIO object for download
    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

def extract_phq9_and_pcl5(uploaded_file):
    phq9_data = {}
    pcl5_data = {}

    with pdfplumber.open(uploaded_file) as pdf:
        # Loop through all pages in the PDF
        for page in pdf.pages:
            text = page.extract_text()

            # Extract PHQ-9 Score
            if not phq9_data:  # Only extract if not already found
                phq9_match = re.search(r'PHQ-9 Score\s*(\d+)', text)
                if phq9_match:
                    phq9_data['PHQ-9 Score'] = int(phq9_match.group(1))

            # Extract PCL-5 Scores (Intrusion, Avoidance, Cognition, Arousal, and Total Score)
            if not pcl5_data:  # Only extract if not already found
                intrusion_match = re.search(r'Intrusion \(Items 1 - 5\)\s*([\d]+)', text)
                avoidance_match = re.search(r'Persistent Avoidance \(Items 6 - 7\)\s*([\d]+)', text)
                cognition_mood_match = re.search(r'Negative Alterations in Cognitions and Mood \(Items 8 - 14\)\s*([\d]+)', text)
                arousal_reactivity_match = re.search(r'Alterations in Arousal and Reactivity \(Items 15 - 20\)\s*([\d]+)', text)
                total_score_match = re.search(r'Total Score \(Items 1 - 20\)\s*([\d]+)', text)

                if intrusion_match and avoidance_match and cognition_mood_match and arousal_reactivity_match and total_score_match:
                    pcl5_data['Intrusion'] = int(intrusion_match.group(1))
                    pcl5_data['Persistent Avoidance'] = int(avoidance_match.group(1))
                    pcl5_data['Negative Alterations in Cognitions and Mood'] = int(cognition_mood_match.group(1))
                    pcl5_data['Alterations in Arousal and Reactivity'] = int(arousal_reactivity_match.group(1))
                    pcl5_data['Total Score'] = int(total_score_match.group(1))

    return phq9_data, pcl5_data

def interpret_phq9(score):
    if 1 <= score <= 4:
        return "Minimal depression"
    elif 5 <= score <= 9:
        return "Mild depression"
    elif 10 <= score <= 14:
        return "Moderate depression"
    elif 15 <= score <= 19:
        return "Moderately severe depression"
    elif 20 <= score <= 27:
        return "Severe depression"
    else:
        return "Unknown interpretation"


if __name__ == "__main__":
    main()
