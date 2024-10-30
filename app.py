import streamlit as st
from datetime import datetime
from docx import Document
import pandas as pd
import re

# Set page configuration with a favicon
st.set_page_config(
    page_title="Skills Boot Camp Weekly Timesheet",
    page_icon="https://lirp.cdn-website.com/d8120025/dms3rep/multi/opt/social-image-88w.png", 
    layout="centered"  # "centered" or "wide"
)

# Initialize session state for screen navigation
if 'page' not in st.session_state:
    st.session_state.page = 1

# Function to load DOCX data
def load_docx_data():
    doc = Document("Skills Boot Camp Week 1 Timesheet.docx")
    
    # Extract first table data
    day, session_activity, facilitator, time, notes_comments = [], [], [], [], []
    attendance_data = []  # for second table
    
    for table_idx, table in enumerate(doc.tables):
        for row in table.rows:
            cells = [cell.text.replace('\n', ' ').strip() for cell in row.cells]
            if table_idx == 0 and len(cells) == 5:
                day.append(cells[0])
                session_activity.append(cells[1])
                facilitator.append(cells[2])
                time.append(cells[3])
                notes_comments.append(cells[4])
            elif table_idx == 1 and len(cells) >= 4:
                while len(cells) < 5:
                    cells.append("")  # Add empty cell if columns are missing
                attendance_data.append(cells[:5])

    df1 = pd.DataFrame({
        "Day": day,
        "Session/Activity": session_activity,
        "Facilitator": facilitator,
        "Time": time,
        "Notes/Comments": notes_comments
    })
    df2 = pd.DataFrame(attendance_data, columns=["Day", "Date", "Arrival Time", "Departure Time", "Learner Signature"])
    
    return df1, df2

# Load data from DOCX
df1, df2 = load_docx_data()

# First Screen: Display the Timesheet table
if st.session_state.page == 1:
    st.header("Skills Boot Camp Week 1 Timesheet")
    st.text("Weekly Timesheet: Week 1 16/09/2024 – 20/09/2024 (10:00 AM - 1:00 PM) ")
    st.markdown(df1.to_html(index=False, header=False), unsafe_allow_html=True)
    
    if st.button("Next"):
        st.session_state.page = 2
        st.experimental_rerun()

# Second Screen: Learner Declaration and Attendance Table
elif st.session_state.page == 2:    
    st.subheader("Attendance Register Declaration (Monday - Friday)")
    learner_name = st.text_input("Enter your full name")    
    st.markdown(f"I, {learner_name} confirm I have attended the scheduled sessions from **16/09/2024** to **20/09/2024** "
                "as outlined in the weekly timetable. I understand that accurate attendance is important for the completion of this programme.")
    
    st.markdown(df2.to_html(index=False, header=False), unsafe_allow_html=True)

    st.subheader("Learner Declaration")
    learner_signature = st.text_input("Signature")
    declaration_date = st.date_input("Date of Declaration", datetime.now().date())
    
    if st.button("Back"):
        st.session_state.page = 1
        st.experimental_rerun()

    if st.button("Save Declaration and Export Document"):
        filled_doc = Document("Skills Boot Camp Week 1 Timesheet.docx")
        
        for paragraph in filled_doc.paragraphs:
            if 'learner_name' in paragraph.text:
                paragraph.text = paragraph.text.replace('learner_name', learner_name)
            if 'learner_signature' in paragraph.text:
                paragraph.text = paragraph.text.replace('learner_signature', learner_signature)
            if 'date' in paragraph.text:
                paragraph.text = paragraph.text.replace('date', declaration_date.strftime("%d/%m/%Y"))
        
        # Generate a unique file name based on the learner's name
        safe_learner_name = re.sub(r'\W+', '_', learner_name)  # Replace non-alphanumeric characters with underscores
        filled_doc_path = f"Filled_Skills_Boot_Camp_Timesheet_{safe_learner_name}.docx"
        filled_doc.save(filled_doc_path)

        # Notify user and provide download link
        st.success("Document filled and saved successfully.")
        with open(filled_doc_path, "rb") as file:
            st.download_button(f"Download Filled Timesheet for {learner_name}", file, filled_doc_path)
