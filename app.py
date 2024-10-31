import streamlit as st
from datetime import datetime, date
from docx import Document
from docx.shared import Inches
import numpy as np
import pandas as pd
import re
from PIL import Image as PILImage
from io import BytesIO
from streamlit_drawable_canvas import st_canvas

# Set page configuration with a favicon
st.set_page_config(
    page_title="Skills Boot Camp Weekly Timesheet",
    page_icon="https://lirp.cdn-website.com/d8120025/dms3rep/multi/opt/social-image-88w.png", 
    layout="centered"  # "centered" or "wide"
)

# Initialize session state for screen navigation
if 'page' not in st.session_state:
    st.session_state.page = 1
if 'learner_signature' not in st.session_state: 
    st.session_state.learner_signature = None
if 'declaration_date' not in st.session_state: 
    st.session_state.declaration_date = None    
if 'checkboxes' not in st.session_state: 
    st.session_state.checkboxes = []  # Store checkbox states for each attendance row

# Function to load DOCX data, skipping header row in the second table
def load_docx_data():
    doc = Document("Skills Boot Camp Week 1 Timesheet.docx")
    day, session_activity, facilitator, time, notes_comments = [], [], [], [], []
    attendance_data = []  # for second table
    
    for table_idx, table in enumerate(doc.tables):
        for row_idx, row in enumerate(table.rows):
            cells = [cell.text.replace('\n', ' ').strip() for cell in row.cells]
            if table_idx == 0 and len(cells) == 5:
                day.append(cells[0])
                session_activity.append(cells[1])
                facilitator.append(cells[2])
                time.append(cells[3])
                notes_comments.append(cells[4])
            elif table_idx == 1 and len(cells) >= 4:
                if row_idx == 0:
                    continue  # Skip the header row
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

def is_signature_drawn(signature):
    if signature is None or not isinstance(signature, np.ndarray) or np.all(signature == 255):
        return False
    return True

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

# Second Screen: Learner Declaration and Attendance Table with Checkboxes
elif st.session_state.page == 2:    
    st.subheader("Attendance Register Declaration (Monday - Friday)")
    st.session_state.learner_name = st.text_input("Enter your full name")    
    st.markdown(f"I, {st.session_state.learner_name} confirm I have attended the scheduled sessions from **16/09/2024** to **20/09/2024** "
                "as outlined in the weekly timetable. I understand that accurate attendance is important for the completion of this programme.")
    
    # Custom header without checkboxes
    st.markdown("### Attendance Table")
    header_cols = st.columns([1, 1, 1, 1, 1])
    header_cols[0].write("Day")
    header_cols[1].write("Date")
    header_cols[2].write("Arrival Time")
    header_cols[3].write("Departure Time")
    header_cols[4].write("Learner Signature")

    # Track checkbox states in session state to retain selections
    checkboxes = []
    for idx, row in df2.iterrows():
        row_cols = st.columns([1, 1, 1, 1, 1])
        row_cols[0].write(row["Day"])
        row_cols[1].write(row["Date"])
        row_cols[2].write(row["Arrival Time"])
        row_cols[3].write(row["Departure Time"])
        # Add a checkbox for each row in the "Learner Signature" column
        checked = row_cols[4].checkbox("Present", key=f"checkbox_{idx}")
        checkboxes.append(checked)
    
    # Update checkboxes in session state
    st.session_state.checkboxes = checkboxes

    # Signature Section
    st.subheader("Learner Declaration")
    st.write("I confirm that the information above is correct and that my attendance has been accurately recorded for this week.")
    st.text("Please draw your signature below:")

    # Signature Box
    canvas_result = st_canvas(
        fill_color="rgba(255, 255, 255, 1)",  
        stroke_width=5,
        stroke_color="rgb(0, 0, 0)",
        background_color="white",
        width=400,
        height=150,
        drawing_mode="freedraw",
        key="canvas",
    )
    st.session_state.learner_signature = canvas_result.image_data

    declaration_date = date.today().strftime("%d-%m-%Y")
    st.write(f"Date: **{declaration_date}**")    

    if st.button("Save Declaration and Export Document"):
        if is_signature_drawn(st.session_state.learner_signature) and st.session_state.learner_name:
            filled_doc = Document("Skills Boot Camp Week 1 Timesheet.docx")
            
            for paragraph in filled_doc.paragraphs:
                if 'learner_name' in paragraph.text:
                    paragraph.text = paragraph.text.replace('learner_name', st.session_state.learner_name)
                if 'date' in paragraph.text:
                    paragraph.text = paragraph.text.replace('date', declaration_date)
                # Replace the placeholder text with the signature image
                if 'learner_signature' in paragraph.text:
                    paragraph.text = paragraph.text.replace('learner_signature', "")
                    
                    # Create a new run within the paragraph to add the image
                    run = paragraph.add_run()
                    signature_image = PILImage.fromarray(st.session_state.learner_signature.astype('uint8'), 'RGBA')
                    signature_image_path = "learner_signature.png"
                    signature_image.save(signature_image_path)
                    
                    # Insert the image into the document in place of the placeholder
                    run.add_picture(signature_image_path, width=Inches(2))  # Adjust width as needed

            # Save the signature as an image file
            signature_image = PILImage.fromarray(st.session_state.learner_signature.astype('uint8'), 'RGBA')
            signature_image_path = "learner_signature.png"
            signature_image.save(signature_image_path)

            # Insert signature image in the "Learner Signature" cell for each checked row
            for table_idx, table in enumerate(filled_doc.tables):
                if table_idx == 1:  # Ensure we're modifying the attendance table
                    for row_idx, row in enumerate(table.rows[1:]):  # Skip the header row
                        cell = row.cells[4]  # "Learner Signature" column
                        if st.session_state.checkboxes[row_idx]:  # If checkbox was checked
                            # Clear text and add image
                            cell.text = ""
                            run = cell.paragraphs[0].add_run()
                            run.add_picture(signature_image_path, width=Inches(0.5))  # Adjust width as needed
                        else:
                            cell.text = ""  # Clear any placeholder text if unchecked

            # Generate a unique file name based on the learner's name
            safe_learner_name = re.sub(r'\W+', '_', st.session_state.learner_name)
            filled_doc_path = f"Filled_Skills_Boot_Camp_Timesheet_{safe_learner_name}.docx"
            filled_doc.save(filled_doc_path)

            st.success("Document filled and saved successfully.")
            with open(filled_doc_path, "rb") as file:
                st.download_button(f"Download Filled Timesheet for {st.session_state.learner_name}", file, filled_doc_path)
        else:
            st.warning("Please enter your nae & draw the signature!")
    if st.button("Back"):
        st.session_state.page = 1
        st.experimental_rerun()
