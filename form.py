import streamlit as st
import pandas as pd
import os
from datetime import datetime
from io import BytesIO

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle
)


# ============================================================
# CONFIGURATION
# ============================================================

EXCEL_FILE = "responses.xlsx"

st.set_page_config(
    page_title="Online Form",
    page_icon="📝",
    layout="centered"
)


# ============================================================
# SAVE RESPONSE TO EXCEL
# ============================================================

def save_to_excel(data):
    df_new = pd.DataFrame([data])

    if os.path.exists(EXCEL_FILE):
        try:
            df_old = pd.read_excel(EXCEL_FILE)
            df_final = pd.concat([df_old, df_new], ignore_index=True)
        except Exception:
            df_final = df_new
    else:
        df_final = df_new

    df_final.to_excel(EXCEL_FILE, index=False)


# ============================================================
# CREATE PDF
# ============================================================

def create_pdf(data):

    buffer = BytesIO()

    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=40,
        leftMargin=40,
        topMargin=40,
        bottomMargin=40
    )

    elements = []

    styles = getSampleStyleSheet()

    title = Paragraph(
        "<b>FORM RESPONSE</b>",
        styles["Title"]
    )

    elements.append(title)
    elements.append(Spacer(1, 20))

    table_data = [
        ["Field", "Response"]
    ]

    for key, value in data.items():
        table_data.append([
            str(key),
            str(value)
        ])

    table = Table(
        table_data,
        colWidths=[180, 300]
    )

    table.setStyle(
        TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.grey),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),

            ("GRID", (0, 0), (-1, -1), 0.5, colors.black),

            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),

            ("FONTNAME", (0, 1), (0, -1), "Helvetica-Bold"),

            ("VALIGN", (0, 0), (-1, -1), "TOP"),

            ("LEFTPADDING", (0, 0), (-1, -1), 8),
            ("RIGHTPADDING", (0, 0), (-1, -1), 8),
            ("TOPPADDING", (0, 0), (-1, -1), 8),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
        ])
    )

    elements.append(table)

    elements.append(Spacer(1, 25))

    footer = Paragraph(
        f"Generated on: {datetime.now().strftime('%d-%m-%Y %I:%M %p')}",
        styles["Normal"]
    )

    elements.append(footer)

    doc.build(elements)

    buffer.seek(0)

    return buffer


# ============================================================
# HEADER
# ============================================================

st.title("📝 Online Form")

st.write("Please fill in the details below and choose an option at the end.")


# ============================================================
# FORM INPUTS
# ============================================================

with st.form("online_form"):

    name = st.text_input(
        "Full Name *"
    )

    employee_id = st.text_input(
        "Employee ID *"
    )

    mobile = st.text_input(
        "Mobile Number *"
    )

    email = st.text_input(
        "Email Address"
    )

    department = st.selectbox(
        "Department *",
        [
            "Select Department",
            "Sales",
            "Service",
            "Marketing",
            "HR",
            "Accounts",
            "Other"
        ]
    )

    response_date = st.date_input(
        "Date"
    )

    remarks = st.text_area(
        "Remarks"
    )

    st.markdown("---")

    st.write("### Select an Option")

    option = st.radio(
        "What would you like to do?",
        [
            "Submit Response",
            "Download PDF"
        ],
        horizontal=True
    )

    action = st.form_submit_button(
        "Continue",
        use_container_width=True
    )


# ============================================================
# PROCESS ACTION
# ============================================================

if action:

    # VALIDATION
    if not name.strip():
        st.error("Please enter your Full Name.")

    elif not employee_id.strip():
        st.error("Please enter your Employee ID.")

    elif not mobile.strip():
        st.error("Please enter your Mobile Number.")

    elif department == "Select Department":
        st.error("Please select a Department.")

    else:

        # CREATE DATA
        form_data = {
            "Timestamp": datetime.now().strftime(
                "%d-%m-%Y %I:%M:%S %p"
            ),
            "Full Name": name,
            "Employee ID": employee_id,
            "Mobile Number": mobile,
            "Email": email,
            "Department": department,
            "Date": response_date.strftime(
                "%d-%m-%Y"
            ),
            "Remarks": remarks
        }


        # ====================================================
        # DOWNLOAD PDF OPTION
        # ====================================================

        if option == "Download PDF":

            pdf_file = create_pdf(form_data)

            st.success(
                "Your PDF is ready."
            )

            st.download_button(
                label="📥 Download Completed Form as PDF",
                data=pdf_file,
                file_name=f"Form_{employee_id}.pdf",
                mime="application/pdf",
                use_container_width=True
            )


        # ====================================================
        # SUBMIT OPTION
        # ====================================================

        elif option == "Submit Response":

            try:

                save_to_excel(form_data)

                st.success(
                    "✅ Your response has been submitted successfully!"
                )

                st.info(
                    "Your response has been saved in responses.xlsx"
                )

            except Exception as e:

                st.error(
                    f"Error while saving the response: {e}"
                )