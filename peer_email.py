import streamlit as st
import pandas as pd
import random
from datetime import datetime
from io import BytesIO
import os
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# -------------------------------------------------
# Streamlit Page Config
# -------------------------------------------------
st.set_page_config(
    page_title="GSCE - Peer to Peer Duties Assignment",
    layout="wide"
)

st.image("gitm.png", width=150)
st.title("GSCE - Peer to Peer Duties Assignment")

st.markdown("""
This system generates **weekly peer duty assignments**  
and **emails peer faculty** with details of where they need to report.
""")

# -------------------------------------------------
# Excel File Path
# -------------------------------------------------
FILE_PATH = "Peer_Job_Fixedslots_withoutsecondperson.xlsx"

if not os.path.exists(FILE_PATH):
    st.error("Required Excel file not found.")
    st.stop()

st.success("Excel file loaded successfully.")

# -------------------------------------------------
# Day Selection
# -------------------------------------------------
days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]
selected_day = st.selectbox("Select Day", days)

# -------------------------------------------------
# Email Function (Peer Faculty)
# -------------------------------------------------
def send_peer_email(
    to_email,
    peer_name,
    subject,
    teaching_faculty,
    day,
    time_slot,
    week
):
    smtp_server = st.secrets["SMTP_SERVER"]
    smtp_port = st.secrets["SMTP_PORT"]
    sender_email = st.secrets["EMAIL_ADDRESS"]
    sender_password = st.secrets["EMAIL_PASSWORD"]
    institute = st.secrets["INSTITUTE_NAME"]

    msg = MIMEMultipart()
    msg["From"] = sender_email
    msg["To"] = to_email
    msg["Subject"] = f"Peer Duty Assignment – {day} ({time_slot})"

    body = f"""
Dear {peer_name},

This is to inform you that you have been assigned **peer duty**
as per the GSCE peer allocation for the week {week}.

Assignment Details:
--------------------------------
Subject            : {subject}
Reporting Faculty  : {teaching_faculty}
Day                : {day}
Time Slot          : {time_slot}
--------------------------------

Kindly report to the respective class/lab on time.

If there is any genuine difficulty, inform the coordinator immediately.

Regards,
GSCE Peer Duty Coordination Team
{institute}
"""

    msg.attach(MIMEText(body, "plain"))

    with smtplib.SMTP(smtp_server, smtp_port) as server:
        server.starttls()
        server.login(sender_email, sender_password)
        server.send_message(msg)

# -------------------------------------------------
# Generate Assignment
# -------------------------------------------------
if st.button("Generate / Regenerate Day-wise Assignment"):
    with st.spinner("Generating assignment..."):

        # Load sheets
        peerslots = pd.read_excel(FILE_PATH, sheet_name="Peerslots")
        busy_fac = pd.read_excel(FILE_PATH, sheet_name="Busy_fac")

        # Filter free peers for selected day
        peerslots = peerslots[
            (peerslots["Status"].str.lower() == "free") &
            (peerslots["Day"] == selected_day)
        ].copy()

        if peerslots.empty:
            st.warning(f"No free peers found for {selected_day}")
            st.stop()

        # Deterministic weekly seed
        week_seed = datetime.now().strftime("%Y-%U")
        random.seed(f"{week_seed}-{selected_day}")

        assigned_subjects = []
        assigned_faculty = []

        weekly_assigned_subjects = set()

        for _, peer in peerslots.iterrows():
            time_slot = peer["Time Slot"]
            peer_emp_id = peer["Emp ID"]

            possible_subjects = busy_fac[
                (busy_fac["Day"] == selected_day) &
                (busy_fac["Time Slot"] == time_slot) &
                (busy_fac["Emp ID"] != peer_emp_id) &
                (~busy_fac["Subject"].isin(weekly_assigned_subjects))
            ]

            if not possible_subjects.empty:
                chosen = possible_subjects.sample(1).iloc[0]

                assigned_subjects.append(chosen["Subject"])
                assigned_faculty.append(chosen["Faculty Name"])
                weekly_assigned_subjects.add(chosen["Subject"])
            else:
                assigned_subjects.append("No Subject Available")
                assigned_faculty.append("NA")

        # Update dataframe
        peerslots["Assigned Subject"] = assigned_subjects
        peerslots["Teaching Faculty"] = assigned_faculty

        # Store for email step
        st.session_state["assignment"] = peerslots
        st.session_state["week"] = week_seed

        st.success(f"{selected_day} assignment generated for Week {week_seed}")
        st.dataframe(peerslots, use_container_width=True)

        # Download Excel
        output = BytesIO()
        peerslots.to_excel(output, index=False, engine="openpyxl")
        output.seek(0)

        st.download_button(
            label="Download Assignment Excel",
            data=output,
            file_name=f"Peer_Assignment_{selected_day}_Week_{week_seed}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# -------------------------------------------------
# Send Emails to Peer Faculty
# -------------------------------------------------
if "assignment" in st.session_state:
    if st.button("Send Email Instructions to Peer Faculty"):
        sent = 0
        failed = 0

        df = st.session_state["assignment"]
        week_seed = st.session_state["week"]

        for _, row in df.iterrows():
            if row["Assigned Subject"] != "No Subject Available":
                try:
                    send_peer_email(
                        to_email=row["Peer Email"],
                        peer_name=row["Peer Name"],
                        subject=row["Assigned Subject"],
                        teaching_faculty=row["Teaching Faculty"],
                        day=row["Day"],
                        time_slot=row["Time Slot"],
                        week=week_seed
                    )
                    sent += 1
                except Exception as e:
                    failed += 1
                    st.error(f"Email failed for {row['Peer Name']}: {e}")

        st.success(f"Emails sent successfully: {sent} | Failed: {failed}")
