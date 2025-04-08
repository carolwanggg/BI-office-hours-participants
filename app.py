import random
import pandas as pd
import numpy as np
import streamlit as st
import datetime as dt
import altair as alt
import io
from azure.storage.blob import BlobServiceClient

# ========================= #
#  AZURE BLOB STORAGE SETUP #
# ========================= #

sas_url = st.secrets.blob_credentials.sas_url  # Only the full SAS URL is needed
container_name = "bicollections"
blob_file_name = "fulfillment/bi_office_hours.xlsx"

blob_service_client = BlobServiceClient(account_url=sas_url)
container_client = blob_service_client.get_container_client(container=container_name)

def read_from_blob_storage(file_name):
    blob_client = container_client.get_blob_client(blob=file_name)
    downloaded_blob = blob_client.download_blob().readall()
    return downloaded_blob

def prepare_xlsx(tuples_list):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        for df, sheet_name in tuples_list:
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    return output.getvalue()

def upload_to_blob_storage(tuples_list, file_name):
    file_data = prepare_xlsx(tuples_list)
    blob_client = container_client.get_blob_client(blob=file_name)
    blob_client.upload_blob(file_data, overwrite=True)

# ===================== #
#  STREAMLIT FUNCTIONS  #
# ===================== #

@st.cache_data
def get_data():
    """Loads participants data from an Azure Blob Excel file."""
    blob_file = read_from_blob_storage(blob_file_name)
    participants_df = pd.read_excel(blob_file, sheet_name="participants")
    session_df = pd.read_excel(blob_file, sheet_name="session_history")
    session_df["date"] = pd.to_datetime(session_df["date"], format="%Y-%m-%d").dt.date
    return participants_df, session_df

def get_next_friday(today):
    days_until_next_friday = (4 - today.weekday()) % 7
    next_friday = today + dt.timedelta(days=days_until_next_friday)
    if (next_friday - dt.date(2024, 1, 5)).days % 14 != 0:
        next_friday += dt.timedelta(days=14)
    return next_friday

def get_next_participants(df, available_team, threshold):
    prev_participants = df.unique().tolist()[:threshold]
    next_participants = random.sample([p for p in available_team if p not in prev_participants], 2)
    return next_participants

def add_participants_entry(df, next_participants, next_date):
    insert_row = {"date": next_date, "participant_1": next_participants[0], "participant_2": next_participants[1]}
    df = pd.concat([df, pd.DataFrame([insert_row])], ignore_index=True)
    return df

# =============== #
#  STREAMLIT APP  #
# =============== #

st.set_page_config(page_title="BI Office Hours Participants", page_icon="📊", layout="wide")


participants_df, session_df = get_data()

today = dt.date.today()
next_friday_default = get_next_friday(today)

selectbox_page = st.selectbox(
    "",
    ["🗕 Assign Participants", "📜 Session History", "➕ Add Participants"],
    label_visibility="collapsed",
)

if selectbox_page == "🗕 Assign Participants":
    participants = participants_df["participant"][participants_df["is_active"] == True].tolist()
    available_team = st.multiselect("Who is available to participate?", participants, participants)
    next_date = st.date_input("Next Office Hours Date", next_friday_default)

    if st.button("Assign Participants!"):
        st.cache_data.clear()
        if len(available_team) < 2:
            st.warning("Please select at least two team members.")
        else:
            next_participants = get_next_participants(
                pd.concat([session_df["participant_1"], session_df["participant_2"]]) if not session_df.empty else pd.Series([]), 
                available_team, 
                threshold=3
            )
            if st.checkbox("Save Results", True):
                session_df = add_participants_entry(session_df, next_participants, next_date)
                upload_to_blob_storage([(participants_df, "participants"), (session_df, "session_history")], blob_file_name)
            st.success(f"Next Office Hours Participants: {next_participants[0]} & {next_participants[1]}")

    st.markdown("## Assignment History 📜")

    if session_df.empty:
        st.write("No sessions recorded yet.")
    else:
        st.table(session_df.sort_values("date", ascending=False))

        st.markdown("### Participation Frequency 📊")
        melted_df = session_df.melt(id_vars=["date"], value_vars=["participant_1", "participant_2"], var_name="Role", value_name="Participant")
        frequency_df = melted_df.groupby("Participant", as_index=False).size().rename(columns={"size": "count"})

        chart = alt.Chart(frequency_df).mark_bar().encode(
            x=alt.X("Participant", sort="-y", title="Participant"),
            y=alt.Y("count", title="Number of Participations"),
            color="Participant",
            tooltip=["Participant", "count"]
        ).properties(width=700, height=400)

        st.altair_chart(chart, use_container_width=True)

elif selectbox_page == "📜 Session History":
    st.markdown("""
        <p style='text-align: center; font-size: 25px; color: #FFB000'><b>Session History</b></p>
    """, unsafe_allow_html=True)
    if session_df.empty:
        st.write("No sessions recorded yet.")
    else:
        st.table(session_df.sort_values("date", ascending=False))

elif selectbox_page == "➕ Add Participants":
    st.markdown("""
        <p style='text-align: center; font-size: 25px; color: #FFB000'><b>Participants</b></p>
        <p style='font-size: 16px; color: #072543'> <span style='font-size: 20px'><b>How To Use This Page</b></span>
        <br>– This page gives an overview of the current and previous members
        <br>– It is best to select if a current participant is active or not rather than delete them, to keep history
        <br>– To <span style='color: #FFB000'><b>add</b></span> a new participant
        <br>&emsp;&emsp; – Click on the <span style='color: #FFB000'><b>+</b></span> at the bottom row
        <br>&emsp;&emsp; – Type their name and check <span style='color: #FFB000'><b>isActive</b></span>
        <br>– To <span style='color: #FFB000'><b>remove</b></span> a participant (not recommended)
        <br>&emsp;&emsp; – Click on the <span style='color: #FFB000'><b>checkbox</b></span> in the leftmost column of their row
        <br>&emsp;&emsp; – Press <span style='color: #FFB000'><b>delete/backspace</b></span> on your keyboard
        <br>– Click on <span style='color: #FFB000'><b>Save</b></span>
        </p>
    """, unsafe_allow_html=True)
    edited_df = st.data_editor(
        participants_df.rename(
            columns={"participant": "Participant", "is_active": "isActive"}
        ),
        num_rows="dynamic",
    )
    if st.button("Save"):
        participants_df = edited_df.rename(
            columns={"Participant": "participant", "isActive": "is_active"}
        )
        participants_df.sort_values("participant", inplace=True, ignore_index=True)
        upload_to_blob_storage([(participants_df, "participants"), (session_df, "session_history")], blob_file_name)
        st.success("Participants have been saved!")
        st.cache_data.clear()