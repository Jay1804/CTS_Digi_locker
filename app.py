"""Streamlit UI for the Digilocker BGV report."""
import io
from datetime import date

import streamlit as st

import main as report

st.set_page_config(page_title="Digilocker BGV Report", layout="wide")
st.title("Digilocker BGV Report")

c1, c2 = st.columns(2)
start = c1.date_input("Check created from", value=date.fromisoformat(report.START_DATE))
end = c2.date_input("Check created till", value=date.today())

if st.button("Generate report", type="primary"):
    try:
        with st.spinner("Fetching data from database..."):
            raw = report.fetch(str(start), str(end))
            st.session_state["raw_count"] = len(raw)
            st.session_state["reports"] = report.split_reports(raw)
            bgv = raw["BGVType"].fillna("").str.strip().str.lower()
            st.session_state["skipped"] = int((~bgv.isin(["lateral", "campus"])).sum())
            st.session_state["stamp"] = end.strftime("%d%b%Y")
    except Exception as e:  # DB/connection errors
        st.session_state.pop("reports", None)
        st.error(f"Failed to fetch data: {e}")

reports = st.session_state.get("reports")
if reports:
    st.success(
        f"Fetched {st.session_state['raw_count']} rows "
        f"({st.session_state['skipped']} skipped: BGVType not Lateral/Campus)."
    )
    tabs = st.tabs(list(reports))
    for tab, (name, df) in zip(tabs, reports.items()):
        with tab:
            st.caption(f"{len(df)} rows")
            status = df[report.STATUS_COL].value_counts().rename_axis("Status").reset_index(name="Count")
            st.dataframe(status, hide_index=True)
            st.dataframe(report.display_frame(df), hide_index=True)
            buf = io.BytesIO()
            report.write_excel(df, buf)
            st.download_button(
                f"Download {name} Excel",
                buf.getvalue(),
                file_name=f"Digilocker_{name}_{st.session_state['stamp']}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"dl_{name}",
            )
