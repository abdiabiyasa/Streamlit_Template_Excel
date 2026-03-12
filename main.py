import streamlit as st
import numpy as np
from modules.excel_a import run_a, save_to_excel_a
from modules.excel_b import run_b, save_to_excel_b
from modules.excel_c import run_c, save_to_excel_c
from modules.excel_d import run_d, save_to_excel_d
from modules.excel_e import run_e, save_to_excel_e
from modules.excel_f import run_f, save_to_excel_f

st.title("Excel Template")

# keep pilihan di station head biar abis download ga langsung balik ke awal step
if "option" not in st.session_state:
    st.session_state.option = "SC"

option = st.selectbox(
    "Choose standar:",
    ["SC", "SC w/Payment Date", "SC + Benefit", "SC (RWC)", "SC + Benefit (NFB)", "Template Report A2000"],
    key="option")

# module sc
if st.session_state.option == "SC":
    uploaded_sc = st.file_uploader("Upload SC file", type=["csv"], key="uploaded_sc_a")

    if uploaded_sc:
        transformed_data = run_a(uploaded_sc)

        if transformed_data is not None:
            st.write("Preview Transformed Data:")
            st.dataframe(transformed_data.head())

            filename = st.text_input("Enter Excel filename:", "SC - - YTD", key="fname_a")
            if filename:
                excel_bytes, fname = save_to_excel_a(transformed_data, filename + ".xlsx")
                st.download_button(
                    label="Download Excel",
                    data=excel_bytes,
                    file_name=fname,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="download_a"
                )

#template: SC + Payment Date
elif st.session_state.option == "SC w/Payment Date":
    uploaded_sc = st.file_uploader("Upload SC file", type=["csv"], key="uploaded_sc_b")

    if uploaded_sc:
        transformed_data = run_b(uploaded_sc)

        if transformed_data is not None:
            st.write("Preview Transformed Data:")
            st.dataframe(transformed_data.head())

            filename = st.text_input("Enter Excel filename:", "SC - - YTD", key="fname_b")
            if filename:
                excel_bytes, fname = save_to_excel_b(transformed_data, filename + ".xlsx")
                st.download_button(
                    label="Download Excel",
                    data=excel_bytes,
                    file_name=fname,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="download_b"
                )

# template: SC + Benefit
elif st.session_state.option == "SC + Benefit":

    # upload files needed
    uploaded_sc = st.file_uploader("Upload SC file", type=["csv"], key="uploaded_sc_c")
    uploaded_benefit = st.file_uploader("Upload Benefit file", type=["csv"], key="uploaded_benefit_c")

    if uploaded_sc and uploaded_benefit:
        df_sc, df_benefit = run_c(uploaded_sc, uploaded_benefit)

        st.write("Preview SC:")
        st.dataframe(df_sc.head())

        st.write("Preview Benefit:")
        st.dataframe(df_benefit.head())

        filename = st.text_input("Enter Excel filename:", "SC & Benefit - - YTD", key="fname_c")

        if filename:
            df_sc = df_sc.replace([np.nan, np.inf, -np.inf], "")
            df_benefit = df_benefit.replace([np.nan, np.inf, -np.inf], "")

            excel_bytes, fname = save_to_excel_c(df_sc, df_benefit, filename + ".xlsx")
            st.download_button(
                label="Download Excel",
                data=excel_bytes,
                file_name=fname,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="download_c"
            )

elif st.session_state.option == "SC (RWC)":
    policy_input = st.text_area("Enter Policy Numbers (one per line):", placeholder="Enter one policy per line")
    policy_filter_list = [p.strip() for p in policy_input.splitlines() if p.strip()]

    uploaded_sc = st.file_uploader("Upload SC file (csv)", type=["csv"], key="uploaded_sc_d")
    uploaded_benefit = st.file_uploader("Upload Benefit file (csv)", type=["csv"], key="uploaded_benefit_d")
    uploaded_cr = st.file_uploader("Upload Claim Ratio file (xlsx)", type=["xlsx", "xls"], key="uploaded_cr_d")

    if uploaded_sc and uploaded_benefit and uploaded_cr:
        df_sc, df_benefit, df_cr = run_d(
        uploaded_sc,
        uploaded_benefit,
        uploaded_cr,
        policy_filter_list
    )

        st.subheader("Preview SC")
        st.dataframe(df_sc.head())

        st.subheader("Preview Benefit")
        st.dataframe(df_benefit.head())

        filename = st.text_input("Enter Excel filename:", "SC - - YTD", key="fname_d")

        if filename:
            # sanitize
            df_sc = df_sc.replace([np.nan, np.inf, -np.inf], "")
            df_benefit = df_benefit.replace([np.nan, np.inf, -np.inf], "")
            df_cr = df_cr.replace([np.nan, np.inf, -np.inf], "")

            excel_bytes, fname = save_to_excel_d(
                df_sc,
                df_benefit,
                df_cr,
                filename + ".xlsx"
            )

            st.download_button(
                label="Download Excel",
                data=excel_bytes,
                file_name=fname,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="download_d"
            )

elif st.session_state.option == "SC + Benefit (NFB)":
    # upload files needed
    uploaded_sc = st.file_uploader("Upload SC file", type=["csv"], key="uploaded_sc_e")
    uploaded_benefit = st.file_uploader("Upload Benefit file", type=["csv"], key="uploaded_benefit_e")

    if uploaded_sc and uploaded_benefit:
        df_sc, df_benefit = run_e(uploaded_sc, uploaded_benefit)

        st.write("Preview SC:")
        st.dataframe(df_sc.head())

        st.write("Preview Benefit:")
        st.dataframe(df_benefit.head())

        filename = st.text_input("Enter Excel filename:", "SC & Benefit - - YTD", key="fname_e")

        if filename:
            df_sc = df_sc.replace([np.nan, np.inf, -np.inf], "")
            df_benefit = df_benefit.replace([np.nan, np.inf, -np.inf], "")

            excel_bytes, fname = save_to_excel_e(df_sc, df_benefit, filename + ".xlsx")
            st.download_button(
                label="Download Excel",
                data=excel_bytes,
                file_name=fname,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="download_e"
            )

# template: A2000
elif st.session_state.option == "Template Report A2000":

    # upload files needed
    uploaded_sc = st.file_uploader("Upload SC file", type=["csv"], key="uploaded_sc_f")
    uploaded_benefit = st.file_uploader("Upload Benefit file", type=["csv"], key="uploaded_benefit_f")

    if uploaded_sc and uploaded_benefit:
        df_sc, df_benefit = run_f(uploaded_sc, uploaded_benefit)

        st.write("Preview SC:")
        st.dataframe(df_sc.head())

        st.write("Preview Benefit:")
        st.dataframe(df_benefit.head())

        filename = st.text_input("Enter Excel filename:", "SC & Benefit - - YTD", key="fname_f")

        if filename:
            df_sc = df_sc.replace([np.nan, np.inf, -np.inf], "")
            df_benefit = df_benefit.replace([np.nan, np.inf, -np.inf], "")

            excel_bytes, fname = save_to_excel_f(df_sc, df_benefit, filename + ".xlsx")
            st.download_button(
                label="Download Excel",
                data=excel_bytes,
                file_name=fname,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="download_f"
            )
