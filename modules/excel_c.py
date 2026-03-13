import pandas as pd
import numpy as np
import streamlit as st
from io import BytesIO

# data cleaning function
def filter_data(df):
    return df[df['ClaimStatus'] == 'R']

def keep_last_duplicate(df):
    duplicate_claims = df[df.duplicated(subset='ClaimNo', keep=False)]
    if not duplicate_claims.empty:
        st.write("Duplicated ClaimNo values:")
        st.write(duplicate_claims)

    df = df.drop_duplicates(subset='ClaimNo', keep='last')
    return df

def filter_benefit_data(df_benefit, df_sc):
    df_benefit = df_benefit.copy()
    df_benefit.columns = df_benefit.columns.str.strip()

    # Filter status claim R
    if 'Status_Claim' in df_benefit.columns:
        df_benefit = df_benefit[df_benefit['Status_Claim'] == 'R']
    elif 'Status Claim' in df_benefit.columns:
        df_benefit = df_benefit[df_benefit['Status Claim'] == 'R']
    else:
        st.warning("Column 'Status Claim' not found. Data not filtered.")

    # Filter claim no supaya hanya yang ada di SC
    if "ClaimNo" in df_benefit.columns and "Claim No" in df_sc.columns:
        df_benefit = df_benefit[df_benefit["ClaimNo"].isin(df_sc["Claim No"])]
    elif "Claim No" in df_benefit.columns and "Claim No" in df_sc.columns:
        df_benefit = df_benefit[df_benefit["Claim No"].isin(df_sc["Claim No"])]

    return df_benefit

# prepro SC Sheet
def template_sc(df):
    new_df = filter_data(df)
    new_df = keep_last_duplicate(new_df)

    # Convert date columns
    date_columns = ["TreatmentStart", "TreatmentFinish", "Date"]
    for col in date_columns:
        if col in new_df.columns:
            new_df[col] = pd.to_datetime(new_df[col], errors='coerce')
            if new_df[col].isnull().any():
                st.warning(f"Invalid date values in '{col}', coerced to NaT.")

    df_transformed = pd.DataFrame({
        "No": range(1, len(new_df) + 1),
        "Policy No": new_df.get("PolicyNo"),
        "Client Name": new_df.get("ClientName"),
        "Claim No": new_df.get("ClaimNo"),
        "Member No": new_df.get("MemberNo"),
        "Emp ID": new_df.get("EmpID"),
        "Emp Name": new_df.get("EmpName"),
        "Patient Name": new_df.get("PatientName"),
        "Membership": new_df.get("Membership"),
        "Product Type": new_df.get("ProductType"),
        "Claim Type": new_df.get("ClaimType"),
        "Room Option": new_df.get("RoomOption", "").fillna('').astype(str).str.upper().str.replace(r"\s+", "", regex=True),
        "Area": new_df.get("Area"),
        "Diagnosis": new_df.get("PrimaryDiagnosis", "").astype(str).str.upper(),
        "Treatment Place": new_df.get("TreatmentPlace", "").astype(str).str.upper(),
        "Treatment Start": new_df.get("TreatmentStart"),
        "Treatment Finish": new_df.get("TreatmentFinish"),
        "Settled Date": new_df.get("Date"),
        "Year": new_df.get("Date").dt.year if "Date" in new_df.columns else None,
        "Month": new_df.get("Date").dt.month if "Date" in new_df.columns else None,
        "Sum of Billed": new_df.get("Billed"),
        "Sum of Accepted": new_df.get("Accepted"),
        "Sum of Excess Coy": new_df.get("ExcessCoy"),
        "Sum of Excess Emp": new_df.get("ExcessEmp"),
        "Sum of Excess Total": new_df.get("ExcessTotal"),
        "Sum of Unpaid": new_df.get("Unpaid"),
    })

    return df_transformed

# prepro Benefit sheet    
def template_benefit(df):
    df = df.copy()
    df.columns = df.columns.str.strip()

    for col in df.columns:
        if df[col].dtype == "object":
            df[col] = df[col].astype(str).str.strip()

    rename_mapping = {
        'ClientName': 'Client Name',
        'PolicyNo': 'Policy No',
        'ClaimNo': 'Claim No',
        'MemberNo': 'Member No',
        'PatientName': 'Patient Name',
        'EmpID': 'Emp ID',
        'EmpName': 'Emp Name',
        'ClaimType': 'Claim Type',
        'TreatmentPlace': 'Treatment Place',
        'RoomOption': 'Room Option',
        'TreatmentRoomClass': 'Treatment Room Class',
        'TreatmentStart': 'Treatment Start',
        'TreatmentFinish': 'Treatment Finish',
        'ProductType': 'Product Type',
        'BenefitName': 'Benefit Name',
        'PaymentDate': 'Payment Date',
        'ExcessTotal': 'Excess Total',
        'ExcessCoy': 'Excess Coy',
        'ExcessEmp': 'Excess Emp'
    }

    df = df.rename(columns=rename_mapping)

    date_cols = ["Treatment Start", "Treatment Finish", "Payment Date"]
    for col in date_cols:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors='coerce')

    if "Room Option" in df.columns:
        df["Room Option"] = df["Room Option"].fillna('').astype(str).str.replace(r"\s+", "", regex=True)
    if "Treatment Room Class" in df.columns:
        df["Treatment Room Class"] = df["Treatment Room Class"].fillna('')

    return df.drop(columns=["Status_Claim", "BAmount"], errors='ignore')

# compile to an excel workbook
def save_to_excel_c(df_sc, df_benefit, filename: str):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        workbook = writer.book

        #formats
        header_fmt = workbook.add_format({'font_name': 'Aptos', 'font_size': 11, 'bold': True,'align': 'center', 'border': 1})
        border_fmt = workbook.add_format({'font_name': 'Aptos', 'font_size': 11, 'border': 1})
        borderbold_fmt = workbook.add_format({'font_name': 'Aptos', 'font_size': 11, 'bold': True,'border': 1})
        num_fmt = workbook.add_format({'font_name': 'Aptos', 'font_size': 11, 'border': 1, 'num_format': '#,##0;[Red]-#,##0;""'})
        date_fmt = workbook.add_format({'font_name': 'Aptos', 'font_size': 11, 'border': 1, 'num_format': 'dd/mm/yyyy'})
        plain_fmt = workbook.add_format({'font_name': 'Aptos', 'font_size': 11})

        # Summary sheet
        summary = workbook.add_worksheet("Summary")
        writer.sheets['Summary'] = summary
        summary.hide_gridlines(2)

        summary.write(0, 0, "List Claim", plain_fmt)
        summary.write_formula("A2", "=SC!A2", plain_fmt)
        summary.write_formula("A3", "=SC!A3", plain_fmt)

        metrics = [
            ("Total Claims", len(df_sc)),
            ("Total Billed", df_sc.get("Sum of Billed").sum()),
            ("Total Accepted", df_sc.get("Sum of Accepted").sum()),
            ("Total Excess", df_sc.get("Sum of Excess Total").sum()),
            ("Total Unpaid", df_sc.get("Sum of Unpaid").sum()),
        ]
        
        col0_max, col1_max = 0, 0

        for name, val in metrics:
            col0_max = max(col0_max, len(str(name)))
            col1_max = max(col1_max, len(f"{val:,}"))

        for i, (name, val) in enumerate(metrics, start=4):
            summary.write(i, 0, name, borderbold_fmt)
            summary.write(i, 1, val, num_fmt)

        summary.set_column(0, 0, col0_max + 2)
        summary.set_column(1, 1, col1_max + 2)

        # SC sheet
        sc = workbook.add_worksheet("SC")
        writer.sheets['SC'] = sc
        sc.hide_gridlines(2)

        sc.write(0, 0, "List Claim", plain_fmt)
        sc.write(1, 0, df_sc["Client Name"].iloc[0] if not df_sc.empty else "", plain_fmt)
        sc.write(2, 0, "YTD", plain_fmt)

        #Header
        for col_idx, col_name in enumerate(df_sc.columns):
            sc.write(4, col_idx, col_name, header_fmt)

        koma_cols = ["Sum of Billed", "Sum of Accepted", "Sum of Excess Coy",
                     "Sum of Excess Emp", "Sum of Excess Total", "Sum of Unpaid"]

        for r, row_data in enumerate(df_sc.to_dict("records"), start=5):
            for c, (col_name, val) in enumerate(row_data.items()):

                if col_name in ["Treatment Start", "Treatment Finish", "Settled Date"] and pd.notna(val):
                    sc.write_datetime(r, c, val, date_fmt)

                elif col_name in koma_cols:
    try:
        if pd.isna(val) or val in [0, "0", "", None]:
            sc.write(r, c, None, border_fmt)
        else:
            if isinstance(val, str):
                clean = val.replace("Rp", "").replace(",", "").strip()
                if clean == "":
                    sc.write(r, c, None, border_fmt)
                else:
                    sc.write_number(r, c, float(clean), num_fmt)
            else:
                try:
                    num = float(str(val).replace(",", ""))
                    sc.write_number(r, c, num, num_fmt)
                except (ValueError, TypeError):
                    sc.write(r, c, val, border_fmt)
    except Exception:
        sc.write_string(r, c, str(val))

                elif col_name == "Emp ID":
                    sc.write(r, c, str(val) if pd.notna(val) else "", border_fmt)

                else:
                    sc.write(r, c, val if pd.notna(val) and val != 0 else "", border_fmt)

        # Benefit sheet
        benefit = workbook.add_worksheet("Benefit")
        writer.sheets['Benefit'] = benefit
        benefit.hide_gridlines(2)

        for col_idx, col_name in enumerate(df_benefit.columns):
            benefit.write(0, col_idx, col_name, header_fmt)

        for r, row_data in enumerate(df_benefit.to_dict("records"), start=1):
            for c, (col_name, val) in enumerate(row_data.items()):
                benefit.write(r, c, val if pd.notna(val) and val != 0 else "", border_fmt)

    output.seek(0)
    return output.getvalue(), filename

# run for main excel    
def run_c(uploaded_sc, uploaded_benefit):
    df_sc_raw = pd.read_csv(uploaded_sc)
    df_benefit_raw = pd.read_csv(uploaded_benefit)

    df_sc_clean = template_sc(df_sc_raw)
    df_benefit_filtered = filter_benefit_data(df_benefit_raw, df_sc_clean)
    df_benefit_clean = template_benefit(df_benefit_filtered)

    return df_sc_clean, df_benefit_clean
