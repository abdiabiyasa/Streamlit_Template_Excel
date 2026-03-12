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
        st.write(duplicate_claims[['ClaimNo']].drop_duplicates())
    return df.drop_duplicates(subset='ClaimNo', keep='last')
    
# prepro function
def move_to_template(df):
    new_df = filter_data(df)
    new_df = keep_last_duplicate(new_df)

    # Convert date columns
    date_columns = ["TreatmentStart", "TreatmentFinish", "Date"]
    for col in date_columns:
        new_df[col] = pd.to_datetime(new_df[col], errors='coerce')
        if new_df[col].isnull().any():
            st.warning(f"Invalid date values in '{col}', coerced to NaT.")

    df_transformed = pd.DataFrame({
        "No": range(1, len(new_df) + 1),
        "Policy No": new_df["PolicyNo"],
        "Client Name": new_df["ClientName"],
        "Claim No": new_df["ClaimNo"],
        "Member No": new_df["MemberNo"],
        "Emp ID": new_df["EmpID"],
        "Emp Name": new_df["EmpName"],
        "Patient Name": new_df["PatientName"],
        "Membership": new_df["Membership"],
        "Product Type": new_df["ProductType"],
        "Claim Type": new_df["ClaimType"],
        "Room Option": new_df["RoomOption"].fillna('').astype(str).str.upper().str.replace(r"\s+", "", regex=True),
        "Area": new_df["Area"],
        "Diagnosis": new_df["PrimaryDiagnosis"].str.upper(),
        "Treatment Place": new_df["TreatmentPlace"].str.upper(),
        "Treatment Start": new_df["TreatmentStart"],
        "Treatment Finish": new_df["TreatmentFinish"],
        "Settled Date": new_df["Date"],
        "Tahun": new_df["Date"].dt.year,
        "Bulan": new_df["Date"].dt.month,
        "Sum of Billed": new_df["Billed"],
        "Sum of Accepted": new_df["Accepted"],
        "Sum of Excess Coy": new_df["ExcessCoy"],
        "Sum of Excess Emp": new_df["ExcessEmp"],
        "Sum of Excess Total": new_df["ExcessTotal"],
        "Sum of Unpaid": new_df["Unpaid"],
    })
    return df_transformed

# compile to an excel workbook
def save_to_excel_a(df, filename: str):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        workbook = writer.book

        # format tampilan
        header_fmt = workbook.add_format({'font_name': 'Aptos', 'font_size': 11, 'bold': True,'align': 'center', 'border': 1})
        border_fmt = workbook.add_format({'font_name': 'Aptos', 'font_size': 11, 'border': 1})
        borderbold_fmt = workbook.add_format({'font_name': 'Aptos', 'font_size': 11, 'bold': True,'border': 1})
        
        # format angka
        num_fmt = workbook.add_format({'font_name': 'Aptos', 'font_size': 11, 'border': 1, 'num_format': '#,##0;[Red]-#,##0;""'})
        date_fmt = workbook.add_format({'font_name': 'Aptos', 'font_size': 11, 'border': 1, 'num_format': 'dd/mm/yyyy'})
        plain_fmt = workbook.add_format({'font_name': 'Aptos', 'font_size': 11})

        # summary sheet
        summary = workbook.add_worksheet("Summary")
        writer.sheets['Summary'] = summary

        summary.hide_gridlines(2)
        summary.write(0, 0, "List Claim", plain_fmt)
        summary.write_formula("A2", "=SC!A2", plain_fmt)
        summary.write_formula("A3", "=SC!A3", plain_fmt)

        metrics = [
            ("Total Claims", len(df["Claim No"])),
            ("Total Billed", df["Sum of Billed"].sum()),
            ("Total Accepted", df["Sum of Accepted"].sum()),
            ("Total Excess", df["Sum of Excess Total"].sum()),
            ("Total Unpaid", df["Sum of Unpaid"].sum()),
        ]
        
        col0_max = 0
        col1_max = 0
        
        for name, val in metrics:
            col0_max = max(col0_max, len(str(name)))
            col1_max = max(col1_max, len(f"{val:,}"))
            
            for i, (name, val) in enumerate(metrics, start=4):
                summary.write(i, 0, name, borderbold_fmt)
                summary.write(i, 1, val, num_fmt)
            
            summary.set_column(0, 0, max(col0_max + 2, 15))
            summary.set_column(1, 1, max(col1_max + 2, 15))

        # sc sheet
        sc = workbook.add_worksheet("SC")
        writer.sheets['SC'] = sc 

        sc.hide_gridlines(2)
        sc.write(0, 0, "List Claim", plain_fmt)
        sc.write(1, 0, df["Client Name"].iloc[0] if not df.empty else "", plain_fmt)
        sc.write(2, 0, "YTD", plain_fmt)
        sc.write(3, 0, "", plain_fmt)

        # table header
        for col_idx, col_name in enumerate(df.columns):
            sc.write(4, col_idx, col_name, header_fmt)


        koma_cols = ["Sum of Billed", "Sum of Accepted", "Sum of Excess Coy",
                     "Sum of Excess Emp", "Sum of Excess Total", "Sum of Unpaid"]
        
        date_cols = ["Treatment Start", "Treatment Finish", "Settled Date"]

        df[koma_cols] = df[koma_cols].replace([np.inf, -np.inf], np.nan)
        df[koma_cols] = df[koma_cols].fillna(0)

        # biar gada (blanks) lbh rapih
        other_cols = [c for c in df.columns if c not in koma_cols and c not in date_cols]
        df[other_cols] = df[other_cols].fillna("")

        for r, row_data in enumerate(df.to_dict("records"), start=5):
            for c, (col_name, val) in enumerate(row_data.items()):
        
                if col_name in koma_cols:
                    if val == 0:
                        sc.write_number(r, c, 0, num_fmt)
                    else:
                        sc.write_number(r, c, float(val), num_fmt)
        
                # write tanggal
                elif col_name in date_cols:
                    if pd.notna(val):
                        sc.write_datetime(r, c, pd.to_datetime(val), date_fmt)
                    else:
                        sc.write(r, c, None, border_fmt) # Tanggal boleh blank murni
        
                # -emp id
                elif col_name == "Emp ID":
                    # hrs text
                    sc.write(r, c, str(val), border_fmt)
        
                else:
                    sc.write(r, c, val, border_fmt)
                        
        # auto width
        for idx, col in enumerate(df.columns):
            series = df[col]
            max_len = max(
                series.astype(str).map(len).max(),
                len(str(col)))
            sc.set_column(idx, idx, max_len + 2)
    
    output.seek(0)
    return output.getvalue(), filename
    
# run function (buat di call di excel page)
def run_a(uploaded_file):
    if uploaded_file is None:
        st.warning("Please upload a CSV file")
        return None
    raw_data = pd.read_csv(uploaded_file)
    transformed_data = move_to_template(raw_data)
    return transformed_data
