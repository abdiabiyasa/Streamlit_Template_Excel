# excel_d.py
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
    if "ClaimNo" in df_benefit.columns:
        df_benefit = df_benefit[df_benefit["ClaimNo"].isin(df_sc["Claim No"])]
    elif "Claim No" in df_benefit.columns:
        df_benefit = df_benefit[df_benefit["Claim No"].isin(df_sc["Claim No"])]

    return df_benefit

def template_sc(df):
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
        "Plan": new_df["PPlan"],
        "PrePost": new_df["isPrePost2"],
        "Primary Diagnosis": new_df["PrimaryDiagnosis"].str.upper(),
        "Secondary Diagnosis": new_df["SecondaryDiagnosis"].fillna('').str.upper(),
        "Treatment Place": new_df["TreatmentPlace"].str.upper(),
        "Treatment Start": new_df["TreatmentStart"],
        "Treatment Finish": new_df["TreatmentFinish"],
        "Treatment Year": new_df["TreatmentStart"].dt.year,
        "Treatment Month": new_df["TreatmentStart"].dt.month,
        "Settled Date": new_df["Date"],
        "Settled Year": new_df["Date"].dt.year,
        "Settled Month": new_df["Date"].dt.month,
        "Length of Stay": new_df["LOS"],
        "Sum of Billed": new_df["Billed"],
        "Sum of Accepted": new_df["Accepted"],
        "Sum of Excess Coy": new_df["ExcessCoy"],
        "Sum of Excess Emp": new_df["ExcessEmp"],
        "Sum of Excess Total": new_df["ExcessTotal"],
        "Sum of Unpaid": new_df["Unpaid"],
    })
    return df_transformed
    
# prepro Benefit sheet    
def template_benefit(df):
    df.columns = df.columns.str.strip()

    for col in df.columns:
        if df[col].dtype == "object":
            df[col] = df[col].astype(str).str.strip()

        # Rename columns
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
    
        # Clean Room Option and Treatment Room Class
        if "Room Option" in df.columns:
            df["Room Option"] = df["Room Option"].fillna('').astype(str).str.replace(r"\s+", "", regex=True)
        if "Treatment Room Class" in df.columns:
            df["Treatment Room Class"] = df["Treatment Room Class"].fillna('')
    
        return df.drop(columns=["Status_Claim", "BAmount"], errors='ignore')


def save_to_excel_d(df_sc, df_benefit, claim_ratio_df, filename: str):
    import pandas as pd
    from io import BytesIO

    # Helper: normalize dataframe column names (strip)
    def _norm_cols(df):
        df = pd.DataFrame(df).copy()
        df.columns = [c.strip() for c in df.columns]
        return df

    df_sc = _norm_cols(df_sc)
    df_benefit = _norm_cols(df_benefit)
    cr = _norm_cols(claim_ratio_df)

    # Map CR columns
    def _map_cr_columns(cr_df):
        col_map = {}
        policy = next((c for c in cr_df.columns if c.lower() in ("policy no","policyno","policy")), None)
        if policy:
            col_map['Policy No'] = policy
        comp = next((c for c in cr_df.columns if c.lower() in ("company","client name","company name","insurer")), None)
        if comp:
            col_map['Company'] = comp
        net = next((c for c in cr_df.columns if c.lower() in ("net premi","net premium","netpremi")), None)
        if net:
            col_map['Net Premi'] = net
        est = next((c for c in cr_df.columns if c.lower() in ("est claim total","estclaimtotal","est claim","est_claim_total")), None)
        if est:
            col_map['Est Claim Total'] = est
        return col_map

    cr_map = _map_cr_columns(cr)
    # Ensure canonical columns exist in CR
    if 'Company' in cr_map:
        cr = cr.rename(columns={cr_map['Company']: 'Company'})
    else:
        cr['Company'] = ''
    if 'Policy No' in cr_map:
        cr = cr.rename(columns={cr_map['Policy No']: 'Policy No'})
    else:
        cr['Policy No'] = ''
    if 'Net Premi' in cr_map:
        cr = cr.rename(columns={cr_map['Net Premi']: 'Net Premi'})
    else:
        cr['Net Premi'] = 0
    if 'Est Claim Total' in cr_map:
        cr = cr.rename(columns={cr_map['Est Claim Total']: 'Est Claim Total'})
    else:
        cr['Est Claim Total'] = 0

    # Clean numeric in CR defensively
    cr['Net Premi'] = pd.to_numeric(
        cr['Net Premi'].astype(str).str.replace('%','',regex=False).str.replace(',','',regex=False),
        errors='coerce'
    ).fillna(0)
    cr['Est Claim Total'] = pd.to_numeric(
        cr['Est Claim Total'].astype(str).str.replace('%','',regex=False).str.replace(',','',regex=False),
        errors='coerce'
    ).fillna(0)

    # Normalize SC columns and map expected names
    def _map_sc_columns(sc_df):
        mapping_candidates = {
            'Client Name': ['Client Name','ClientName','Client','Company'],
            'Policy No': ['Policy No','PolicyNo','Policy'],
            'Sum of Billed': ['Sum of Billed','Billed','Total Billed'],
            'Sum of Accepted': ['Sum of Accepted','Accepted','Claim'],
            'Sum of Unpaid': ['Sum of Unpaid','Unpaid'],
            'Sum of Excess Total': ['Sum of Excess Total','Excess Total','ExcessTotal'],
            'Sum of Excess Coy': ['Sum of Excess Coy','Excess Coy','ExcessCoy'],
            'Sum of Excess Emp': ['Sum of Excess Emp','Excess Emp','ExcessEmp']
        }
        col_map = {}
        for canon, candidates in mapping_candidates.items():
            found = next((c for c in sc_df.columns if c in candidates or c.strip().lower() in [x.lower() for x in candidates]), None)
            col_map[canon] = found if found else None
        return col_map

    sc_map = _map_sc_columns(df_sc)

    # Ensure canonical columns exist in df_sc (create if missing)
    for canon, orig in sc_map.items():
        if orig is None:
            if canon in ('Client Name','Policy No'):
                df_sc[canon] = ''
            else:
                df_sc[canon] = 0
        else:
            df_sc = df_sc.rename(columns={orig: canon})

    # Defensive conversions for df_sc
    df_sc['Client Name'] = df_sc['Client Name'].astype(str).fillna('')
    df_sc['Policy No'] = df_sc['Policy No'].astype(str).fillna('')
    for c in ['Sum of Billed','Sum of Accepted','Sum of Unpaid','Sum of Excess Total','Sum of Excess Coy','Sum of Excess Emp']:
        if c in df_sc.columns:
            df_sc[c] = pd.to_numeric(df_sc[c].astype(str).str.replace(',','',regex=False), errors='coerce').fillna(0)
        else:
            df_sc[c] = 0

    # -------------------------
    # AGGREGATE SC
    # -------------------------
    sc_grouped = df_sc.groupby(['Client Name','Policy No'], dropna=False).agg({
        'Sum of Billed':'sum',
        'Sum of Accepted':'sum',
        'Sum of Unpaid':'sum',
        'Sum of Excess Total':'sum',
        'Sum of Excess Coy':'sum',
        'Sum of Excess Emp':'sum'
    }).reset_index().rename(columns={'Sum of Accepted':'Claim'})

    # -------------------------
    # MERGE with CR
    # -------------------------
    merged = cr.merge(
        sc_grouped,
        left_on=['Company','Policy No'],
        right_on=['Client Name','Policy No'],
        how='left',
        suffixes=('','_sc')
    )

    # Ensure merged numeric
    for col in ['Sum of Billed','Sum of Unpaid','Sum of Excess Total','Sum of Excess Coy','Sum of Excess Emp','Claim']:
        merged[col] = pd.to_numeric(merged.get(col, 0), errors='coerce').fillna(0)

    merged['Billed'] = merged['Sum of Billed']
    merged['Unpaid'] = merged['Sum of Unpaid']
    merged['Excess Total'] = merged['Sum of Excess Total']
    merged['Excess Coy'] = merged['Sum of Excess Coy']
    merged['Excess Emp'] = merged['Sum of Excess Emp']

    merged['Net Premi'] = pd.to_numeric(merged.get('Net Premi', 0), errors='coerce').fillna(0)
    merged['Est Claim Total'] = pd.to_numeric(merged.get('Est Claim Total', 0), errors='coerce').fillna(0)

    merged['CR'] = merged.apply(lambda r: (r['Claim'] / r['Net Premi'] * 100) if r['Net Premi'] else 0, axis=1)
    merged['Est CR'] = merged.apply(lambda r: (r['Est Claim Total'] / r['Net Premi'] * 100) if r['Net Premi'] else 0, axis=1)

    cr_columns_header = ["Company","Net Premi","Billed","Unpaid","Excess Total","Excess Coy","Excess Emp","Claim","CR","Est CR"]
    for c in cr_columns_header:
        if c not in merged.columns:
            merged[c] = 0

    # totals
    grand = {
        'Net Premi': merged['Net Premi'].sum(),
        'Est Claim Total': merged['Est Claim Total'].sum(),
        'Billed': merged['Billed'].sum(),
        'Unpaid': merged['Unpaid'].sum(),
        'Excess Total': merged['Excess Total'].sum(),
        'Excess Coy': merged['Excess Coy'].sum(),
        'Excess Emp': merged['Excess Emp'].sum(),
        'Claim': merged['Claim'].sum()
    }
    grand_cr = (grand['Claim']/grand['Net Premi']*100) if grand['Net Premi'] else 0
    grand_est_cr = (grand['Est Claim Total']/grand['Net Premi']*100) if grand['Net Premi'] else 0

    # Excel
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book

        # formats
        header_fmt = workbook.add_format({'font_name':'Aptos','font_size':11,'bold':True,'align':'center','border':1})
        border_fmt = workbook.add_format({'font_name':'Aptos','font_size':11,'border':1})
        borderbold_fmt = workbook.add_format({'font_name':'Aptos','font_size':11,'bold':True,'border':1})
        num_fmt = workbook.add_format({'font_name':'Aptos','font_size':11,'border':1,'num_format':'#,##0;[Red]-#,##0;""'})
        date_fmt = workbook.add_format({'font_name':'Aptos','font_size':11,'border':1,'num_format':'dd/mm/yyyy'})
        plain_fmt = workbook.add_format({'font_name':'Aptos','font_size':11})
        plain_border = workbook.add_format({'border':1,'font_name':'Aptos','num_format':'#,##0;[Red]-#,##0;"-";@'})
        bold_plain_border = workbook.add_format({'bold':True,'border':1,'font_name':'Aptos','num_format':'#,##0;[Red]-#,##0;"-";@'})
        header_border = workbook.add_format({'bold':True,'border':1,'align':'center','font_name':'Aptos','num_format':'#,##0;[Red]-#,##0;"-";@'})
        highlight_yellow = workbook.add_format({'bg_color':'#FFFF00','border':1,'num_format':'0.00"%"','font_name':'Aptos'})
        highlight_yellow_bold = workbook.add_format({'bg_color':'#FFFF00','border':1,'bold':True,'num_format':'0.00"%"','font_name':'Aptos'})
        percent_format = workbook.add_format({'border': 1, 'num_format': '0.00"%"', 'font_name': 'Aptos'})
        boolean_format = workbook.add_format({'border': 1, 'font_name': 'Aptos', 'num_format': '"TRUE";;"FALSE"' })

        # Summary
        summary_sheet = workbook.add_worksheet('Summary')
        writer.sheets['Summary'] = summary_sheet
        summary_sheet.hide_gridlines(2)

        summary_sheet.write(0,0,'List Claim', plain_fmt)
        summary_sheet.write_formula('A2','=SC!A2')
        summary_sheet.write_formula('A3','=SC!A3')

        metrics = [
            ("Total Claims", len(df_sc), num_fmt),
            ("Employee Claims", len(df_sc[df_sc.get('Membership','') == '1. EMP']), num_fmt),
            ("Spouse Claims", len(df_sc[df_sc.get('Membership','') == '2. SPO']), num_fmt),
            ("Children Claims", len(df_sc[df_sc.get('Membership','') == '3. CHI']), num_fmt),
            ("Total Billed", df_sc['Sum of Billed'].sum() if 'Sum of Billed' in df_sc.columns else 0, num_fmt),
            ("Total Accepted", df_sc['Sum of Accepted'].sum() if 'Sum of Accepted' in df_sc.columns else 0, num_fmt),
            ("Total Excess", df_sc['Sum of Excess Total'].sum() if 'Sum of Excess Total' in df_sc.columns else 0, num_fmt),
            ("Total Unpaid", df_sc['Sum of Unpaid'].sum() if 'Sum of Unpaid' in df_sc.columns else 0, num_fmt),
            ("Claim Ratio (%)", grand_cr, percent_format)
        ]

        for i,(name,val,fmt) in enumerate(metrics,start=4):
            summary_sheet.write(i,0,name,borderbold_fmt)
            summary_sheet.write(i,1,val,fmt)

        cr_start = 4 + len(metrics) + 3
        for ci,col_name in enumerate(cr_columns_header):
            summary_sheet.write(cr_start,ci,col_name,header_border)
        r = cr_start + 1

        if not merged.empty:
            for _, rowdata in merged.iterrows():
                for ci, col_name in enumerate(cr_columns_header):
                    val = rowdata.get(col_name, 0)

                    if col_name in ('CR', 'Est CR'):
                        summary_sheet.write_number(r, ci, float(val), highlight_yellow)
                    elif col_name in ('Net Premi','Est Claim Total','Billed','Unpaid','Excess Total','Excess Coy','Excess Emp','Claim'):
                        try:
                            numeric_val = float(val) if pd.notna(val) else 0
                        except Exception:
                            numeric_val = 0
                        summary_sheet.write_number(r, ci, numeric_val, num_fmt)
                    else:
                        summary_sheet.write(r, ci, val, plain_border)
                r += 1
        else:
            summary_sheet.write(r,0,'No Claim Ratio data',plain_border)
            r += 1

        # Grand total
        summary_sheet.write(r,0,'Grand Total',bold_plain_border)
        for ci,col_name in enumerate(cr_columns_header[1:],start=1):
            if col_name == 'CR':
                summary_sheet.write_number(r,ci,grand_cr,highlight_yellow_bold)
            elif col_name == 'Est CR':
                summary_sheet.write_number(r,ci,grand_est_cr,highlight_yellow_bold)
            else:
                v = grand.get(col_name, '')
                if v == '' or pd.isna(v):
                    summary_sheet.write(r,ci,None,bold_plain_border)
                else:
                    try:
                        summary_sheet.write_number(r,ci,float(v),bold_plain_border)
                    except:
                        summary_sheet.write(r,ci,v,bold_plain_border)
        r += 1

        # SC sheet — unchanged
        sc_sheet = workbook.add_worksheet('SC')
        writer.sheets['SC'] = sc_sheet
        sc_sheet.hide_gridlines(2)
        sc_sheet.write(0,0,'List Claim', plain_fmt)
        sc_sheet.write(1,0, df_sc['Client Name'].iloc[0] if not df_sc.empty else '', plain_fmt)
        sc_sheet.write(2,0,'YTD', plain_fmt)

        for ci,col_name in enumerate(df_sc.columns):
            sc_sheet.write(4,ci,col_name,header_fmt)

        koma_cols = ['Sum of Billed','Sum of Accepted','Sum of Excess Coy','Sum of Excess Emp','Sum of Excess Total','Sum of Unpaid']

        for rr, rowdata in enumerate(df_sc.to_dict("records"), start=5):
            for ci, (col_name, val) in enumerate(rowdata.items()):
        
                # koma cols ( 0 -> cell kosong)
                if col_name in koma_cols:
                    if pd.isna(val) or val == 0:
                        sc_sheet.write(rr, ci, None, num_fmt)  # blank cell
                    else:
                        sc_sheet.write_number(rr, ci, float(val), num_fmt)
        
                # Date columns
                elif col_name in ('Treatment Start', 'Treatment Finish', 'Settled Date'):
                    if pd.notna(val) and val != '':
                        try:
                            sc_sheet.write_datetime(rr, ci, pd.to_datetime(val), date_fmt)
                        except:
                            sc_sheet.write(rr, ci, None, border_fmt)
                    else:
                        sc_sheet.write(rr, ci, None, border_fmt)
        
                # PrePost write as boolean
                elif col_name == 'PrePost':
                    if val in [1, "1", True]:
                        sc_sheet.write(rr, ci, "True", border_fmt)
                    elif val in [0, "0", False, None]:
                        sc_sheet.write(rr, ci, "False", border_fmt)
                    else:
                        sc_sheet.write(rr, ci, None, border_fmt)
        
                # Emp ID write as text
                elif col_name == 'Emp ID':
                    sc_sheet.write(rr, ci, str(val) if pd.notna(val) else "", border_fmt)
        
                # dll klo value 0 -> cell jd kosong
                else:
                    if pd.isna(val) or val == 0:
                        sc_sheet.write(rr, ci, None, border_fmt)
                    else:
                        sc_sheet.write(rr, ci, val, border_fmt)


        # Benefit sheet
        benefit_sheet = workbook.add_worksheet('Benefit')
        writer.sheets['Benefit'] = benefit_sheet
        benefit_sheet.hide_gridlines(2)

        for ci,col_name in enumerate(df_benefit.columns):
            benefit_sheet.write(0,ci,col_name,header_fmt)

        koma_cols_benefit = ['Billed','Accepted','Unpaid','Excess Total','Excess Coy','Excess Emp']

        for rr, rowdata in enumerate(df_benefit.to_dict("records"), start=1):
            for ci, (col_name, val) in enumerate(rowdata.items()):
        
                # koma kols (0 -> cell kosong)
                if col_name in koma_cols_benefit:
                    if pd.isna(val) or val == 0:
                        benefit_sheet.write(rr, ci, None, num_fmt)  # blank cell
                    else:
                        benefit_sheet.write_number(rr, ci, float(val), num_fmt)
        
                # Date columns
                elif col_name in ('Treatment Start', 'Treatment Finish', 'Payment Date'):
                    if pd.notna(val) and val != '':
                        try:
                            benefit_sheet.write_datetime(rr, ci, pd.to_datetime(val), date_fmt)
                        except:
                            benefit_sheet.write(rr, ci, None, border_fmt)
                    else:
                        benefit_sheet.write(rr, ci, None, border_fmt)
        
                # Emp ID write as text
                elif col_name == 'Emp ID':
                    benefit_sheet.write(rr, ci, str(val) if pd.notna(val) else "", border_fmt)
        
                # dll klo value 0 -> cell jd kosong
                else:
                    if pd.isna(val) or val == 0:
                        benefit_sheet.write(rr, ci, None, border_fmt)
                    else:
                        benefit_sheet.write(rr, ci, val, border_fmt)


        # Autofit SC & Benefit only
        def autofit(sheet, df):
            for idx,col in enumerate(df.columns):
                series = df[col].astype(str)
                try:
                    max_len = max(series.map(len).max(), len(col))
                except Exception:
                    max_len = len(col)
                sheet.set_column(idx, idx, max_len + 5)

        try:
            autofit(sc_sheet, df_sc)
            autofit(benefit_sheet, df_benefit)
            autofit(summary_sheet, pd.DataFrame(merged[cr_columns_header]))
        except Exception:
            pass

    output.seek(0)
    return output.getvalue(), filename
    
# compile to excel main
def run_d(uploaded_sc, uploaded_benefit, uploaded_cr, policy_filter_list):
    # load SC and Benefit CSVs
    df_sc_raw = pd.read_csv(uploaded_sc)
    df_benefit_raw = pd.read_csv(uploaded_benefit)

    # load Claim Ratio (xlsx)
    try:
        df_cr_raw = pd.read_excel(uploaded_cr)
    except Exception as e:
        st.error(f"Error reading Claim Ratio file: {e}")
        df_cr_raw = pd.DataFrame()

    # prepro SC and benefit using excel_c logic
    df_sc_clean = template_sc(df_sc_raw)

    # filter SC by policy list if provided
    if policy_filter_list:
        # ensure comparable types
        df_sc_clean["Policy No"] = df_sc_clean["Policy No"].astype(str).str.strip()
        df_sc_clean = df_sc_clean[df_sc_clean["Policy No"].isin([str(p).strip() for p in policy_filter_list])]

    # filter benefit using df_sc_clean (reuse filter logic)
    df_benefit_filtered = filter_benefit_data(df_benefit_raw, df_sc_clean)
    df_benefit_clean = template_benefit(df_benefit_filtered)

    # Process claim ratio: filter by policy no (assume CR has column 'Policy No' or 'PolicyNo')
    if not df_cr_raw.empty:
        cr_cols = [c for c in df_cr_raw.columns if c.strip().lower() in ("policy no", "policyno", "policy")]
        if cr_cols:
            policy_col = cr_cols[0]
            df_cr_raw[policy_col] = df_cr_raw[policy_col].astype(str).str.strip()
            if policy_filter_list:
                df_cr_filtered = df_cr_raw[df_cr_raw[policy_col].isin([str(p).strip() for p in policy_filter_list])]
            else:
                df_cr_filtered = df_cr_raw.copy()
        else:
            # no policy column found, keep whole df but warn
            st.warning("No 'Policy No' column detected in Claim Ratio file. Using full CR dataset.")
            df_cr_filtered = df_cr_raw.copy()
    else:
        df_cr_filtered = pd.DataFrame()

    # normalize column names in claim ratio (strip)
    df_cr_filtered.columns = df_cr_filtered.columns.str.strip()

    return df_sc_clean, df_benefit_clean, df_cr_filtered

