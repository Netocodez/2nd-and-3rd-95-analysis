import pandas as pd
import os
from io import BytesIO
from datetime import datetime

emr_df = pd.read_csv("LAMISNMRS.csv", encoding='utf-8')

#Filter and process line list        
columns_to_select = ["S/N", "State", "LGA", "FacilityName", "PEPID", "PatientHospitalNo", "uuid", "Sex", "Current_Age", "Surname", "Firstname", "MaritalStatus", "PhoneNo", "Address", "State_of_Residence", "LGA_of_Residence", "DateConfirmedHIV+", "ARTStartDate", "Pharmacy_LastPickupdate", "DaysOfARVRefill", "CurrentARTRegimen", "NextAppt", "CurrentPregnancyStatus", "CurrentViralLoad", "DateResultReceivedFacility", "Alphanumeric_Viral_Load_Result", "LastDateOfSampleCollection", "Outcomes", "Outcomes_Date", "IIT_Date", "ARTStatus_PreviousQuarter", "CurrentARTStatus", "First_TPT_Pickupdate", "Current_TPT_Received", "PBS_Capturee", "PBS_Capture_Date", "Date_Generated", "CaseManager"]
columns_to_select2 = ["S/N", "State", "LGA", "FacilityName", "PEPID", "PatientHospitalNo", "uuid", "Sex", "Current_Age", "Surname", "Firstname", "MaritalStatus", "PhoneNo", "Address", "State_of_Residence", "LGA_of_Residence", "DateConfirmedHIV+", "ARTStartDate", "Pharmacy_LastPickupdate", "DaysOfARVRefill", "CurrentARTRegimen", "NextAppt", "CurrentPregnancyStatus", "CurrentViralLoad", "DateResultReceivedFacility", "Alphanumeric_Viral_Load_Result", "LastDateOfSampleCollection", "Outcomes", "Outcomes_Date", "IIT_Date", "ARTStatus_PreviousQuarter", "CurrentARTStatus", "First_TPT_Pickupdate", "Current_TPT_Received", "PBS_Capturee", "PBS_Capture_Date", "Date_Generated"]

#function to process line list
def process_Linelist(df, column_name, filter_value, columns_to_select, sort_by=None, ascending=True):
    # Filter the DataFrame
    df_filtered = df[df[column_name] == filter_value].reset_index(drop=True)

    # Sort if sort_by is specified and column(s) exist
    if sort_by:
        if isinstance(sort_by, list):
            valid_sort_cols = [col for col in sort_by if col in df_filtered.columns]
        else:
            valid_sort_cols = [sort_by] if sort_by in df_filtered.columns else []

        if valid_sort_cols:
            df_filtered = df_filtered.sort_values(by=valid_sort_cols, ascending=ascending).reset_index(drop=True)

    # Insert Serial Number
    df_filtered.insert(0, 'S/N', df_filtered.index + 1)

    # Select only existing columns
    existing_columns = [col for col in columns_to_select if col in df_filtered.columns]
    df_selected = df_filtered[existing_columns]

    return df_selected

def clean_id(val):
    return str(val).strip().lower().replace(' ', '').lstrip('0') if pd.notna(val) else ''

def appendLamisData(df, dfbaseline, emr_df):
    # Remove rows with any blank fields in mapping
    emr_df = emr_df[(emr_df != '').all(axis=1)]
    
    # Select and deduplicate necessary columns from emr_df
    emr_subset = emr_df[['Name on NMRS', 'LGA', 'STATE', 'Name on Lamis']].drop_duplicates(subset='Name on NMRS')

    # Merge once using FacilityName <-> Name on NMRS
    df = df.merge(
        emr_subset,
        how='left',
        left_on='FacilityName',
        right_on='Name on NMRS',
        suffixes=('', '_emr')
    )

    # Fill missing LGA and State from EMR
    df['LGA'] = df['LGA'].fillna(df['LGA_emr'])
    df['State'] = df['State'].fillna(df['STATE'])

    # Replace FacilityName if different
    df.loc[df['Name on Lamis'] != df['FacilityName'], 'FacilityName'] = df['Name on Lamis']

    # Drop extra columns
    df.drop(['Name on NMRS', 'LGA_emr', 'STATE', 'Name on Lamis'], axis=1, inplace=True)

    # Normalize hospital numbers and unique IDs
    df['PatientHospitalNo1'] = df['PatientHospitalNo'].apply(clean_id)
    df['PatientUniqueID1'] = df['PEPID'].apply(clean_id)
    dfbaseline['Hospital Number1'] = dfbaseline['Hospital Number'].apply(clean_id)
    dfbaseline['Unique ID1'] = dfbaseline['Unique ID'].apply(clean_id)

    # Create consistent unique identifiers for both datasets
    dfbaseline['unique identifiers'] = (
        dfbaseline["LGA"].astype(str).str.lower().str.strip().str.replace(' ', '') +
        dfbaseline["Facility"].astype(str).str.lower().str.strip().str.replace(' ', '') +
        dfbaseline["Hospital Number1"] +
        dfbaseline["Unique ID1"]
    )

    df['unique identifiers'] = (
        df["LGA"].astype(str).str.lower().str.strip().str.replace(' ', '') +
        df["FacilityName"].astype(str).str.lower().str.strip().str.replace(' ', '') +
        df["PatientHospitalNo1"] +
        df["PatientUniqueID1"]
    )

    # Drop duplicates from baseline data
    dfbaseline = dfbaseline.drop_duplicates(subset=['unique identifiers'], keep=False)

    # Identify duplicates in 'unique identifiers'
    dup_mask = df.duplicated('unique identifiers', keep=False)

    # Only modify duplicates
    df.loc[dup_mask, 'unique identifiers'] = (
        df.loc[dup_mask]
        .groupby('unique identifiers')
        .cumcount()
        .astype(str)
        .radd(df.loc[dup_mask, 'unique identifiers'] + '_')
    )

    # Merge into df
    df = df.merge(
        dfbaseline[['unique identifiers', 'Date of TPT Start (yyyy-mm-dd)', 'TPT Type']],
        on='unique identifiers',
        how='left',
        suffixes=('', '_baseline')
    )

    # Fill missing TPT values
    df['Date of TPT Start (yyyy-mm-dd)'] = pd.to_datetime(df['Date of TPT Start (yyyy-mm-dd)'], errors='coerce', dayfirst=True)
    df['First_TPT_Pickupdate'] = pd.to_datetime(df['First_TPT_Pickupdate'], errors='coerce', dayfirst=True)
    df['First_TPT_Pickupdate'] = df['First_TPT_Pickupdate'].fillna(df['Date of TPT Start (yyyy-mm-dd)'])
    df['Current_TPT_Received'] = df['Current_TPT_Received'].fillna(df['TPT Type'])

    return df


def ensureLGAState(df, emr_df):
    # Remove rows with any blank fields in mapping
    emr_df = emr_df[(emr_df != '').all(axis=1)]
    
    # Select and deduplicate necessary columns from emr_df
    emr_subset = emr_df[['Name on NMRS', 'LGA', 'STATE', 'Name on Lamis']].drop_duplicates(subset='Name on NMRS')

    # Merge once using FacilityName <-> Name on NMRS
    df = df.merge(
        emr_subset,
        how='left',
        left_on='FacilityName',
        right_on='Name on NMRS',
        suffixes=('', '_emr')
    )

    # Fill missing LGA and State from EMR
    df['LGA'] = df['LGA'].fillna(df['LGA_emr'])
    df['State'] = df['State'].fillna(df['STATE'])

    # Drop extra columns
    df.drop(['Name on NMRS', 'LGA_emr', 'STATE', 'Name on Lamis'], axis=1, inplace=True)
    
    return df

def calculate_age_vectorized(df, dob_col='DOB', ref_date=None):
        # pick the reference date
        if ref_date is None:
            today = pd.Timestamp.today().normalize()  # current day
        else:
            today = pd.to_datetime(ref_date) # use provided reference date

        # fully vectorized age calculation
        dob = df[dob_col]
        dob = dob.astype(str).str.strip()
        dob = pd.to_datetime(dob, errors='coerce', infer_datetime_format=True).fillna(pd.to_datetime('1900'))
        age = (today.year - dob.dt.year 
            - ((dob.dt.month > today.month) | 
                ((dob.dt.month == today.month) & (dob.dt.day > today.day))).astype(int))

        return age

def sc_gap_mask(
        df: pd.DataFrame,
        end_date: str | pd.Timestamp,
        last_sample_col: str = "LastDateOfSampleCollection",
        art_start_col: str = "ARTStartDate",
        age_col: str = "Age"
    ) -> pd.Series:

    today = pd.to_datetime(end_date)
    end_of_quarter = today + pd.tseries.offsets.QuarterEnd(0)
    one_year_ago_quarter_end = (today - pd.DateOffset(years=1)) + pd.tseries.offsets.QuarterEnd(0)
    six_months_ago = today - pd.DateOffset(months=6)

    art_start = pd.to_datetime(df[art_start_col], errors="coerce", dayfirst=True)
    sample_date = pd.to_datetime(df[last_sample_col], errors="coerce", dayfirst=True)
    
    # months on ART as at end_of_quarter
    months_on_art = (
        (end_of_quarter.year - art_start.dt.year) * 12 +
        (end_of_quarter.month - art_start.dt.month) -
        (art_start.dt.day > end_of_quarter.day)
    )

    # Adults: no sample or last sample > 12 months ago
    adult_mask = (df[age_col] >= 15) & ((sample_date.isna()) | (sample_date <= one_year_ago_quarter_end))

    # Pediatrics: no sample or last sample > 6 months ago
    ped_mask = (df[age_col] < 15) & ((sample_date.isna()) | (sample_date <= six_months_ago))

    mask = (
        (df['CurrentARTStatus'] == 'Active') &
        (df['ARTStatus_PreviousQuarter'] == 'Active') &
        (months_on_art >= 6) &
        (adult_mask | ped_mask)
    )

    return mask

def export_to_excel_with_formatting(dataframes, formatted_period, summaryName="2ND 95 SUMMARY",
                                    division_columns=None, color_column="%Weekly Refill Rate",
                                    column_widths=None, mergeNum=2, row_masks=None):

    if row_masks is None:
        row_masks = {}

    output = BytesIO()
    writer = pd.ExcelWriter(output, engine="xlsxwriter")

    # ----------------------------
    # Strip datetime columns
    # ----------------------------
    for df_name, df in dataframes.items():
        for col in df.columns:
            if pd.api.types.is_datetime64_any_dtype(df[col]):
                df[col] = df[col].dt.date

    # ----------------------------
    # Process each sheet
    # ----------------------------
    for sheet_name, df in dataframes.items():
        df.to_excel(writer, sheet_name=sheet_name, startrow=1, header=False, index=False)
        workbook = writer.book
        worksheet = writer.sheets[sheet_name]

        # ----------------------------
        # Formats
        # ----------------------------
        header_format = workbook.add_format({
            "bold": True, "text_wrap": True, "valign": "bottom",
            "fg_color": "#D7E4BC", "border": 1
        })
        percentage_format = workbook.add_format({'num_format': '0.00%'})
        percentage_bold = workbook.add_format({
            "bold": True, "text_wrap": True, "valign": "bottom",
            "fg_color": "#D7E4BC", "border": 1, 'num_format': '0.00%'
        })
        row_band_format = workbook.add_format({'bg_color': '#F9F9F9'})
        alert_format = workbook.add_format({'bg_color': '#FFA500'})
        mask_format = workbook.add_format({'bg_color': '#FFEB9C'})
        title_format = workbook.add_format({'bold': True, 'font_size': 14, 'align': 'center'})

        # ----------------------------
        # Write headers
        # ----------------------------
        for col_num, value in enumerate(df.columns):
            worksheet.write(0, col_num, value, header_format)

        # ----------------------------
        # Summary Sheet
        # ----------------------------
        if sheet_name == summaryName:
            # Title
            worksheet.merge_range(0, 0, 0, len(df.columns)-1, f'{summaryName} AS AT {formatted_period}', title_format)
            # Rewrite headers/data starting row 2
            df.to_excel(writer, sheet_name=sheet_name, startrow=2, header=False, index=False)
            for col_num, value in enumerate(df.columns):
                worksheet.write(1, col_num, value, header_format)

            # Color scale
            color_columns = color_column if isinstance(color_column, (list, tuple)) else [color_column]
            for col in color_columns:
                if col in df.columns and pd.api.types.is_numeric_dtype(df[col]):
                    idx = df.columns.get_loc(col)
                    worksheet.conditional_format(
                        2, idx, len(df)+1, idx,
                        {'type': '3_color_scale', 'min_color': '#F8696B', 'mid_color': '#FFEB84', 'max_color': '#63BE7B'}
                    )

            # Row banding
            worksheet.conditional_format(2, 0, len(df)+1, len(df.columns)-1,
                                         {'type': 'formula', 'criteria': 'MOD(ROW(),2)=0', 'format': row_band_format})
            worksheet.hide_gridlines(2)

            # Column widths
            if column_widths:
                for col_range, width in column_widths.items():
                    worksheet.set_column(col_range, width)

            # Division formatting
            if division_columns is None:
                division_columns = {}
            for col in division_columns.keys():
                worksheet.set_column(f"{col}:{col}", None, percentage_format)

            # Total row
            total_row = len(df)+2
            header_format.set_align('center')
            worksheet.merge_range(total_row, 0, total_row, mergeNum, 'Total', header_format)
            for col_num in range(1, len(df.columns)):
                col_letter = chr(65 + col_num)
                if col_letter in division_columns:
                    num_col, denom_col = division_columns[col_letter]
                    worksheet.write_formula(total_row, col_num,
                                            f"SUM({num_col}3:{num_col}{total_row})/SUM({denom_col}3:{denom_col}{total_row})",
                                            percentage_bold)
                else:
                    worksheet.write_formula(total_row, col_num, f"SUM({col_letter}3:{col_letter}{total_row})", header_format)

        # ----------------------------
        # Non-summary Sheets
        # ----------------------------
        else:
            first_col = 0
            last_col = len(df.columns)-1

            # Notes column for comments
            notes_col_name = "PatientHospitalNo"
            if notes_col_name in df.columns:
                notes_idx = df.columns.get_loc(notes_col_name)
                for row_idx in range(len(df)):
                    alerts = []

                    # Biometrics alert
                    if "PBS_Capture_Date" in df.columns:
                        val = df.at[row_idx, "PBS_Capture_Date"]
                        if pd.isna(val) or val == "":
                            alerts.append("Needs biometrics capture")

                    # Mask alert
                    if sheet_name in row_masks and row_idx < len(row_masks[sheet_name]) and row_masks[sheet_name][row_idx]:
                        alerts.append("Due for sample collection")

                    if alerts:
                        text = " | ".join(alerts)
                        worksheet.write_comment(row_idx+1, notes_idx, text, {'visible': False, 'x_scale': 1.5, 'y_scale': 1.5})
                        # Highlight only rows with comments
                        worksheet.conditional_format(row_idx+1, notes_idx, row_idx+1, notes_idx,
                                                     {'type': 'no_errors', 'format': alert_format})

            # Highlight PBS_Capturee / PBS_Capture_Date
            if "PBS_Capture_Date" in df.columns and "PBS_Capturee" in df.columns:
                date_idx = df.columns.get_loc("PBS_Capture_Date")
                capturee_idx = df.columns.get_loc("PBS_Capturee")
                for row_idx in range(len(df)):
                    date_val = df.at[row_idx, "PBS_Capture_Date"]
                    if pd.isna(date_val) or date_val == "":
                        worksheet.conditional_format(row_idx+1, date_idx, row_idx+1, date_idx, {'type':'no_errors','format':alert_format})
                        worksheet.conditional_format(row_idx+1, capturee_idx, row_idx+1, capturee_idx, {'type':'no_errors','format':alert_format})
                        # Add comment for capturee
                        name = df.at[row_idx, "Surname"] if "Surname" in df.columns else ""
                        worksheet.write_comment(row_idx+1, capturee_idx,
                                                f"{name} needs to be captured for biometrics",
                                                {'visible': False, 'x_scale': 1.5, 'y_scale': 1.5})

            # Mask highlighting
            if sheet_name in row_masks:
                mask = row_masks[sheet_name]
                for row_idx, flag in enumerate(mask):
                    if flag:
                        worksheet.conditional_format(row_idx+1, first_col, row_idx+1, last_col,
                                                    {'type': 'formula', 'criteria': '=TRUE', 'format': mask_format})

            # Row banding
            worksheet.conditional_format(1, 0, len(df), len(df.columns)-1,
                                         {'type': 'formula', 'criteria': 'MOD(ROW(),2)=0', 'format': row_band_format})

            # Column widths
            col_ranges = {'O:AF':10, 'J:M':12, 'N:N':40, 'E:F':15}
            for col_range, width in col_ranges.items():
                worksheet.set_column(col_range, width)
            worksheet.set_column('G:G', None, None, {'hidden': True})

    # ----------------------------
    # Save workbook
    # ----------------------------
    writer.close()
    output.seek(0)

    output_dir = os.path.join(os.path.dirname(__file__), "..", "outputs")
    os.makedirs(output_dir, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"{summaryName}_{timestamp}.xlsx"
    output_path = os.path.join(output_dir, filename)
    with open(output_path, "wb") as f:
        f.write(output.getbuffer())

    return filename
