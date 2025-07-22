import pandas as pd
import os
from io import BytesIO
from datetime import datetime

emr_df = pd.read_csv("LAMISNMRS.csv", encoding='utf-8')

#Filter and process line list        
columns_to_select = ["S/N", "State", "LGA", "FacilityName", "PEPID", "PatientHospitalNo", "uuid", "Sex", "Current_Age", "Surname", "Firstname", "MaritalStatus", "PhoneNo", "Address", "State_of_Residence", "LGA_of_Residence", "DateConfirmedHIV+", "ARTStartDate", "Pharmacy_LastPickupdate", "DaysOfARVRefill", "CurrentARTRegimen", "NextAppt", "CurrentPregnancyStatus", "CurrentViralLoad", "DateResultReceivedFacility", "Alphanumeric_Viral_Load_Result", "LastDateOfSampleCollection", "Outcomes", "Outcomes_Date", "IIT_Date", "CurrentARTStatus", "First_TPT_Pickupdate", "Current_TPT_Received", "PBS_Capturee", "PBS_Capture_Date", "Date_Generated", "CaseManager"]
columns_to_select2 = ["S/N", "State", "LGA", "FacilityName", "PEPID", "PatientHospitalNo", "uuid", "Sex", "Current_Age", "Surname", "Firstname", "MaritalStatus", "PhoneNo", "Address", "State_of_Residence", "LGA_of_Residence", "DateConfirmedHIV+", "ARTStartDate", "Pharmacy_LastPickupdate", "DaysOfARVRefill", "CurrentARTRegimen", "NextAppt", "CurrentPregnancyStatus", "CurrentViralLoad", "DateResultReceivedFacility", "Alphanumeric_Viral_Load_Result", "LastDateOfSampleCollection", "Outcomes", "Outcomes_Date", "IIT_Date", "CurrentARTStatus", "First_TPT_Pickupdate", "Current_TPT_Received", "PBS_Capturee", "PBS_Capture_Date", "Date_Generated"]

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

def export_to_excel_with_formatting(dataframes, formatted_period, summaryName="2ND 95 SUMMARY", division_columns=None, color_column="%Weekly Refill Rate", column_widths=None, mergeNum=2):
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine="xlsxwriter")
    
    # Strip time from datetime columns to ensure clean yyyy-mm-dd export
    for df_name, df_data in dataframes.items():
        for col in df_data.columns:
            if pd.api.types.is_datetime64_any_dtype(df_data[col]):
                df_data[col] = df_data[col].dt.date

    for sheet_name, dataframe in dataframes.items():
        dataframe.to_excel(writer, sheet_name=sheet_name, startrow=1, header=False, index=False)

        workbook = writer.book
        worksheet = writer.sheets[sheet_name]

        # Header format
        header_format = workbook.add_format({
            "bold": True,
            "text_wrap": True,
            "valign": "bottom",
            "fg_color": "#D7E4BC",
            "border": 1,
        })

        # Percentage formats
        percentage_format = workbook.add_format({'num_format': '0.00%'})
        percentage_format_bold = workbook.add_format({
            "bold": True,
            "text_wrap": True,
            "valign": "bottom",
            "fg_color": "#D7E4BC",
            "border": 1,
            'num_format': '0.00%',
        })

        # Write column headers
        for col_num, value in enumerate(dataframe.columns.values):
            worksheet.write(0, col_num, value, header_format)

        if sheet_name == summaryName:
            # Title
            title_format = workbook.add_format({'bold': True, 'font_size': 14, 'align': 'center'})
            worksheet.merge_range(0, 0, 0, len(dataframe.columns) - 1, f'{summaryName} AS AT {formatted_period}', title_format)

            # Rewrite headers and data starting from row 2
            dataframe.to_excel(writer, sheet_name=sheet_name, startrow=2, header=False, index=False)
            for col_num, value in enumerate(dataframe.columns.values):
                worksheet.write(1, col_num, value, header_format)

            # Apply color scale to multiple columns if specified
            if isinstance(color_column, (list, tuple)):
                color_columns = color_column
            else:
                color_columns = [color_column]

            for col in color_columns:
                if col in dataframe.columns and pd.api.types.is_numeric_dtype(dataframe[col]):
                    col_index = dataframe.columns.get_loc(col)
                    worksheet.conditional_format(
                        2, col_index, len(dataframe) + 1, col_index,
                        {
                            'type': '3_color_scale',
                            'min_color': '#F8696B',
                            'mid_color': '#FFEB84',
                            'max_color': '#63BE7B',
                        }
                    )

            # Row banding
            row_band_format = workbook.add_format({'bg_color': '#F9F9F9'})
            worksheet.conditional_format(2, 0, len(dataframe) + 1, len(dataframe.columns) - 1,
                                         {'type': 'formula', 'criteria': 'MOD(ROW(), 2) = 0', 'format': row_band_format})
            worksheet.hide_gridlines(2)

            # Set column widths
            if column_widths:
                for col_range, width in column_widths.items():
                    worksheet.set_column(col_range, width)
            
            if division_columns is None:
                division_columns = {}  # fallback to empty dict
            
            if division_columns:
                for target_col in division_columns.keys():
                    worksheet.set_column(f"{target_col}:{target_col}", None, percentage_format)

            # Total row with formulas
            header_format.set_align('center')
            total_row = len(dataframe) + 2
            worksheet.merge_range(total_row, 0, total_row, mergeNum, 'Total', header_format)
            for col_num in range(1, len(dataframe.columns)):
                col_letter = chr(65 + col_num)

                if col_letter in division_columns:
                    numerator_col, denominator_col = division_columns[col_letter]
                    worksheet.write_formula(
                        total_row,
                        col_num,
                        f"SUM({numerator_col}3:{numerator_col}{total_row}) / SUM({denominator_col}3:{denominator_col}{total_row})",
                        percentage_format_bold
                    )
                else:
                    worksheet.write_formula(
                        total_row,
                        col_num,
                        f"SUM({col_letter}3:{col_letter}{total_row})",
                        header_format
                    )

        elif sheet_name != summaryName:
            # Row banding
            row_band_format = workbook.add_format({'bg_color': '#F9F9F9'})
            worksheet.conditional_format(1, 0, len(dataframe), len(dataframe.columns) - 1,
                                         {'type': 'formula', 'criteria': 'MOD(ROW(), 2) = 0', 'format': row_band_format})

            # Set column widths
            column_ranges = {
                'O:AF': 10, 'J:M': 12, 'N:N': 40, 'E:F': 15
            }
            for col_range, width in column_ranges.items():
                worksheet.set_column(col_range, width)

            worksheet.set_column('G:G', None, None, {'hidden': True})

    # Save the workbook
    writer.close()
    output.seek(0)
    
    # Resolve absolute path to outputs folder
    output_dir = os.path.join(os.path.dirname(__file__), "..", "outputs")
    os.makedirs(output_dir, exist_ok=True)
    
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"{summaryName}_{timestamp}.xlsx"
    output_path = os.path.join(output_dir, filename)

    with open(output_path, "wb") as f:
        f.write(output.getbuffer())

    return filename