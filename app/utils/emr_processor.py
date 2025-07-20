import pandas as pd

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