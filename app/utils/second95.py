# app.py
from flask import Flask, render_template, request, jsonify, send_file
import pandas as pd
import os
import logging
import traceback
from io import BytesIO
from datetime import datetime
from .emr_processor import process_Linelist, columns_to_select, columns_to_select2

def third95(df, end_date):
    
    # All processing and Excel writing logic goes here.    
    endDate = pd.to_datetime(end_date)
    formatted_period = endDate.strftime('%d %B %Y')  # ✅ use the datetime object

    # --- The rest of the logic goes here ---
    
    # Ensure the 'Pharmacy_LastPickupdate' column is in datetime format and fill NaNs with a specific date
    df['Pharmacy_LastPickupdate2'] = pd.to_datetime(df['Pharmacy_LastPickupdate'], errors='coerce').fillna(pd.to_datetime('1900'))

    #Fill zero if the column contains number greater than 180
    df['DaysOfARVRefill2'] = df['DaysOfARVRefill'].apply(lambda x: 0 if x > 180 else x)

    # Calculate the 'NextAppt' column by adding the 'DaysOfARVRefill2' to 'Pharmacy_LastPickupdate2'
    df['NextAppt'] = df['Pharmacy_LastPickupdate2'] + pd.to_timedelta(df['DaysOfARVRefill2'], unit='D') 

    df.insert(0, 'S/N.', df.index + 1)
        
    #extract actives in both current and previous quarter
    df = df[(df['CurrentARTStatus'] == 'Active') & (df['ARTStatus_PreviousQuarter'] =='Active')]

    today = pd.to_datetime(endDate)
    df['end_date'] = today
    df['ARTStartDate2'] = pd.to_datetime(df['ARTStartDate'].fillna('1900'))

    # Function to calculate difference in months
    def date_diff_in_months2(date1, date2):
        return (date2.year - date1.year) * 12 + date2.month - date1.month - (1 if date2.day < date1.day else 0)

    # Apply the function to the DataFrame
    df['durationOnART'] = df.apply(lambda row: date_diff_in_months2(row['ARTStartDate2'], row['end_date']), axis=1)

    #drop irrelevant columns and extract actual VL eligible from dataframe
    df = df.drop(['end_date','ARTStartDate2'], axis=1)
    df = df[df['durationOnART'] >= 6]

    # Get the current date
    #today = pd.Timestamp('today')
    today = pd.to_datetime(endDate)
    last_30_days = today - pd.Timedelta(days=29)
    next_30_days = today + pd.Timedelta(days=29)

    # Get the last day of the first quarter in the last one year
    last_year = today - pd.DateOffset(years=1)
    first_quarter_last_year = pd.to_datetime(last_year) + pd.tseries.offsets.QuarterEnd(0)

    logging.info(f"First quarter end from last year: {first_quarter_last_year}")  
    
    #convert NextAppt to datetime
    df['NextAppt'] = pd.to_datetime(df['NextAppt'])

    #3rd 95 columns integration
    df['validVlResult'] = df.apply(lambda row: 'Valid Result' if row['DateResultReceivedFacility'] > first_quarter_last_year else 'Invalid Result', axis=1)
    df['validVlSampleCollection'] = df.apply(lambda row: 'Valid SC' if row['LastDateOfSampleCollection'] > first_quarter_last_year else 'Invalid SC', axis=1)
    df['vlSCGap'] = df.apply(lambda row: 'SC Gap' if row['validVlSampleCollection'] == 'Invalid SC' else 'Not SC Gap', axis=1)
    df['PendingResult'] = df.apply(lambda row: 'Pending' if ((row['validVlSampleCollection'] == 'Valid SC') & (row['validVlResult'] == 'Invalid Result')) else 'Not pending', axis=1)  
    df['last30daysmissedSC'] = df.apply(lambda row: 'Missed SC' if ((row['vlSCGap'] == 'SC Gap') & (row['Pharmacy_LastPickupdate'] >= last_30_days) & (row['Pharmacy_LastPickupdate'] < today)) else 'Not Missed SC', axis=1)  
    df['expNext30daysdueforSC'] = df.apply(lambda row: 'due for SC' if ((row['vlSCGap'] == 'SC Gap') & (row['NextAppt'] >= today) & (row['NextAppt'] < next_30_days)) else 'Not due for SC', axis=1)  
    df['Suppression'] = df.apply(lambda row: 'Suppressed' if ((row['validVlResult'] == 'Valid Result') & (row['CurrentViralLoad'] < 1000)) else ('Unsuppressed' if ((row['validVlResult'] == 'Valid Result') & (row['CurrentViralLoad'] > 1000)) else 'Invalid Result'), axis=1)
    
    #df.to_excel('3rd95.xlsx')           
    
    df_sc_Gap = process_Linelist(df, 'vlSCGap', 'SC Gap', columns_to_select2)
    df_pending_results = process_Linelist(df, 'PendingResult', 'Pending', columns_to_select2)
    df_last30daysmissedSC  = process_Linelist(df, 'last30daysmissedSC', 'Missed SC', columns_to_select2)
    df_expNext30daysdueforSC = process_Linelist(df, 'expNext30daysdueforSC', 'due for SC', columns_to_select2)
    df_Suppression = process_Linelist(df, 'Suppression', 'Unsuppressed', columns_to_select2)            
    
    df_active_Eligible = df[df['CurrentARTStatus']=="Active"]

    df_active_Eligible['vlEligible'] = df_active_Eligible['CurrentARTStatus'].apply(lambda x: 1 if x == "Active" else 0)
    df_active_Eligible['validVlResult_valid'] = df_active_Eligible['validVlResult'].apply(lambda x: 1 if x == "Valid Result" else 0)
    df_active_Eligible['validVlSampleCollection'] = df_active_Eligible['validVlSampleCollection'].apply(lambda x: 1 if x == "Valid SC" else 0)
    df_active_Eligible['vlSCGap'] = df_active_Eligible['vlSCGap'].apply(lambda x: 1 if x == "SC Gap" else 0)
    df_active_Eligible['PendingResult'] = df_active_Eligible['PendingResult'].apply(lambda x: 1 if x == "Pending" else 0)
    df_active_Eligible['last30daysmissedSC'] = df_active_Eligible['last30daysmissedSC'].apply(lambda x: 1 if x == "Missed SC" else 0)
    df_active_Eligible['expNext30daysdueforSC'] = df_active_Eligible['expNext30daysdueforSC'].apply(lambda x: 1 if x == "due for SC" else 0)
    df_active_Eligible['Suppressed'] = df_active_Eligible['Suppression'].apply(lambda x: 1 if x == "Suppressed" else 0)

    result = df_active_Eligible.groupby(['LGA','FacilityName'])[['vlEligible', 'validVlResult_valid', 'validVlSampleCollection','vlSCGap', 'PendingResult','last30daysmissedSC','expNext30daysdueforSC','Suppressed']].sum().reset_index()

    result['vl_result_rate'] = ((result['validVlResult_valid'] / result['vlEligible'])).round(4)
    result['sample_collection_rate'] = ((result['validVlSampleCollection'] / result['vlEligible'])).round(4)
    result['suppression_rate'] = ((result['Suppressed'] / result['validVlResult_valid'])).round(4)
    
    result = result[['LGA','FacilityName','vlEligible','validVlResult_valid','Suppressed','vl_result_rate','suppression_rate','validVlSampleCollection','sample_collection_rate','vlSCGap','PendingResult','last30daysmissedSC','expNext30daysdueforSC']]

    result = result.rename(columns={'validVlSampleCollection': 'Valid VL Samples',
                                    'validVlResult_valid': 'Valid VL Results',
                                    'vlEligible': 'Eligible for VL',
                                    'sample_collection_rate': '%VL Sample Collection Rate',
                                    'vl_result_rate': '%VL Coverage',
                                    'suppression_rate': '%VL Suppression Rate',
                                    'PendingResult': 'Pending VL Results',
                                    'vlSCGap': 'VL Sample Collection Gap',
                                    'last30daysmissedSC': 'Last 30 days Missed VL SC',
                                    'expNext30daysdueforSC': 'Expected Next 30 days due for VL SC',
                                    'Suppression': 'Suppressed',})

    # For demonstration, we'll assume `result` is your final summary DataFrame
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    workbook = writer.book

    # Format definitions
    header_format = workbook.add_format({
        "bold": True,
        "text_wrap": True,
        "valign": "bottom",
        "fg_color": "#D7E4BC",
        "border": 1,
    })
    percentage_format = workbook.add_format({'num_format': '0.00%'})
    row_band_format = workbook.add_format({'bg_color': '#F9F9F9'})

    # DataFrames and sheet names
    dataframes = {
        "SC GAP": df_sc_Gap,
        "PENDING RESULT": df_pending_results,
        "LAST 30 DAYS MISSED SC": df_last30daysmissedSC,
        "EXP NEXT 30 DAYS DUE FOR SC": df_expNext30daysdueforSC,
        "UNSUPPRESSED RESULTS": df_Suppression,
        "3RD 95 SUMMARY": result
    }
    
    # Strip time from datetime columns to ensure clean yyyy-mm-dd export
    for df_name, df_data in dataframes.items():
        for col in df_data.columns:
            if pd.api.types.is_datetime64_any_dtype(df_data[col]):
                df_data[col] = df_data[col].dt.date

    for sheet_name, dataframe in dataframes.items():
        dataframe.to_excel(writer, sheet_name=sheet_name, startrow=1, header=False, index=False)
        worksheet = writer.sheets[sheet_name]

        # Write headers
        for col_num, column_name in enumerate(dataframe.columns.values):
            worksheet.write(0, col_num, column_name, header_format)

        # Alternating row background
        worksheet.conditional_format(
            2, 0, len(dataframe) + 1, len(dataframe.columns) - 1,
            {'type': 'formula', 'criteria': 'MOD(ROW(),2)=0', 'format': row_band_format}
        )

        # Special formatting for summary
        if sheet_name == "3RD 95 SUMMARY":
            title_format = workbook.add_format({'bold': True, 'font_size': 14, 'align': 'center'})
            worksheet.merge_range(0, 0, 0, len(dataframe.columns) - 1, f'3RD 95 SUMMARY AS AT {formatted_period}', title_format)

            # Apply % formatting
            worksheet.set_column('F:F', None, percentage_format)
            worksheet.set_column('G:G', None, percentage_format)
            worksheet.set_column('I:I', None, percentage_format)


    writer.close()
    output.seek(0)

    # Resolve absolute path to outputs folder
    output_dir = os.path.join(os.path.dirname(__file__), "..", "outputs")
    os.makedirs(output_dir, exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"3RD_95_ANALYSIS_{timestamp}.xlsx"
    output_path = os.path.abspath(os.path.join(output_dir, filename))

    # Save Excel file
    with open(output_path, "wb") as f:
        f.write(output.getbuffer())

    return filename  # just filename, not full path


def third95CMG(df, end_date):
        # All processing and Excel writing logic goes here.    
        endDate = pd.to_datetime(end_date)
        formatted_period = endDate.strftime('%d %B %Y')  # ✅ use the datetime object
        df['CaseManager'] = df['CaseManager'].fillna('UNASSIGNED')
              
        # Ensure the 'Pharmacy_LastPickupdate' column is in datetime format and fill NaNs with a specific date
        df['Pharmacy_LastPickupdate2'] = pd.to_datetime(df['Pharmacy_LastPickupdate'], errors='coerce').fillna(pd.to_datetime('1900'))

        #Fill zero if the column contains number greater than 180
        df['DaysOfARVRefill2'] = df['DaysOfARVRefill'].apply(lambda x: 0 if x > 180 else x)

        # Calculate the 'NextAppt' column by adding the 'DaysOfARVRefill2' to 'Pharmacy_LastPickupdate2'
        df['NextAppt'] = df['Pharmacy_LastPickupdate2'] + pd.to_timedelta(df['DaysOfARVRefill2'], unit='D')     
        
        df = df.drop(['LGA2','STATE2'], axis=1)
        df.insert(0, 'S/N.', df.index + 1)
            
        #extract actives in both current and previous quarter
        df = df[(df['CurrentARTStatus'] == 'Active') & (df['ARTStatus_PreviousQuarter'] =='Active')]

        today = pd.to_datetime(endDate)
        df['end_date'] = today
        df['ARTStartDate2'] = pd.to_datetime(df['ARTStartDate'].fillna('1900'))

        # Function to calculate difference in months
        def date_diff_in_months2(date1, date2):
            return (date2.year - date1.year) * 12 + date2.month - date1.month - (1 if date2.day < date1.day else 0)

        # Apply the function to the DataFrame
        df['durationOnART'] = df.apply(lambda row: date_diff_in_months2(row['ARTStartDate2'], row['end_date']), axis=1)

        #drop irrelevant columns and extract actual VL eligible from dataframe
        df = df.drop(['end_date','ARTStartDate2'], axis=1)
        df = df[df['durationOnART'] >= 6]

        # Get the current date
        #today = pd.Timestamp('today')
        today = pd.to_datetime(endDate)
        last_30_days = today - pd.Timedelta(days=29)
        next_30_days = today + pd.Timedelta(days=29)

        # Get the last day of the first quarter in the last one year
        last_year = today - pd.DateOffset(years=1)
        first_quarter_last_year = pd.to_datetime(last_year) + pd.tseries.offsets.QuarterEnd(0)

        print(first_quarter_last_year)     
        
        #convert VL dates to datetime
        df['DateResultReceivedFacility'] = pd.to_datetime(df['DateResultReceivedFacility'])
        df['LastDateOfSampleCollection'] = pd.to_datetime(df['LastDateOfSampleCollection'])
        df['Pharmacy_LastPickupdate'] = pd.to_datetime(df['Pharmacy_LastPickupdate'])
        df['NextAppt'] = pd.to_datetime(df['NextAppt'])

        #3rd 95 columns integration
        df['validVlResult'] = df.apply(lambda row: 'Valid Result' if row['DateResultReceivedFacility'] > first_quarter_last_year else 'Invalid Result', axis=1)
        df['validVlSampleCollection'] = df.apply(lambda row: 'Valid SC' if row['LastDateOfSampleCollection'] > first_quarter_last_year else 'Invalid SC', axis=1)
        df['vlSCGap'] = df.apply(lambda row: 'SC Gap' if row['validVlSampleCollection'] == 'Invalid SC' else 'Not SC Gap', axis=1)
        df['PendingResult'] = df.apply(lambda row: 'Pending' if ((row['validVlSampleCollection'] == 'Valid SC') & (row['validVlResult'] == 'Invalid Result')) else 'Not pending', axis=1)    
        df['last30daysmissedSC'] = df.apply(lambda row: 'Missed SC' if ((row['vlSCGap'] == 'SC Gap') & (row['Pharmacy_LastPickupdate'] >= last_30_days) & (row['Pharmacy_LastPickupdate'] < today)) else 'Not Missed SC', axis=1)  
        df['expNext30daysdueforSC'] = df.apply(lambda row: 'due for SC' if ((row['vlSCGap'] == 'SC Gap') & (row['NextAppt'] >= today) & (row['NextAppt'] < next_30_days)) else 'Not due for SC', axis=1)  
        #df['Suppression'] = df.apply(lambda row: 'Suppressed' if ((row['validVlResult'] == 'Valid Result') & (row['CurrentViralLoad2'] < 1000)) else 'Not Suppressed', axis=1)
        df['Suppression'] = df.apply(lambda row: 'Suppressed' if ((row['validVlResult'] == 'Valid Result') & (row['CurrentViralLoad'] < 1000)) else ('Unsuppressed' if ((row['validVlResult'] == 'Valid Result') & (row['CurrentViralLoad'] > 1000)) else 'Invalid Result'), axis=1)
        #df.to_excel("3rd95.xlsx")
        
        df_sc_Gap = df[df['vlSCGap'] == 'SC Gap']
        df_sc_Gap = df_sc_Gap.sort_values(by=['CaseManager'], ascending=True).reset_index()
        df_sc_Gap.insert(0, 'S/N', df_sc_Gap.index + 1)
        
        df_pending_results = df[df['PendingResult'] == 'Pending']
        df_pending_results = df_pending_results.sort_values(by=['CaseManager'], ascending=True).reset_index()
        df_pending_results.insert(0, 'S/N', df_pending_results.index + 1)
        
        df_last30daysmissedSC = df[df['last30daysmissedSC'] == 'Missed SC']
        df_last30daysmissedSC = df_last30daysmissedSC.sort_values(by=['CaseManager'], ascending=True).reset_index()
        df_last30daysmissedSC.insert(0, 'S/N', df_last30daysmissedSC.index + 1)
        
        df_expNext30daysdueforSC = df[df['expNext30daysdueforSC'] == 'due for SC']
        df_expNext30daysdueforSC= df_expNext30daysdueforSC.sort_values(by=['CaseManager'], ascending=True).reset_index()
        df_expNext30daysdueforSC.insert(0, 'S/N', df_expNext30daysdueforSC.index + 1)
        
        df_Suppression = df[df['Suppression'] == 'Unsuppressed']
        df_Suppression = df_Suppression.sort_values(by=['CaseManager'], ascending=True).reset_index()
        df_Suppression.insert(0, 'S/N', df_Suppression.index + 1)
        
        #Trime Line lists
        df_sc_Gap = df_sc_Gap[["S/N", "State", "LGA", "FacilityName", "PEPID", "PatientHospitalNo", "uuid", "Sex", "Current_Age", "Surname", "Firstname", "MaritalStatus", "PhoneNo", "Address", "State_of_Residence", "LGA_of_Residence", "DateConfirmedHIV+", "ARTStartDate", "Pharmacy_LastPickupdate", "DaysOfARVRefill", "CurrentARTRegimen", "NextAppt", "CurrentPregnancyStatus", "CurrentViralLoad", "DateResultReceivedFacility", "Alphanumeric_Viral_Load_Result", "LastDateOfSampleCollection", "Outcomes", "Outcomes_Date", "IIT_Date", "CurrentARTStatus", "First_TPT_Pickupdate", "Current_TPT_Received", "PBS_Capturee", "PBS_Capture_Date", "Date_Generated","CaseManager"]]
        df_pending_results = df_pending_results[["S/N", "State", "LGA", "FacilityName", "PEPID", "PatientHospitalNo", "uuid", "Sex", "Current_Age", "Surname", "Firstname", "MaritalStatus", "PhoneNo", "Address", "State_of_Residence", "LGA_of_Residence", "DateConfirmedHIV+", "ARTStartDate", "Pharmacy_LastPickupdate", "DaysOfARVRefill", "CurrentARTRegimen", "NextAppt", "CurrentPregnancyStatus", "CurrentViralLoad", "DateResultReceivedFacility", "Alphanumeric_Viral_Load_Result", "LastDateOfSampleCollection", "Outcomes", "Outcomes_Date", "IIT_Date", "CurrentARTStatus", "First_TPT_Pickupdate", "Current_TPT_Received", "PBS_Capturee", "PBS_Capture_Date", "Date_Generated","CaseManager"]]
        df_last30daysmissedSC = df_last30daysmissedSC[["S/N", "State", "LGA", "FacilityName", "PEPID", "PatientHospitalNo", "uuid", "Sex", "Current_Age", "Surname", "Firstname", "MaritalStatus", "PhoneNo", "Address", "State_of_Residence", "LGA_of_Residence", "DateConfirmedHIV+", "ARTStartDate", "Pharmacy_LastPickupdate", "DaysOfARVRefill", "CurrentARTRegimen", "NextAppt", "CurrentPregnancyStatus", "CurrentViralLoad", "DateResultReceivedFacility", "Alphanumeric_Viral_Load_Result", "LastDateOfSampleCollection", "Outcomes", "Outcomes_Date", "IIT_Date", "CurrentARTStatus", "First_TPT_Pickupdate", "Current_TPT_Received", "PBS_Capturee", "PBS_Capture_Date", "Date_Generated","CaseManager"]]
        df_expNext30daysdueforSC = df_expNext30daysdueforSC[["S/N", "State", "LGA", "FacilityName", "PEPID", "PatientHospitalNo", "uuid", "Sex", "Current_Age", "Surname", "Firstname", "MaritalStatus", "PhoneNo", "Address", "State_of_Residence", "LGA_of_Residence", "DateConfirmedHIV+", "ARTStartDate", "Pharmacy_LastPickupdate", "DaysOfARVRefill", "CurrentARTRegimen", "NextAppt", "CurrentPregnancyStatus", "CurrentViralLoad", "DateResultReceivedFacility", "Alphanumeric_Viral_Load_Result", "LastDateOfSampleCollection", "Outcomes", "Outcomes_Date", "IIT_Date", "CurrentARTStatus", "First_TPT_Pickupdate", "Current_TPT_Received", "PBS_Capturee", "PBS_Capture_Date", "Date_Generated","CaseManager"]]
        df_Suppression = df_Suppression[["S/N", "State", "LGA", "FacilityName", "PEPID", "PatientHospitalNo", "uuid", "Sex", "Current_Age", "Surname", "Firstname", "MaritalStatus", "PhoneNo", "Address", "State_of_Residence", "LGA_of_Residence", "DateConfirmedHIV+", "ARTStartDate", "Pharmacy_LastPickupdate", "DaysOfARVRefill", "CurrentARTRegimen", "NextAppt", "CurrentPregnancyStatus", "CurrentViralLoad", "DateResultReceivedFacility", "Alphanumeric_Viral_Load_Result", "LastDateOfSampleCollection", "Outcomes", "Outcomes_Date", "IIT_Date", "CurrentARTStatus", "First_TPT_Pickupdate", "Current_TPT_Received", "PBS_Capturee", "PBS_Capture_Date", "Date_Generated","CaseManager"]]
        
        df_active_Eligible = df[df['CurrentARTStatus']=="Active"]

        df_active_Eligible['vlEligible'] = df_active_Eligible['CurrentARTStatus'].apply(lambda x: 1 if x == "Active" else 0)
        df_active_Eligible['validVlResult_valid'] = df_active_Eligible['validVlResult'].apply(lambda x: 1 if x == "Valid Result" else 0)
        df_active_Eligible['validVlSampleCollection'] = df_active_Eligible['validVlSampleCollection'].apply(lambda x: 1 if x == "Valid SC" else 0)
        df_active_Eligible['vlSCGap'] = df_active_Eligible['vlSCGap'].apply(lambda x: 1 if x == "SC Gap" else 0)
        df_active_Eligible['PendingResult'] = df_active_Eligible['PendingResult'].apply(lambda x: 1 if x == "Pending" else 0)
        df_active_Eligible['last30daysmissedSC'] = df_active_Eligible['last30daysmissedSC'].apply(lambda x: 1 if x == "Missed SC" else 0)
        df_active_Eligible['expNext30daysdueforSC'] = df_active_Eligible['expNext30daysdueforSC'].apply(lambda x: 1 if x == "due for SC" else 0)
        df_active_Eligible['Suppressed'] = df_active_Eligible['Suppression'].apply(lambda x: 1 if x == "Suppressed" else 0)

        #result = df_active_Eligible.groupby(['LGA','FacilityName','CaseManager'])[['vlEligible', 'validVlResult_valid', 'validVlSampleCollection','vlSCGap', 'PendingResult']].sum().reset_index()
        result = df_active_Eligible.groupby(['LGA','FacilityName','CaseManager'])[['vlEligible', 'validVlResult_valid', 'validVlSampleCollection','vlSCGap', 'PendingResult','last30daysmissedSC','expNext30daysdueforSC','Suppressed']].sum().reset_index()

        result['vl_result_rate'] = ((result['validVlResult_valid'] / result['vlEligible'])).round(4)
        result['sample_collection_rate'] = ((result['validVlSampleCollection'] / result['vlEligible'])).round(4)
        result['suppression_rate'] = ((result['Suppressed'] / result['validVlResult_valid'])).round(4)
        
        result = result[['LGA','FacilityName','CaseManager','vlEligible','validVlResult_valid','Suppressed','vl_result_rate','suppression_rate','validVlSampleCollection','sample_collection_rate','vlSCGap','PendingResult','last30daysmissedSC','expNext30daysdueforSC']]

        result = result.rename(columns={'validVlSampleCollection': 'Valid VL Samples',
                                        'validVlResult_valid': 'Valid VL Results',
                                        'vlEligible': 'Eligible for VL',
                                        'sample_collection_rate': '%VL Sample Collection Rate',
                                        'vl_result_rate': '%VL Coverage',
                                        'suppression_rate': '%VL Suppression Rate',
                                        'PendingResult': 'Pending VL Results',
                                        'vlSCGap': 'VL Sample Collection Gap',
                                        'last30daysmissedSC': 'Last 30 days Missed VL SC',
                                        'expNext30daysdueforSC': 'Expected Next 30 days due for VL SC',
                                        'Suppression': 'Suppressed',})
        
        #format and export
        sdate = endDate.strftime("%d-%m-%Y")


        # Create a Pandas Excel writer using XlsxWriter as the engine.
        #writer = pd.ExcelWriter(output_file_path, engine="xlsxwriter")
        output = BytesIO()
        writer = pd.ExcelWriter(output, engine='xlsxwriter')
        workbook = writer.book

        # List of dataframes and their corresponding sheet names
        dataframes = {
            "SC GAP": df_sc_Gap,
            "PENDING RESULT": df_pending_results,
            "LAST 30 DAYS MISSED SC": df_last30daysmissedSC,
            "EXP NEXT 30 DAYS DUE FOR SC": df_expNext30daysdueforSC,
            "UNSUPPRESSED RESULTS": df_Suppression,
            "3RD 95 SUMMARY": result,
            #"Sheet3": df3,
            # Add more dataframes and sheet names as needed
        }

        # Write each dataframe to a different sheet
        for sheet_name, dataframe in dataframes.items():
            dataframe.to_excel(writer, sheet_name=sheet_name, startrow=1, header=False, index=False)
            
            workbook = writer.book
            worksheet = writer.sheets[sheet_name]
            
            # Add a header format.
            header_format = workbook.add_format(
                {
                    "bold": True,
                    "text_wrap": True,
                    "valign": "bottom",
                    "fg_color": "#D7E4BC",
                    "border": 1,
                }
            )
            
            # Write the column headers with the defined format.
            for col_num, value in enumerate(dataframe.columns.values):
                worksheet.write(0, col_num, value, header_format)                 
            
            # Add percentage format to a specific column in "2ND 95 SUMMARY" (e.g., column B which is index 1)
            if sheet_name == "3RD 95 SUMMARY":
                # Add title to the worksheet
                title_format = workbook.add_format({'bold': True, 'font_size': 14, 'align': 'center'})
                worksheet.merge_range(0, 0, 0, len(dataframe.columns) - 1, f'3RD 95 SUMMARY AS AT {formatted_period}', title_format)
                
                # Write the dataframe to the sheet, starting from row 2 to leave space for the title and headers
                dataframe.to_excel(writer, sheet_name=sheet_name, startrow=2, header=False, index=False)
                
                # Write the column headers with the defined format, starting from row 1
                for col_num, value in enumerate(dataframe.columns.values):
                    worksheet.write(1, col_num, value, header_format)
                
                # Add percentage format to columns
                percentage_format2 = workbook.add_format(
                    {
                        "bold": True,
                        "text_wrap": True,
                        "valign": "bottom",
                        "fg_color": "#D7E4BC",
                        "border": 1,
                        'num_format': '0.00%',
                    }
                )
                percentage_format = workbook.add_format({'num_format': '0.00%'})
                worksheet.set_column('G:G', None, percentage_format)
                worksheet.set_column('H:H', None, percentage_format)
                worksheet.set_column('J:J', None, percentage_format)
                
                colorscales = ['%VL Sample Collection Rate', '%VL Coverage', '%VL Suppression Rate']
                for col_name in colorscales:
                    if col_name in dataframe.columns and pd.api.types.is_numeric_dtype(dataframe[col_name]):
                        col_index = dataframe.columns.get_loc(col_name)
                
                        start_row = 2  # Excel rows are 1-indexed; row 1 is header
                        end_row = len(dataframe) + 1
                
                        # Apply 3-color scale conditional formatting
                        worksheet.conditional_format(
                            start_row, col_index, end_row, col_index,
                            {
                                'type': '3_color_scale',
                                'min_color': '#F8696B',
                                'mid_color': '#FFEB84',
                                'max_color': '#63BE7B',
                            }
                        )
                
                row_band_format = workbook.add_format({'bg_color': '#F9F9F9'})
                worksheet.conditional_format(2, 0, len(dataframe) + 1, len(dataframe.columns) - 1, 
                                            {'type': 'formula', 'criteria': 'MOD(ROW(), 2) = 0', 'format': row_band_format})
                worksheet.hide_gridlines(2)
                
                column_ranges = {
                    'A:A': 20,  # Example width for columns A to A
                    'B:B': 35,  # Example width for columns B to B
                    'C:C': 25,  # Example width for columns C to C
                    # Add more column ranges and their widths as needed
                }
                for col_range, width in column_ranges.items():
                    worksheet.set_column(col_range, width)
                
                # Add total row
                header_format.set_align('center')
                total_row = len(dataframe) + 2
                worksheet.merge_range(total_row, 0, total_row, 2, 'Total', header_format)
                for col_num in range(1, len(dataframe.columns)):
                    col_letter = chr(65 + col_num)  # Convert column number to letter
                    #worksheet.write_formula(total_row, col_num, f'SUM({col_letter}2:{col_letter}{total_row})', header_format)
                    if col_letter == 'G':  # Column F to be calculated as D / C
                        worksheet.write_formula(total_row, col_num, f'SUM(E3:E{total_row}) / SUM(D3:D{total_row})', percentage_format2)
                    elif col_letter == 'H':  # Column G to be calculated as E / C
                        worksheet.write_formula(total_row, col_num, f'SUM(F3:F{total_row}) / SUM(E3:E{total_row})', percentage_format2)
                    elif col_letter == 'J':  # Column I to be calculated as H / C
                        worksheet.write_formula(total_row, col_num, f'SUM(I3:I{total_row}) / SUM(D3:D{total_row})', percentage_format2)
                    else:
                        worksheet.write_formula(total_row, col_num, f'SUM({col_letter}3:{col_letter}{total_row})', header_format)
            
            else:
                # Write the dataframe to the sheet, starting from row 1
                dataframe.to_excel(writer, sheet_name=sheet_name, startrow=1, header=False, index=False)
                
                # Write the column headers with the defined format, starting from row 0
                for col_num, value in enumerate(dataframe.columns.values):
                    worksheet.write(0, col_num, value, header_format)
            
            # Add row band format to "SC GAP" and "PENDING RESULT"
            if sheet_name in ["SC GAP", "PENDING RESULT", "LAST 30 DAYS MISSED SC", "EXP NEXT 30 DAYS DUE FOR SC", "UNSUPPRESSED RESULTS"]:
                row_band_format = workbook.add_format({'bg_color': '#F9F9F9'})
                worksheet.conditional_format(1, 0, len(dataframe), len(dataframe.columns) - 1, 
                                            {'type': 'formula', 'criteria': 'MOD(ROW(), 2) = 0', 'format': row_band_format})
                                
                # Specify column widths for a range of columns
                column_ranges = {
                    'O:AG': 10,  # Example width for columns A to C
                    'J:M': 12,  # Example width for columns D to E
                    'N:N': 40,  # Example width for columns G to H
                    'E:F': 15,  # Example width for columns G to H
                    # Add more column ranges and their widths as needed
                }
                for col_range, width in column_ranges.items():
                    worksheet.set_column(col_range, width)
                
                worksheet.set_column('G:G', None, None, {'hidden': True})
                
        # Close the Pandas Excel writer and output the Excel file.
        writer.close()


def second95(df, endDate):
    # Ensure the 'Pharmacy_LastPickupdate' column is in datetime format and fill NaNs with a specific date
    df['Pharmacy_LastPickupdate2'] = pd.to_datetime(df['Pharmacy_LastPickupdate'], errors='coerce').fillna(pd.to_datetime('1900'))

    #Fill zero if the column contains number greater than 180
    df['DaysOfARVRefill2'] = df['DaysOfARVRefill'].apply(lambda x: 0 if x > 180 else x)

    # Calculate the 'NextAppt' column by adding the 'DaysOfARVRefill2' to 'Pharmacy_LastPickupdate2'
    df['NextAppt'] = df['Pharmacy_LastPickupdate2'] + pd.to_timedelta(df['DaysOfARVRefill2'], unit='D') 
    df['IITDate2'] = (df['NextAppt'] + pd.Timedelta(days=29)).fillna('1900')
    
    df = df.drop(['LGA2','STATE2'], axis=1)
    df.insert(0, 'S/N.', df.index + 1)
    
    today = pd.Timestamp(endDate)
    currentyear = today.year
    previousyear = today.year-1
    currentmonthyear = str(today.month) + '_' + str(today.year)
    print(currentmonthyear)
    sevenDaysIIT = today + pd.Timedelta(days=7)

    print(currentyear)
    print(previousyear)
    
    #2nd 95 columns integration
    df['IITDate2'] = pd.to_datetime(df['IITDate2'])
    df['IITYear'] = df['IITDate2'].dt.year
    df['NextApptMonthYear'] = df['NextAppt'].fillna('1900').dt.month.astype(int).astype(str) + '_' + df['NextAppt'].fillna('1900').dt.year.astype(int).astype(str)
    print(df['NextApptMonthYear'])
    print(df['IITYear'])
    df['CurrentYearIIT'] = df.apply(lambda row: 'CurrentYearIIT' if ((row['IITYear'] == currentyear) & (row['IITDate2'] <= today) & (row['CurrentARTStatus'] in ['LTFU', 'Lost to followup'])) else 'notCurrentYearIIT', axis=1)
    df['previousyearIIT'] = df.apply(lambda row: 'previousyearIIT' if ((row['IITYear'] == previousyear) & (row['IITDate2'] <= today) & (row['CurrentARTStatus'] in ['LTFU', 'Lost to followup'])) else 'notpreviousyearIIT', axis=1) 
    df['ImminentIIT'] = df.apply(lambda row: 'ImminentIIT' if ((row['NextAppt'] <= today) & ((row['CurrentARTStatus'] == 'Active'))) else 'notImminentIIT', axis=1)
    df['sevendaysIIT'] = df.apply(lambda row: 'sevendaysIIT' if ((row['IITDate2'] <= sevenDaysIIT) & ((row['CurrentARTStatus'] == 'Active'))) else 'notsevenDaysIIT', axis=1)
    df['currentmonthexpected'] = df.apply(lambda row: 'currentmonthexpected' if ((row['NextApptMonthYear'] == currentmonthyear) & ((row['CurrentARTStatus'] == 'Active'))) else 'notcurrentmonthexpected', axis=1)
    
    #convert NextAppt to datetime
    df['NextAppt'] = pd.to_datetime(df['NextAppt'])
    
    dfCurrentYearIIT = df[df['CurrentYearIIT'] == 'CurrentYearIIT'].reset_index()
    dfCurrentYearIIT.insert(0, 'S/N', dfCurrentYearIIT.index + 1)
    
    dfpreviousyearIIT = df[df['previousyearIIT'] == 'previousyearIIT'].reset_index()
    dfpreviousyearIIT.insert(0, 'S/N', dfpreviousyearIIT.index + 1)
    
    dfImminentIIT = df[df['ImminentIIT'] == 'ImminentIIT'].reset_index()
    dfImminentIIT.insert(0, 'S/N', dfImminentIIT.index + 1)
    
    dfsevendaysIIT = df[df['sevendaysIIT'] == 'sevendaysIIT'].reset_index()
    dfsevendaysIIT.insert(0, 'S/N', dfsevendaysIIT.index + 1)
    
    dfcurrentmonthexpected = df[df['currentmonthexpected'] == 'currentmonthexpected'].reset_index()
    dfcurrentmonthexpected.insert(0, 'S/N', dfcurrentmonthexpected.index + 1)
    
    #Trime Line lists
    dfCurrentYearIIT = dfCurrentYearIIT[["S/N", "State", "LGA", "FacilityName", "PEPID", "PatientHospitalNo", "uuid", "Sex", "Current_Age", "Surname", "Firstname", "MaritalStatus", "PhoneNo", "Address", "State_of_Residence", "LGA_of_Residence", "DateConfirmedHIV+", "ARTStartDate", "Pharmacy_LastPickupdate", "DaysOfARVRefill", "CurrentARTRegimen", "NextAppt", "CurrentPregnancyStatus", "CurrentViralLoad", "DateResultReceivedFacility", "Alphanumeric_Viral_Load_Result", "LastDateOfSampleCollection", "Outcomes", "Outcomes_Date", "IIT_Date", "CurrentARTStatus", "First_TPT_Pickupdate", "Current_TPT_Received", "PBS_Capturee", "PBS_Capture_Date", "Date_Generated"]]
    dfpreviousyearIIT = dfpreviousyearIIT[["S/N", "State", "LGA", "FacilityName", "PEPID", "PatientHospitalNo", "uuid", "Sex", "Current_Age", "Surname", "Firstname", "MaritalStatus", "PhoneNo", "Address", "State_of_Residence", "LGA_of_Residence", "DateConfirmedHIV+", "ARTStartDate", "Pharmacy_LastPickupdate", "DaysOfARVRefill", "CurrentARTRegimen", "NextAppt", "CurrentPregnancyStatus", "CurrentViralLoad", "DateResultReceivedFacility", "Alphanumeric_Viral_Load_Result", "LastDateOfSampleCollection", "Outcomes", "Outcomes_Date", "IIT_Date", "CurrentARTStatus", "First_TPT_Pickupdate", "Current_TPT_Received", "PBS_Capturee", "PBS_Capture_Date", "Date_Generated"]]
    dfImminentIIT = dfImminentIIT[["S/N", "State", "LGA", "FacilityName", "PEPID", "PatientHospitalNo", "uuid", "Sex", "Current_Age", "Surname", "Firstname", "MaritalStatus", "PhoneNo", "Address", "State_of_Residence", "LGA_of_Residence", "DateConfirmedHIV+", "ARTStartDate", "Pharmacy_LastPickupdate", "DaysOfARVRefill", "CurrentARTRegimen", "NextAppt", "CurrentPregnancyStatus", "CurrentViralLoad", "DateResultReceivedFacility", "Alphanumeric_Viral_Load_Result", "LastDateOfSampleCollection", "Outcomes", "Outcomes_Date", "IIT_Date", "CurrentARTStatus", "First_TPT_Pickupdate", "Current_TPT_Received", "PBS_Capturee", "PBS_Capture_Date", "Date_Generated"]]
    dfsevendaysIIT = dfsevendaysIIT[["S/N", "State", "LGA", "FacilityName", "PEPID", "PatientHospitalNo", "uuid", "Sex", "Current_Age", "Surname", "Firstname", "MaritalStatus", "PhoneNo", "Address", "State_of_Residence", "LGA_of_Residence", "DateConfirmedHIV+", "ARTStartDate", "Pharmacy_LastPickupdate", "DaysOfARVRefill", "CurrentARTRegimen", "NextAppt", "CurrentPregnancyStatus", "CurrentViralLoad", "DateResultReceivedFacility", "Alphanumeric_Viral_Load_Result", "LastDateOfSampleCollection", "Outcomes", "Outcomes_Date", "IIT_Date", "CurrentARTStatus", "First_TPT_Pickupdate", "Current_TPT_Received", "PBS_Capturee", "PBS_Capture_Date", "Date_Generated"]]
    dfcurrentmonthexpected = dfcurrentmonthexpected[["S/N", "State", "LGA", "FacilityName", "PEPID", "PatientHospitalNo", "uuid", "Sex", "Current_Age", "Surname", "Firstname", "MaritalStatus", "PhoneNo", "Address", "State_of_Residence", "LGA_of_Residence", "DateConfirmedHIV+", "ARTStartDate", "Pharmacy_LastPickupdate", "DaysOfARVRefill", "CurrentARTRegimen", "NextAppt", "CurrentPregnancyStatus", "CurrentViralLoad", "DateResultReceivedFacility", "Alphanumeric_Viral_Load_Result", "LastDateOfSampleCollection", "Outcomes", "Outcomes_Date", "IIT_Date", "CurrentARTStatus", "First_TPT_Pickupdate", "Current_TPT_Received", "PBS_Capturee", "PBS_Capture_Date", "Date_Generated"]]
    
    df2nd95Summary = df

    df2nd95Summary['ActiveClients'] = df2nd95Summary['CurrentARTStatus'].apply(lambda x: 1 if x == "Active" else 0)
    df2nd95Summary['currentYearIIT'] = df2nd95Summary['CurrentYearIIT'].apply(lambda x: 1 if x == "CurrentYearIIT" else 0)
    df2nd95Summary['previousyearIIT'] = df2nd95Summary['previousyearIIT'].apply(lambda x: 1 if x == "previousyearIIT" else 0)
    df2nd95Summary['ImminentIIT'] = df2nd95Summary['ImminentIIT'].apply(lambda x: 1 if x == "ImminentIIT" else 0)
    df2nd95Summary['sevendaysIIT'] = df2nd95Summary['sevendaysIIT'].apply(lambda x: 1 if x == "sevendaysIIT" else 0)
    df2nd95Summary['currentmonthexpected'] = df2nd95Summary['currentmonthexpected'].apply(lambda x: 1 if x == "currentmonthexpected" else 0)

    result = df2nd95Summary.groupby(['LGA','FacilityName'])[['ActiveClients', 'currentYearIIT', 'previousyearIIT','ImminentIIT','sevendaysIIT','currentmonthexpected']].sum().reset_index()
    result_check = ['ActiveClients', 'currentYearIIT', 'previousyearIIT', 'ImminentIIT', 'sevendaysIIT', 'currentmonthexpected']
    result = result[(result[result_check] != 0).any(axis=1)]


    result = result.rename(columns={'ActiveClients': 'Active Clients',
                                    'currentYearIIT': 'current FY IIT',
                                    'previousyearIIT': 'previous FY IIT',
                                    'ImminentIIT': 'Imminent IIT',
                                    'sevendaysIIT': 'Next 7 days IIT',
                                    'currentmonthexpected': 'EXPECTED THIS MONTH',})
    result   
    
    #format and export
    sdate = endDate.strftime("%d-%m-%Y")

    print(sdate)
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    workbook = writer.book

    # List of dataframes and their corresponding sheet names
    dataframes = {
        "CURRENT FY IIT": dfCurrentYearIIT,
        "PREVIOUS FY IIT": dfpreviousyearIIT,
        "IMMINENT IIT": dfImminentIIT,
        "NEXT 7 DAYS IIT": dfsevendaysIIT,
        "EXPECTED THIS MONTH": dfcurrentmonthexpected,
        "2ND 95 SUMMARY": result,
        # Add more dataframes and sheet names as needed
    }

    # Write each dataframe to a different sheet
    for sheet_name, dataframe in dataframes.items():
        dataframe.to_excel(writer, sheet_name=sheet_name, startrow=1, header=False, index=False)
        
        workbook = writer.book
        worksheet = writer.sheets[sheet_name]
        
        # Add a header format.
        header_format = workbook.add_format(
            {
                "bold": True,
                "text_wrap": True,
                "valign": "bottom",
                "fg_color": "#D7E4BC",
                "border": 1,
            }
        )
        
        # Write the column headers with the defined format.
        for col_num, value in enumerate(dataframe.columns.values):
            worksheet.write(0, col_num, value, header_format)
            
        # Add percentage format to a specific column in "2ND 95 SUMMARY" (e.g., column B which is index 1)
        if sheet_name == "2ND 95 SUMMARY":
            # Add title to the worksheet
            title_format = workbook.add_format({'bold': True, 'font_size': 14, 'align': 'center'})
            worksheet.merge_range(0, 0, 0, len(dataframe.columns) - 1, '2ND 95 SUMMARY', title_format)
            
            # Write the dataframe to the sheet, starting from row 2 to leave space for the title and headers
            dataframe.to_excel(writer, sheet_name=sheet_name, startrow=2, header=False, index=False)
            
            # Write the column headers with the defined format, starting from row 1
            for col_num, value in enumerate(dataframe.columns.values):
                worksheet.write(1, col_num, value, header_format)
            
            # Add percentage format to column B
            percentage_format = workbook.add_format({'num_format': '0.00%'})
            worksheet.set_column('B:B', None, percentage_format)
            
            row_band_format = workbook.add_format({'bg_color': '#F9F9F9'})
            worksheet.conditional_format(2, 0, len(dataframe) + 1, len(dataframe.columns) - 1, 
                                        {'type': 'formula', 'criteria': 'MOD(ROW(), 2) = 0', 'format': row_band_format})
            worksheet.hide_gridlines(2)
            
            column_ranges = {
                'A:A': 20,  # Example width for columns A to A
                'B:B': 35,  # Example width for columns B to B
                # Add more column ranges and their widths as needed
            }
            for col_range, width in column_ranges.items():
                worksheet.set_column(col_range, width)
            
            # Add total row
            header_format.set_align('center')
            total_row = len(dataframe) + 2
            worksheet.merge_range(total_row, 0, total_row, 1, 'Total', header_format)
            for col_num in range(1, len(dataframe.columns)):
                col_letter = chr(65 + col_num)  # Convert column number to letter
                worksheet.write_formula(total_row, col_num, f'SUM({col_letter}2:{col_letter}{total_row})', header_format)
        
        else:
            # Write the dataframe to the sheet, starting from row 1
            dataframe.to_excel(writer, sheet_name=sheet_name, startrow=1, header=False, index=False)
            
            # Write the column headers with the defined format, starting from row 0
            for col_num, value in enumerate(dataframe.columns.values):
                worksheet.write(0, col_num, value, header_format)
        
        # Add row band format to specific sheets
        if sheet_name in ["CURRENT FY IIT", "PREVIOUS FY IIT", "IMMINENT IIT", "NEXT 7 DAYS IIT", "EXPECTED THIS MONTH"]:
            row_band_format = workbook.add_format({'bg_color': '#F9F9F9'})
            worksheet.conditional_format(1, 0, len(dataframe), len(dataframe.columns) - 1, 
                                        {'type': 'formula', 'criteria': 'MOD(ROW(), 2) = 0', 'format': row_band_format})
            
            # Specify column widths for a range of columns
            column_ranges = {
                'O:AF': 10,  # Example width for columns O to AF
                'J:M': 12,  # Example width for columns J to M
                'N:N': 40,  # Example width for column N
                'E:F': 15,  # Example width for columns E to F
                # Add more column ranges and their widths as needed
            }
            for col_range, width in column_ranges.items():
                worksheet.set_column(col_range, width)
            
            worksheet.set_column('G:G', None, None, {'hidden': True})
        
    writer.close()
    output.seek(0)

    # Resolve absolute path to outputs folder
    output_dir = os.path.join(os.path.dirname(__file__), "..", "outputs")
    os.makedirs(output_dir, exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"2nd_95_ANALYSIS_{timestamp}.xlsx"
    output_path = os.path.abspath(os.path.join(output_dir, filename))

    # Save Excel file
    with open(output_path, "wb") as f:
        f.write(output.getbuffer())

    return filename  # just filename, not full path
