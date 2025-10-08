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
    df['Pharmacy_LastPickupdate2'] = pd.to_datetime(df['Pharmacy_LastPickupdate'], errors='coerce', dayfirst=True).fillna(pd.to_datetime('1900'))

    #Fill zero if the column contains number greater than 180
    df['DaysOfARVRefill2'] = df['DaysOfARVRefill'].apply(lambda x: 0 if x > 180 else x)

    # Calculate the 'NextAppt' column by adding the 'DaysOfARVRefill2' to 'Pharmacy_LastPickupdate2'
    df['NextAppt'] = df['Pharmacy_LastPickupdate2'] + pd.to_timedelta(df['DaysOfARVRefill2'], unit='D') 

    df.insert(0, 'S/N.', df.index + 1)
        
    #extract actives in both current and previous quarter
    df = df[(df['CurrentARTStatus'] == 'Active') & (df['ARTStatus_PreviousQuarter'] =='Active')]

    today = pd.to_datetime(endDate)
    endOfThisQuarter = pd.to_datetime(today) + pd.tseries.offsets.QuarterEnd(0)
    logging.info(f"End of Current Quarter: {endOfThisQuarter}")    
    
    df['end_date'] = endOfThisQuarter
    df['ARTStartDate2'] = pd.to_datetime(df['ARTStartDate'].fillna('1900'))
    #df['Pharmacy_LastPickupdate2'] = pd.to_datetime(df['Pharmacy_LastPickupdate'], errors='coerce')

    # Function to calculate difference in months
    def date_diff_in_months2(date1, date2):
        return (date2.year - date1.year) * 12 + date2.month - date1.month - (1 if date2.day < date1.day else 0)
    
    #function to calculate current age and mirror excel datedif function
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
    
    #print(repr(df.loc[df['DOB'].str.contains('1976')]['DOB'].iloc[0]))
    df['DOB'] = df['DOB'].astype(str).str.strip()
    df['DOB'] = pd.to_datetime(df['DOB'], errors='coerce', infer_datetime_format=True).fillna(pd.to_datetime('1900'))
    
    df['Age'] = calculate_age_vectorized(df, 'DOB', ref_date=end_date)

    # Apply the function to the DataFrame
    df['durationOnART'] = df.apply(lambda row: date_diff_in_months2(row['ARTStartDate2'], row['end_date']), axis=1)
    
    # Flag those who are already ≥ 6 months as at today
    df['eligible_today'] = df['ARTStartDate2'].apply(lambda d: date_diff_in_months2(d, today) >= 6)
             
    #drop irrelevant columns and extract actual VL eligible from dataframe
    df = df.drop(['end_date','ARTStartDate2'], axis=1)
    df = df[df['durationOnART'] >= 6]

    # Get the current date
    #today = pd.Timestamp('today')
    today = pd.to_datetime(endDate)
    currentWeek = today.isocalendar().week
    last_30_days = today - pd.Timedelta(days=29)
    next_30_days = today + pd.Timedelta(days=29)

    # Get the last day of the first quarter in the last one year
    last_year = today - pd.DateOffset(years=1)
    first_quarter_last_year = pd.to_datetime(last_year) + pd.tseries.offsets.QuarterEnd(0)
    
    # End of the quarter in the last 6 months
    six_months_ago = today - pd.DateOffset(months=6)
    end_of_quarter_last_6_months = pd.to_datetime(six_months_ago) + pd.tseries.offsets.QuarterEnd(0)

    logging.info(f"First quarter end from last year: {first_quarter_last_year}")
    
    #convert NextAppt to datetime
    df['NextAppt'] = pd.to_datetime(df['NextAppt'])

    #3rd 95 columns integration
    #df['validVlResult'] = df.apply(lambda row: 'Valid Result' if (row['DateResultReceivedFacility'] > first_quarter_last_year and row['LastDateOfSampleCollection'] > first_quarter_last_year) else 'Invalid Result', axis=1)
    df['validVlResult'] = df.apply(lambda row: 'Valid Result' if (((row['Age'] >= 15 and row['DateResultReceivedFacility'] > first_quarter_last_year) or (row['Age'] < 15 and row['DateResultReceivedFacility'] > six_months_ago)) and ((row['Age'] >= 15 and row['LastDateOfSampleCollection'] > first_quarter_last_year) or (row['Age'] < 15 and row['LastDateOfSampleCollection'] > six_months_ago))) else 'Invalid Result', axis=1)
    #df['validVlSampleCollection'] = df.apply(lambda row: 'Valid SC' if row['LastDateOfSampleCollection'] > first_quarter_last_year else 'Invalid SC', axis=1)
    df['validVlSampleCollection'] = df.apply(lambda row: 'Valid SC' if ((row['Age'] >= 15 and row['LastDateOfSampleCollection'] > first_quarter_last_year) or (row['Age'] < 15 and row['LastDateOfSampleCollection'] > six_months_ago)) else 'Invalid SC', axis=1)
    df['vlSCGap'] = df.apply(lambda row: 'SC Gap' if row['validVlSampleCollection'] == 'Invalid SC' else 'Not SC Gap', axis=1)
    df['vlWKMissedSC'] = df.apply(lambda row: 'vlWKMissedSC' if ((row['vlSCGap'] == 'SC Gap') and (row['Pharmacy_LastPickupdate2'].isocalendar().week == currentWeek) and row['eligible_today']) else 'NotvlWKMissedSC', axis=1)
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
    df_vlWKMissedSC = process_Linelist(df, 'vlWKMissedSC', 'vlWKMissedSC', columns_to_select2)        
    
    df_active_Eligible = df[df['CurrentARTStatus']=="Active"]

    df_active_Eligible['vlEligible'] = df_active_Eligible['CurrentARTStatus'].apply(lambda x: 1 if x == "Active" else 0)
    df_active_Eligible['validVlResult_valid'] = df_active_Eligible['validVlResult'].apply(lambda x: 1 if x == "Valid Result" else 0)
    df_active_Eligible['validVlSampleCollection'] = df_active_Eligible['validVlSampleCollection'].apply(lambda x: 1 if x == "Valid SC" else 0)
    df_active_Eligible['vlSCGap'] = df_active_Eligible['vlSCGap'].apply(lambda x: 1 if x == "SC Gap" else 0)
    df_active_Eligible['PendingResult'] = df_active_Eligible['PendingResult'].apply(lambda x: 1 if x == "Pending" else 0)
    df_active_Eligible['last30daysmissedSC'] = df_active_Eligible['last30daysmissedSC'].apply(lambda x: 1 if x == "Missed SC" else 0)
    df_active_Eligible['expNext30daysdueforSC'] = df_active_Eligible['expNext30daysdueforSC'].apply(lambda x: 1 if x == "due for SC" else 0)
    df_active_Eligible['Suppressed'] = df_active_Eligible['Suppression'].apply(lambda x: 1 if x == "Suppressed" else 0)
    df_active_Eligible['vlWKMissedSC'] = df_active_Eligible['vlWKMissedSC'].apply(lambda x: 1 if x == "vlWKMissedSC" else 0)

    result = df_active_Eligible.groupby(['LGA','FacilityName'])[['vlEligible', 'validVlResult_valid', 'validVlSampleCollection','vlSCGap', 'PendingResult','last30daysmissedSC','expNext30daysdueforSC','Suppressed','vlWKMissedSC']].sum().reset_index()

    result['vl_result_rate'] = ((result['validVlResult_valid'] / result['vlEligible'])).round(4)
    result['sample_collection_rate'] = ((result['validVlSampleCollection'] / result['vlEligible'])).round(4)
    result['suppression_rate'] = ((result['Suppressed'] / result['validVlResult_valid'])).round(4)
    
    result = result[['LGA','FacilityName','vlEligible','validVlResult_valid','Suppressed','vl_result_rate','suppression_rate','validVlSampleCollection','sample_collection_rate','vlSCGap','PendingResult','last30daysmissedSC','expNext30daysdueforSC','vlWKMissedSC']]

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
                                    'vlWKMissedSC': 'Weekly VL SC Missed Oppurtunity',
                                    'Suppression': 'Suppressed',})

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
        "WK SC Missed Oppurtunity": df_vlWKMissedSC,
        "3RD 95 SUMMARY": result,
        #"Sheet3": df3,
        # Add more dataframes and sheet names as needed
    }
    
    # Strip time from datetime columns to ensure clean yyyy-mm-dd export
    for df_name, df_data in dataframes.items():
        for col in df_data.columns:
            if pd.api.types.is_datetime64_any_dtype(df_data[col]):
                df_data[col] = df_data[col].dt.date

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
            worksheet.set_column('F:F', None, percentage_format)
            worksheet.set_column('G:G', None, percentage_format)
            worksheet.set_column('I:I', None, percentage_format)
            
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
                if col_letter == 'F':  # Column F to be calculated as D / C
                    worksheet.write_formula(total_row, col_num, f'SUM(D3:D{total_row}) / SUM(C3:C{total_row})', percentage_format2)
                elif col_letter == 'G':  # Column G to be calculated as E / C
                    worksheet.write_formula(total_row, col_num, f'SUM(E3:E{total_row}) / SUM(D3:D{total_row})', percentage_format2)
                elif col_letter == 'I':  # Column I to be calculated as H / C
                    worksheet.write_formula(total_row, col_num, f'SUM(H3:H{total_row}) / SUM(C3:C{total_row})', percentage_format2)
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
        df['Pharmacy_LastPickupdate2'] = pd.to_datetime(df['Pharmacy_LastPickupdate'], errors='coerce', dayfirst=True).fillna(pd.to_datetime('1900'))

        #Fill zero if the column contains number greater than 180
        df['DaysOfARVRefill2'] = df['DaysOfARVRefill'].apply(lambda x: 0 if x > 180 else x)

        # Calculate the 'NextAppt' column by adding the 'DaysOfARVRefill2' to 'Pharmacy_LastPickupdate2'
        df['NextAppt'] = df['Pharmacy_LastPickupdate2'] + pd.to_timedelta(df['DaysOfARVRefill2'], unit='D')     
        
        df.insert(0, 'S/N.', df.index + 1)
            
        #extract actives in both current and previous quarter
        df = df[(df['CurrentARTStatus'] == 'Active') & (df['ARTStatus_PreviousQuarter'] =='Active')]

        today = pd.to_datetime(endDate)
        currentWeek = today.isocalendar().week
        
        endOfThisQuarter = pd.to_datetime(today) + pd.tseries.offsets.QuarterEnd(0)
        logging.info(f"End of Current Quarter: {endOfThisQuarter}")    
        
        df['end_date'] = endOfThisQuarter
        df['ARTStartDate2'] = pd.to_datetime(df['ARTStartDate'].fillna('1900'))

        # Function to calculate difference in months
        def date_diff_in_months2(date1, date2):
            return (date2.year - date1.year) * 12 + date2.month - date1.month - (1 if date2.day < date1.day else 0)

        # Apply the function to the DataFrame
        df['durationOnART'] = df.apply(lambda row: date_diff_in_months2(row['ARTStartDate2'], row['end_date']), axis=1)
        
        # Flag those who are already ≥ 6 months as at today
        df['eligible_today'] = df['ARTStartDate2'].apply(lambda d: date_diff_in_months2(d, today) >= 6)

        #drop irrelevant columns and extract actual VL eligible from dataframe
        df = df.drop(['end_date','ARTStartDate2'], axis=1)
        df = df[df['durationOnART'] >= 6]

        # Get the current date
        #today = pd.Timestamp('today')
        today = pd.to_datetime(endDate)
        last_30_days = today - pd.Timedelta(days=29)
        next_30_days = today + pd.Timedelta(days=29)
        
        #function to calculate current age and mirror excel datedif function
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
        
        #print(repr(df.loc[df['DOB'].str.contains('1976')]['DOB'].iloc[0]))
        df['DOB'] = df['DOB'].astype(str).str.strip()
        df['DOB'] = pd.to_datetime(df['DOB'], errors='coerce', infer_datetime_format=True).fillna(pd.to_datetime('1900'))
        
        df['Age'] = calculate_age_vectorized(df, 'DOB', ref_date=end_date)

        # Get the last day of the first quarter in the last one year
        last_year = today - pd.DateOffset(years=1)
        first_quarter_last_year = pd.to_datetime(last_year) + pd.tseries.offsets.QuarterEnd(0)
        
        six_months_ago = today - pd.DateOffset(months=6)
        end_of_quarter_last_6_months = pd.to_datetime(six_months_ago) + pd.tseries.offsets.QuarterEnd(0)

        print(first_quarter_last_year)     
        
        #convert VL dates to datetime
        df['DateResultReceivedFacility'] = pd.to_datetime(df['DateResultReceivedFacility'])
        df['LastDateOfSampleCollection'] = pd.to_datetime(df['LastDateOfSampleCollection'])
        df['Pharmacy_LastPickupdate'] = pd.to_datetime(df['Pharmacy_LastPickupdate'])
        df['NextAppt'] = pd.to_datetime(df['NextAppt'])

        #3rd 95 columns integration
        df['validVlResult'] = df.apply(lambda row: 'Valid Result' if (((row['Age'] >= 15 and row['DateResultReceivedFacility'] > first_quarter_last_year) or (row['Age'] < 15 and row['DateResultReceivedFacility'] > six_months_ago)) and ((row['Age'] >= 15 and row['LastDateOfSampleCollection'] > first_quarter_last_year) or (row['Age'] < 15 and row['LastDateOfSampleCollection'] > six_months_ago))) else 'Invalid Result', axis=1)
        df['validVlSampleCollection'] = df.apply(lambda row: 'Valid SC' if ((row['Age'] >= 15 and row['LastDateOfSampleCollection'] > first_quarter_last_year) or (row['Age'] < 15 and row['LastDateOfSampleCollection'] > six_months_ago)) else 'Invalid SC', axis=1)
        #df['validVlResult'] = df.apply(lambda row: 'Valid Result' if (row['DateResultReceivedFacility'] > first_quarter_last_year and row['LastDateOfSampleCollection'] > first_quarter_last_year) else 'Invalid Result', axis=1)
        #df['validVlSampleCollection'] = df.apply(lambda row: 'Valid SC' if row['LastDateOfSampleCollection'] > first_quarter_last_year else 'Invalid SC', axis=1)
        df['vlSCGap'] = df.apply(lambda row: 'SC Gap' if row['validVlSampleCollection'] == 'Invalid SC' else 'Not SC Gap', axis=1)
        df['vlWKMissedSC'] = df.apply(lambda row: 'vlWKMissedSC' if ((row['vlSCGap'] == 'SC Gap') and (row['Pharmacy_LastPickupdate2'].isocalendar().week == currentWeek) and row['eligible_today']) else 'NotvlWKMissedSC', axis=1)
        df['PendingResult'] = df.apply(lambda row: 'Pending' if ((row['validVlSampleCollection'] == 'Valid SC') & (row['validVlResult'] == 'Invalid Result')) else 'Not pending', axis=1)    
        df['last30daysmissedSC'] = df.apply(lambda row: 'Missed SC' if ((row['vlSCGap'] == 'SC Gap') & (row['Pharmacy_LastPickupdate'] >= last_30_days) & (row['Pharmacy_LastPickupdate'] < today)) else 'Not Missed SC', axis=1)  
        df['expNext30daysdueforSC'] = df.apply(lambda row: 'due for SC' if ((row['vlSCGap'] == 'SC Gap') & (row['NextAppt'] >= today) & (row['NextAppt'] < next_30_days)) else 'Not due for SC', axis=1)  
        #df['Suppression'] = df.apply(lambda row: 'Suppressed' if ((row['validVlResult'] == 'Valid Result') & (row['CurrentViralLoad2'] < 1000)) else 'Not Suppressed', axis=1)
        df['Suppression'] = df.apply(lambda row: 'Suppressed' if ((row['validVlResult'] == 'Valid Result') & (row['CurrentViralLoad'] < 1000)) else ('Unsuppressed' if ((row['validVlResult'] == 'Valid Result') & (row['CurrentViralLoad'] > 1000)) else 'Invalid Result'), axis=1)
        #df.to_excel("3rd95.xlsx")
        
        df_sc_Gap = process_Linelist(df, 'vlSCGap', 'SC Gap', columns_to_select, sort_by='CaseManager')
        df_pending_results = process_Linelist(df, 'PendingResult', 'Pending', columns_to_select, sort_by='CaseManager')
        df_last30daysmissedSC  = process_Linelist(df, 'last30daysmissedSC', 'Missed SC', columns_to_select, sort_by='CaseManager')
        df_expNext30daysdueforSC = process_Linelist(df, 'expNext30daysdueforSC', 'due for SC', columns_to_select, sort_by='CaseManager')
        df_Suppression = process_Linelist(df, 'Suppression', 'Unsuppressed', columns_to_select, sort_by='CaseManager')  
        df_vlWKMissedSC = process_Linelist(df, 'vlWKMissedSC', 'vlWKMissedSC', columns_to_select, sort_by='CaseManager')   
        
        df_active_Eligible = df[df['CurrentARTStatus']=="Active"]

        df_active_Eligible['vlEligible'] = df_active_Eligible['CurrentARTStatus'].apply(lambda x: 1 if x == "Active" else 0)
        df_active_Eligible['validVlResult_valid'] = df_active_Eligible['validVlResult'].apply(lambda x: 1 if x == "Valid Result" else 0)
        df_active_Eligible['validVlSampleCollection'] = df_active_Eligible['validVlSampleCollection'].apply(lambda x: 1 if x == "Valid SC" else 0)
        df_active_Eligible['vlSCGap'] = df_active_Eligible['vlSCGap'].apply(lambda x: 1 if x == "SC Gap" else 0)
        df_active_Eligible['PendingResult'] = df_active_Eligible['PendingResult'].apply(lambda x: 1 if x == "Pending" else 0)
        df_active_Eligible['last30daysmissedSC'] = df_active_Eligible['last30daysmissedSC'].apply(lambda x: 1 if x == "Missed SC" else 0)
        df_active_Eligible['expNext30daysdueforSC'] = df_active_Eligible['expNext30daysdueforSC'].apply(lambda x: 1 if x == "due for SC" else 0)
        df_active_Eligible['Suppressed'] = df_active_Eligible['Suppression'].apply(lambda x: 1 if x == "Suppressed" else 0)
        df_active_Eligible['vlWKMissedSC'] = df_active_Eligible['vlWKMissedSC'].apply(lambda x: 1 if x == "vlWKMissedSC" else 0)

        #result = df_active_Eligible.groupby(['LGA','FacilityName','CaseManager'])[['vlEligible', 'validVlResult_valid', 'validVlSampleCollection','vlSCGap', 'PendingResult']].sum().reset_index()
        result = df_active_Eligible.groupby(['LGA','FacilityName','CaseManager'])[['vlEligible', 'validVlResult_valid', 'validVlSampleCollection','vlSCGap', 'PendingResult','last30daysmissedSC','expNext30daysdueforSC','Suppressed','vlWKMissedSC']].sum().reset_index()

        result['vl_result_rate'] = ((result['validVlResult_valid'] / result['vlEligible'])).round(4)
        result['sample_collection_rate'] = ((result['validVlSampleCollection'] / result['vlEligible'])).round(4)
        result['suppression_rate'] = ((result['Suppressed'] / result['validVlResult_valid'])).round(4)
        
        result = result[['LGA','FacilityName','CaseManager','vlEligible','validVlResult_valid','Suppressed','vl_result_rate','suppression_rate','validVlSampleCollection','sample_collection_rate','vlSCGap','PendingResult','last30daysmissedSC','expNext30daysdueforSC','vlWKMissedSC']]

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
                                        'vlWKMissedSC': 'Weekly VL SC Missed Oppurtunity',
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
            "WK SC Missed Oppurtunity": df_vlWKMissedSC,
            "3RD 95 SUMMARY": result,
            #"Sheet3": df3,
            # Add more dataframes and sheet names as needed
        }
        
        # Strip time from datetime columns to ensure clean yyyy-mm-dd export
        for df_name, df_data in dataframes.items():
            for col in df_data.columns:
                if pd.api.types.is_datetime64_any_dtype(df_data[col]):
                    df_data[col] = df_data[col].dt.date

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
