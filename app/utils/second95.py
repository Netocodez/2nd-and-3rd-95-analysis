# app.py
from flask import Flask, render_template, request, jsonify, send_file
import pandas as pd
import os
import logging
import traceback
from io import BytesIO
from datetime import datetime
from .emr_processor import process_Linelist, columns_to_select, columns_to_select2


def second95(df, endDate):
    
    endDate = pd.to_datetime(endDate)
    
    # Ensure the 'Pharmacy_LastPickupdate' column is in datetime format and fill NaNs with a specific date
    df['Pharmacy_LastPickupdate2'] = pd.to_datetime(df['Pharmacy_LastPickupdate'], errors='coerce').fillna(pd.to_datetime('1900'))

    #Fill zero if the column contains number greater than 180
    df['DaysOfARVRefill2'] = df['DaysOfARVRefill'].apply(lambda x: 0 if x > 180 else x)

    # Calculate the 'NextAppt' column by adding the 'DaysOfARVRefill2' to 'Pharmacy_LastPickupdate2'
    df['NextAppt'] = df['Pharmacy_LastPickupdate2'] + pd.to_timedelta(df['DaysOfARVRefill2'], unit='D') 
    df['IITDate2'] = (df['NextAppt'] + pd.Timedelta(days=29)).fillna('1900')
    
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


def Second95R(df, dfbaseline, endDate): 
    
        endDate = pd.to_datetime(endDate)

        #declaver variables for analysis
        today = pd.Timestamp(endDate)
        currentyear = today.year
        previousyear = today.year-1
        currentmonthyear = str(today.month) + '_' + str(today.year)
        sevenDaysIIT = today + pd.Timedelta(days=7)
        currentWeekYear = f"{today.isocalendar().week}_{today.year}"
        print(currentWeekYear)
        
        #Bring in Values from baseline ART line list
        df['BaselineCurrentARTStatus']=df['uuid'].map(dfbaseline.set_index('uuid')['CurrentARTStatus'])
        df['BaselinePharmacy_LastPickupdate']=df['uuid'].map(dfbaseline.set_index('uuid')['Pharmacy_LastPickupdate'])
        df['BaselineDaysOfARVRefill']=df['uuid'].map(dfbaseline.set_index('uuid')['DaysOfARVRefill'])
        
        #Calculate Next Appointment for Added baseline columns
        df['BaselinePharmacy_LastPickupdate'] = pd.to_datetime(df['BaselinePharmacy_LastPickupdate'], errors='coerce').fillna(pd.to_datetime('1900'))
        df['BaselineDaysOfARVRefill'] = pd.to_numeric(df['BaselineDaysOfARVRefill']).apply(lambda x: 0 if x > 180 else x)
        df['BaselineNextAppt'] = pd.to_datetime(df['BaselinePharmacy_LastPickupdate'], errors='coerce', dayfirst=True) + pd.to_timedelta(df['BaselineDaysOfARVRefill'], unit='D') 
        
        #Create week year columns for baseline last pickup date and baseline next appointment
        df['BaselinePharmacy_LastPickupdate_week_year'] = df['BaselinePharmacy_LastPickupdate'].dt.isocalendar().week.astype(str) + '_' + df['BaselinePharmacy_LastPickupdate'].dt.year.astype(str)
        #df['BaselineNextAppt_week_year'] = df['BaselineNextAppt'].dt.isocalendar().week.astype(str) + '_' + df['BaselineNextAppt'].dt.year.astype(int).astype(str)
        df['BaselineNextAppt_week_year'] = df['BaselineNextAppt'].apply(
                lambda x: f"{x.isocalendar().week}_{x.year}" if pd.notna(x) else ""
            )
         
        # Ensure the 'Pharmacy_LastPickupdate' column is in datetime format and fill NaNs with a specific date
        df['Pharmacy_LastPickupdate2'] = pd.to_datetime(df['Pharmacy_LastPickupdate'], errors='coerce').fillna(pd.to_datetime('1900'))

        #Fill zero if the column contains number greater than 180
        df['DaysOfARVRefill2'] = df['DaysOfARVRefill'].apply(lambda x: 0 if x > 180 else x)
        
        #create week_year column for current pharmacy refill
        df['Pharmacy_LastPickupdate2_week_year'] = df['Pharmacy_LastPickupdate2'].dt.isocalendar().week.astype(str) + '_' + df['Pharmacy_LastPickupdate2'].dt.year.astype(str)

        # Calculate the 'NextAppt' column by adding the 'DaysOfARVRefill2' to 'Pharmacy_LastPickupdate2'
        df['NextAppt'] = df['Pharmacy_LastPickupdate2'] + pd.to_timedelta(df['DaysOfARVRefill2'], unit='D') 
        df['IITDate2'] = (df['NextAppt'] + pd.Timedelta(days=29)).fillna('1900')
        
        #Clean Next Appointments
        df['NextAppt'] = pd.to_datetime(df['NextAppt'], errors='coerce').fillna(pd.to_datetime('1900'))
        df['BaselineNextAppt'] = pd.to_datetime(df['BaselineNextAppt'], errors='coerce').fillna(pd.to_datetime('1900'))
         
        df.insert(0, 'S/N.', df.index + 1)
        
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
        df['currentweekexpected'] = df.apply(lambda row: 'currentweekexpected' if ((row['BaselineNextAppt_week_year'] == currentWeekYear) & ((row['CurrentARTStatus'] == 'Active')) & ((row['BaselineCurrentARTStatus'] == 'Active'))) else 'notcurrentweekexpected', axis=1)
        df['weeklyexpectedrefilled'] = df.apply(lambda row: 'weeklyexpectedrefilled' if ((row['currentweekexpected'] == 'currentweekexpected') & ((row['Pharmacy_LastPickupdate2'] > row['BaselinePharmacy_LastPickupdate']) & (row['NextAppt'] > row['BaselineNextAppt']))) else 'notweeklyexpectedrefilled', axis=1)
        df['pendingweeklyrefill'] = df.apply(lambda row: 'pendingweeklyrefill' if ((row['BaselineNextAppt_week_year'] == currentWeekYear) & ((row['CurrentARTStatus'] == 'Active')) & ((row['BaselineCurrentARTStatus'] == 'Active')) & ((row['weeklyexpectedrefilled'] == 'notweeklyexpectedrefilled'))) else 'notpendingweeklyrefill', axis=1)
        #df.to_excel("refillrate.xlsx")
        
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
        
        dfpendingweeklyrefill = df[df['pendingweeklyrefill'] == 'pendingweeklyrefill'].reset_index()
        dfpendingweeklyrefill.insert(0, 'S/N', dfpendingweeklyrefill.index + 1)
        
        #Trime Line lists
        dfCurrentYearIIT = dfCurrentYearIIT[["S/N", "State", "LGA", "FacilityName", "PEPID", "PatientHospitalNo", "uuid", "Sex", "Current_Age", "Surname", "Firstname", "MaritalStatus", "PhoneNo", "Address", "State_of_Residence", "LGA_of_Residence", "DateConfirmedHIV+", "ARTStartDate", "Pharmacy_LastPickupdate", "DaysOfARVRefill", "CurrentARTRegimen", "NextAppt", "CurrentPregnancyStatus", "CurrentViralLoad", "DateResultReceivedFacility", "Alphanumeric_Viral_Load_Result", "LastDateOfSampleCollection", "Outcomes", "Outcomes_Date", "IIT_Date", "CurrentARTStatus", "First_TPT_Pickupdate", "Current_TPT_Received", "PBS_Capturee", "PBS_Capture_Date", "Date_Generated"]]
        dfpreviousyearIIT = dfpreviousyearIIT[["S/N", "State", "LGA", "FacilityName", "PEPID", "PatientHospitalNo", "uuid", "Sex", "Current_Age", "Surname", "Firstname", "MaritalStatus", "PhoneNo", "Address", "State_of_Residence", "LGA_of_Residence", "DateConfirmedHIV+", "ARTStartDate", "Pharmacy_LastPickupdate", "DaysOfARVRefill", "CurrentARTRegimen", "NextAppt", "CurrentPregnancyStatus", "CurrentViralLoad", "DateResultReceivedFacility", "Alphanumeric_Viral_Load_Result", "LastDateOfSampleCollection", "Outcomes", "Outcomes_Date", "IIT_Date", "CurrentARTStatus", "First_TPT_Pickupdate", "Current_TPT_Received", "PBS_Capturee", "PBS_Capture_Date", "Date_Generated"]]
        dfImminentIIT = dfImminentIIT[["S/N", "State", "LGA", "FacilityName", "PEPID", "PatientHospitalNo", "uuid", "Sex", "Current_Age", "Surname", "Firstname", "MaritalStatus", "PhoneNo", "Address", "State_of_Residence", "LGA_of_Residence", "DateConfirmedHIV+", "ARTStartDate", "Pharmacy_LastPickupdate", "DaysOfARVRefill", "CurrentARTRegimen", "NextAppt", "CurrentPregnancyStatus", "CurrentViralLoad", "DateResultReceivedFacility", "Alphanumeric_Viral_Load_Result", "LastDateOfSampleCollection", "Outcomes", "Outcomes_Date", "IIT_Date", "CurrentARTStatus", "First_TPT_Pickupdate", "Current_TPT_Received", "PBS_Capturee", "PBS_Capture_Date", "Date_Generated"]]
        dfsevendaysIIT = dfsevendaysIIT[["S/N", "State", "LGA", "FacilityName", "PEPID", "PatientHospitalNo", "uuid", "Sex", "Current_Age", "Surname", "Firstname", "MaritalStatus", "PhoneNo", "Address", "State_of_Residence", "LGA_of_Residence", "DateConfirmedHIV+", "ARTStartDate", "Pharmacy_LastPickupdate", "DaysOfARVRefill", "CurrentARTRegimen", "NextAppt", "CurrentPregnancyStatus", "CurrentViralLoad", "DateResultReceivedFacility", "Alphanumeric_Viral_Load_Result", "LastDateOfSampleCollection", "Outcomes", "Outcomes_Date", "IIT_Date", "CurrentARTStatus", "First_TPT_Pickupdate", "Current_TPT_Received", "PBS_Capturee", "PBS_Capture_Date", "Date_Generated"]]
        dfcurrentmonthexpected = dfcurrentmonthexpected[["S/N", "State", "LGA", "FacilityName", "PEPID", "PatientHospitalNo", "uuid", "Sex", "Current_Age", "Surname", "Firstname", "MaritalStatus", "PhoneNo", "Address", "State_of_Residence", "LGA_of_Residence", "DateConfirmedHIV+", "ARTStartDate", "Pharmacy_LastPickupdate", "DaysOfARVRefill", "CurrentARTRegimen", "NextAppt", "CurrentPregnancyStatus", "CurrentViralLoad", "DateResultReceivedFacility", "Alphanumeric_Viral_Load_Result", "LastDateOfSampleCollection", "Outcomes", "Outcomes_Date", "IIT_Date", "CurrentARTStatus", "First_TPT_Pickupdate", "Current_TPT_Received", "PBS_Capturee", "PBS_Capture_Date", "Date_Generated"]]
        dfpendingweeklyrefill = dfpendingweeklyrefill[["S/N", "State", "LGA", "FacilityName", "PEPID", "PatientHospitalNo", "uuid", "Sex", "Current_Age", "Surname", "Firstname", "MaritalStatus", "PhoneNo", "Address", "State_of_Residence", "LGA_of_Residence", "DateConfirmedHIV+", "ARTStartDate", "Pharmacy_LastPickupdate", "DaysOfARVRefill", "CurrentARTRegimen", "NextAppt", "CurrentPregnancyStatus", "CurrentViralLoad", "DateResultReceivedFacility", "Alphanumeric_Viral_Load_Result", "LastDateOfSampleCollection", "Outcomes", "Outcomes_Date", "IIT_Date", "CurrentARTStatus", "First_TPT_Pickupdate", "Current_TPT_Received", "PBS_Capturee", "PBS_Capture_Date", "Date_Generated"]]
        
        df2nd95Summary = df

        df2nd95Summary['ActiveClients'] = df2nd95Summary['CurrentARTStatus'].apply(lambda x: 1 if x == "Active" else 0)
        df2nd95Summary['currentYearIIT'] = df2nd95Summary['CurrentYearIIT'].apply(lambda x: 1 if x == "CurrentYearIIT" else 0)
        df2nd95Summary['previousyearIIT'] = df2nd95Summary['previousyearIIT'].apply(lambda x: 1 if x == "previousyearIIT" else 0)
        df2nd95Summary['ImminentIIT'] = df2nd95Summary['ImminentIIT'].apply(lambda x: 1 if x == "ImminentIIT" else 0)
        df2nd95Summary['sevendaysIIT'] = df2nd95Summary['sevendaysIIT'].apply(lambda x: 1 if x == "sevendaysIIT" else 0)
        df2nd95Summary['currentmonthexpected'] = df2nd95Summary['currentmonthexpected'].apply(lambda x: 1 if x == "currentmonthexpected" else 0)
        df2nd95Summary['currentweekexpected'] = df2nd95Summary['currentweekexpected'].apply(lambda x: 1 if x == "currentweekexpected" else 0)
        df2nd95Summary['weeklyexpectedrefilled'] = df2nd95Summary['weeklyexpectedrefilled'].apply(lambda x: 1 if x == "weeklyexpectedrefilled" else 0)

        result = df2nd95Summary.groupby(['LGA','FacilityName'])[['ActiveClients', 'currentYearIIT', 'previousyearIIT','ImminentIIT','sevendaysIIT','currentmonthexpected', 'currentweekexpected', 'weeklyexpectedrefilled']].sum().reset_index()
        result_check = ['ActiveClients', 'currentYearIIT', 'previousyearIIT', 'ImminentIIT', 'sevendaysIIT', 'currentmonthexpected', 'currentweekexpected', 'weeklyexpectedrefilled']
        result = result[(result[result_check] != 0).any(axis=1)]
        
        result['Weekly_Refill_Rate'] = ((result['weeklyexpectedrefilled'] / result['currentweekexpected'])).round(4)
        result = result[['LGA','FacilityName','ActiveClients','currentmonthexpected','currentweekexpected','weeklyexpectedrefilled','Weekly_Refill_Rate', 'currentYearIIT', 'previousyearIIT', 'ImminentIIT', 'sevendaysIIT']]


        result = result.rename(columns={'ActiveClients': 'Active Clients',
                                        'currentYearIIT': 'current FY IIT',
                                        'previousyearIIT': 'previous FY IIT',
                                        'ImminentIIT': 'Imminent IIT',
                                        'sevendaysIIT': 'Next 7 days IIT',
                                        'currentmonthexpected': 'EXPECTED THIS MONTH',
                                        'currentweekexpected': 'EXPECTED THIS WEEK',
                                        'weeklyexpectedrefilled': 'REFILLED FROM WEEKLY EXPECTED',
                                        'Weekly_Refill_Rate': '%Weekly Refill Rate',})
        result   
        
        #format and export
        sdate = endDate.strftime("%d-%m-%Y")

        # Create a Pandas Excel writer using XlsxWriter as the engine.
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
            "EXPECTED THIS WEEK": dfpendingweeklyrefill,
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
                
                row_band_format = workbook.add_format({'bg_color': '#F9F9F9'})
                worksheet.conditional_format(2, 0, len(dataframe) + 1, len(dataframe.columns) - 1, 
                                            {'type': 'formula', 'criteria': 'MOD(ROW(), 2) = 0', 'format': row_band_format})
                worksheet.hide_gridlines(2)
                
                column_ranges = {
                    'A:A': 20,  # Example width for columns A to A
                    'B:B': 35,  # Example width for columns B to B
                    #'C:C': 25,  # Example width for columns C to C
                    # Add more column ranges and their widths as needed
                }
                for col_range, width in column_ranges.items():
                    worksheet.set_column(col_range, width)
                    
                worksheet.set_column('G:G', None, percentage_format) #add percentage to column G
                
                # Add total row
                header_format.set_align('center')
                total_row = len(dataframe) + 2
                worksheet.merge_range(total_row, 0, total_row, 1, 'Total', header_format)
                for col_num in range(1, len(dataframe.columns)):
                    col_letter = chr(65 + col_num)  # Convert column number to letter
                    #worksheet.write_formula(total_row, col_num, f'SUM({col_letter}2:{col_letter}{total_row})', header_format)
                    if col_letter == 'G':  # Column G to be calculated as E / C
                        worksheet.write_formula(total_row, col_num, f'SUM(F3:F{total_row}) / SUM(E3:E{total_row})', percentage_format2)
                    #elif col_letter == 'I':  # Column I to be calculated as H / C
                        #worksheet.write_formula(total_row, col_num, f'SUM(H3:H{total_row}) / SUM(C3:C{total_row})', percentage_format2)
                    else:
                        worksheet.write_formula(total_row, col_num, f'SUM({col_letter}3:{col_letter}{total_row})', header_format)
                            
            else:
                # Write the dataframe to the sheet, starting from row 1
                dataframe.to_excel(writer, sheet_name=sheet_name, startrow=1, header=False, index=False)
                
                # Write the column headers with the defined format, starting from row 0
                for col_num, value in enumerate(dataframe.columns.values):
                    worksheet.write(0, col_num, value, header_format)
            
            # Add row band format to specific sheets
            if sheet_name in ["CURRENT FY IIT", "PREVIOUS FY IIT", "IMMINENT IIT", "NEXT 7 DAYS IIT", "EXPECTED THIS MONTH", "EXPECTED THIS WEEK"]:
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

        # Close the Pandas Excel writer and output the Excel file.
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
    

def Second95RCMG(df, dfbaseline, endDate): 
    
        endDate = pd.to_datetime(endDate)
                
        print(dfbaseline)
                
        #add case managers and next appt to the dataframe
        df['CaseManager'] = df['CaseManager'].fillna('UNASSIGNED')
        
        #declaver variables for analysis
        today = pd.Timestamp(endDate)
        currentyear = today.year
        previousyear = today.year-1
        currentmonthyear = str(today.month) + '_' + str(today.year)
        sevenDaysIIT = today + pd.Timedelta(days=7)
        currentWeekYear = f"{today.isocalendar().week}_{today.year}"
        currentWeek = today.isocalendar().week
        print(currentWeekYear)
        
        #Bring in Values from baseline ART line list
        df['BaselineCurrentARTStatus']=df['uuid'].map(dfbaseline.set_index('uuid')['CurrentARTStatus'])
        df['BaselinePharmacy_LastPickupdate']=df['uuid'].map(dfbaseline.set_index('uuid')['Pharmacy_LastPickupdate'])
        df['BaselineDaysOfARVRefill']=df['uuid'].map(dfbaseline.set_index('uuid')['DaysOfARVRefill'])
        
        #Calculate Next Appointment for Added baseline columns
        df['BaselinePharmacy_LastPickupdate'] = pd.to_datetime(df['BaselinePharmacy_LastPickupdate'], errors='coerce', dayfirst=True).fillna(pd.to_datetime('1900'))
        df['BaselineDaysOfARVRefill'] = pd.to_numeric(df['BaselineDaysOfARVRefill'], errors='coerce')
        df['BaselineDaysOfARVRefill'] = df['BaselineDaysOfARVRefill'].apply(lambda x: 0 if pd.notnull(x) and x > 180 else x)
        df['BaselineNextAppt'] = df['BaselinePharmacy_LastPickupdate'] + pd.to_timedelta(df['BaselineDaysOfARVRefill'], unit='D') 
        
        #Create week year columns for baseline last pickup date and baseline next appointment
        df['BaselinePharmacy_LastPickupdate_week_year'] = df['BaselinePharmacy_LastPickupdate'].dt.isocalendar().week.astype(str) + '_' + df['BaselinePharmacy_LastPickupdate'].dt.year.astype(str)
        #df['BaselineNextAppt_week_year'] = df['BaselineNextAppt'].dt.isocalendar().week.astype(str) + '_' + df['BaselineNextAppt'].dt.year.astype(int).astype(str)
        df['BaselineNextAppt_week_year'] = df['BaselineNextAppt'].apply(
                lambda x: f"{x.isocalendar().week}_{x.year}" if pd.notna(x) else ""
            )
         
        # Ensure the 'Pharmacy_LastPickupdate' column is in datetime format and fill NaNs with a specific date
        df['Pharmacy_LastPickupdate2'] = pd.to_datetime(df['Pharmacy_LastPickupdate'], errors='coerce', dayfirst=True).fillna(pd.to_datetime('1900'))

        #Fill zero if the column contains number greater than 180
        df['DaysOfARVRefill2'] = df['DaysOfARVRefill'].apply(lambda x: 0 if x > 180 else x)
        
        #create week_year column for current pharmacy refill
        df['Pharmacy_LastPickupdate2_week_year'] = df['Pharmacy_LastPickupdate2'].dt.isocalendar().week.astype(str) + '_' + df['Pharmacy_LastPickupdate2'].dt.year.astype(str)

        # Calculate the 'NextAppt' column by adding the 'DaysOfARVRefill2' to 'Pharmacy_LastPickupdate2'
        df['NextAppt'] = df['Pharmacy_LastPickupdate2'] + pd.to_timedelta(df['DaysOfARVRefill2'], unit='D') 
        df['IITDate2'] = (df['NextAppt'] + pd.Timedelta(days=29)).fillna('1900')
         
        #Clean Next Appointments
        df['NextAppt'] = pd.to_datetime(df['NextAppt'], errors='coerce').fillna(pd.to_datetime('1900'))
        df['BaselineNextAppt'] = pd.to_datetime(df['BaselineNextAppt'], errors='coerce').fillna(pd.to_datetime('1900'))
         
        df.insert(0, 'S/N.', df.index + 1)
        
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
        df['currentweekexpected'] = df.apply(lambda row: 'currentweekexpected' if ((row['BaselineNextAppt_week_year'] == currentWeekYear) & ((row['CurrentARTStatus'] == 'Active')) & ((row['BaselineCurrentARTStatus'] == 'Active'))) else 'notcurrentweekexpected', axis=1)
        df['weeklyexpectedrefilled'] = df.apply(lambda row: 'weeklyexpectedrefilled' if ((row['currentweekexpected'] == 'currentweekexpected') & ((row['Pharmacy_LastPickupdate2'] > row['BaselinePharmacy_LastPickupdate']) & (row['NextAppt'] > row['BaselineNextAppt']))) else 'notweeklyexpectedrefilled', axis=1)
        df['pendingweeklyrefill'] = df.apply(lambda row: 'pendingweeklyrefill' if ((row['BaselineNextAppt_week_year'] == currentWeekYear) & ((row['CurrentARTStatus'] == 'Active')) & ((row['BaselineCurrentARTStatus'] == 'Active')) & ((row['weeklyexpectedrefilled'] == 'notweeklyexpectedrefilled'))) else 'notpendingweeklyrefill', axis=1)
        #df.to_excel("refillrate.xlsx")
        
        #Filter and process line list        
        columns_to_select = ["S/N", "State", "LGA", "FacilityName", "PEPID", "PatientHospitalNo", "uuid", "Sex", "Current_Age", "Surname", "Firstname", "MaritalStatus", "PhoneNo", "Address", "State_of_Residence", "LGA_of_Residence", "DateConfirmedHIV+", "ARTStartDate", "Pharmacy_LastPickupdate", "DaysOfARVRefill", "CurrentARTRegimen", "NextAppt", "CurrentPregnancyStatus", "CurrentViralLoad", "DateResultReceivedFacility", "Alphanumeric_Viral_Load_Result", "LastDateOfSampleCollection", "Outcomes", "Outcomes_Date", "IIT_Date", "CurrentARTStatus", "First_TPT_Pickupdate", "Current_TPT_Received", "PBS_Capturee", "PBS_Capture_Date", "Date_Generated", "CaseManager"]
        
        #function to process line list
        def process_Linelist(df, column_name, filter_value, columns_to_select):
            df_filtered = df[df[column_name] == filter_value].reset_index(drop=True)
            df_sorted = df_filtered.sort_values(by='CaseManager', ascending=True).reset_index(drop=True)
            df_sorted.insert(0, 'S/N', df_sorted.index + 1)
            df_selected = df_sorted[columns_to_select]
            return df_selected
        
        #apply function to process line list
        dfCurrentYearIIT = process_Linelist(df, 'CurrentYearIIT', 'CurrentYearIIT', columns_to_select)
        dfpreviousyearIIT = process_Linelist(df, 'previousyearIIT', 'previousyearIIT', columns_to_select)
        dfImminentIIT = process_Linelist(df, 'ImminentIIT', 'ImminentIIT', columns_to_select)
        dfsevendaysIIT = process_Linelist(df, 'sevendaysIIT', 'sevendaysIIT', columns_to_select)
        dfcurrentmonthexpected = process_Linelist(df, 'currentmonthexpected', 'currentmonthexpected', columns_to_select)
        dfpendingweeklyrefill = process_Linelist(df, 'pendingweeklyrefill', 'pendingweeklyrefill', columns_to_select)
               
        
        df2nd95Summary = df

        df2nd95Summary['ActiveClients'] = df2nd95Summary['CurrentARTStatus'].apply(lambda x: 1 if x == "Active" else 0)
        df2nd95Summary['currentYearIIT'] = df2nd95Summary['CurrentYearIIT'].apply(lambda x: 1 if x == "CurrentYearIIT" else 0)
        df2nd95Summary['previousyearIIT'] = df2nd95Summary['previousyearIIT'].apply(lambda x: 1 if x == "previousyearIIT" else 0)
        df2nd95Summary['ImminentIIT'] = df2nd95Summary['ImminentIIT'].apply(lambda x: 1 if x == "ImminentIIT" else 0)
        df2nd95Summary['sevendaysIIT'] = df2nd95Summary['sevendaysIIT'].apply(lambda x: 1 if x == "sevendaysIIT" else 0)
        df2nd95Summary['currentmonthexpected'] = df2nd95Summary['currentmonthexpected'].apply(lambda x: 1 if x == "currentmonthexpected" else 0)
        df2nd95Summary['currentweekexpected'] = df2nd95Summary['currentweekexpected'].apply(lambda x: 1 if x == "currentweekexpected" else 0)
        df2nd95Summary['weeklyexpectedrefilled'] = df2nd95Summary['weeklyexpectedrefilled'].apply(lambda x: 1 if x == "weeklyexpectedrefilled" else 0)

        result = df2nd95Summary.groupby(['LGA','FacilityName','CaseManager'])[['ActiveClients', 'currentYearIIT', 'previousyearIIT','ImminentIIT','sevendaysIIT','currentmonthexpected', 'currentweekexpected', 'weeklyexpectedrefilled']].sum().reset_index()
        result_check = ['ActiveClients', 'currentYearIIT', 'previousyearIIT', 'ImminentIIT', 'sevendaysIIT', 'currentmonthexpected', 'currentweekexpected', 'weeklyexpectedrefilled']
        result = result[(result[result_check] != 0).any(axis=1)]
        
        result['Weekly_Refill_Rate'] = ((result['weeklyexpectedrefilled'] / result['currentweekexpected'])).round(4)
        result = result[['LGA','FacilityName','CaseManager','ActiveClients','currentmonthexpected','currentweekexpected','weeklyexpectedrefilled','Weekly_Refill_Rate', 'currentYearIIT', 'previousyearIIT', 'ImminentIIT', 'sevendaysIIT']]


        result = result.rename(columns={'ActiveClients': 'Active Clients',
                                        'currentYearIIT': 'current FY IIT',
                                        'previousyearIIT': 'previous FY IIT',
                                        'ImminentIIT': 'Imminent IIT',
                                        'sevendaysIIT': 'Next 7 days IIT',
                                        'currentmonthexpected': 'EXPECTED THIS MONTH',
                                        'currentweekexpected': 'EXPECTED THIS WEEK',
                                        'weeklyexpectedrefilled': 'REFILLED FROM WEEKLY EXPECTED',
                                        'Weekly_Refill_Rate': '%Weekly Refill Rate',})
        result   
        
        #format and export
        sdate = endDate.strftime("%d-%m-%Y")

        # Create a Pandas Excel writer using XlsxWriter as the engine.
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
            "EXPECTED THIS WEEK": dfpendingweeklyrefill,
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
                
                colorscales = ['%Weekly Refill Rate']
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
                    'C:C': 35,  # Example width for columns B to B
                    # Add more column ranges and their widths as needed
                }
                for col_range, width in column_ranges.items():
                    worksheet.set_column(col_range, width)
                    
                worksheet.set_column('H:H', None, percentage_format) #add percentage to column G
                
                # Add total row
                header_format.set_align('center')
                total_row = len(dataframe) + 2
                worksheet.merge_range(total_row, 0, total_row, 1, 'Total', header_format)
                for col_num in range(1, len(dataframe.columns)):
                    col_letter = chr(65 + col_num)  # Convert column number to letter
                    #worksheet.write_formula(total_row, col_num, f'SUM({col_letter}2:{col_letter}{total_row})', header_format)
                    if col_letter == 'H':  # Column G to be calculated as E / C
                        worksheet.write_formula(total_row, col_num, f'SUM(G3:G{total_row}) / SUM(F3:F{total_row})', percentage_format2)
                    #elif col_letter == 'I':  # Column I to be calculated as H / C
                        #worksheet.write_formula(total_row, col_num, f'SUM(H3:H{total_row}) / SUM(C3:C{total_row})', percentage_format2)
                    else:
                        worksheet.write_formula(total_row, col_num, f'SUM({col_letter}3:{col_letter}{total_row})', header_format)
                            
            else:
                # Write the dataframe to the sheet, starting from row 1
                dataframe.to_excel(writer, sheet_name=sheet_name, startrow=1, header=False, index=False)
                
                # Write the column headers with the defined format, starting from row 0
                for col_num, value in enumerate(dataframe.columns.values):
                    worksheet.write(0, col_num, value, header_format)
            
            # Add row band format to specific sheets
            if sheet_name in ["CURRENT FY IIT", "PREVIOUS FY IIT", "IMMINENT IIT", "NEXT 7 DAYS IIT", "EXPECTED THIS MONTH", "EXPECTED THIS WEEK"]:
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

        # Close the Pandas Excel writer and output the Excel file.
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
