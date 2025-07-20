from tkinter import ttk
import time
import pandas as pd
import numpy as np
from tkcalendar import DateEntry
from datetime import datetime
from dateutil import parser
import tkinter as tk
from tkinter import filedialog
from tkinter import *
import string
import csv
import os

# Declare global variables
start_date = None
end_date = None

def get_selected_date():
    global start_date, end_date
    start_date = cal.get_date()
    end_date = cal2.get_date()
    start_date_label.config(text=f"Start Date: {start_date}")
    end_date_label.config(text=f"End Date: {end_date}")
    startDate = start_date
    endDate = end_date
    todayDate = endDate
    todayDate = str(todayDate)
    return startDate, endDate

#function for cleaning and formating dates 
def parse_date(date):

    if pd.isna(date):  # Handle NaN values
        return pd.NaT
    
    if isinstance(date, pd.Timestamp):  # If already a datetime object
        return date.date()
    
    if isinstance(date, (int, float)):  # Handle Excel serial numbers
        try:
            return pd.to_datetime(date, origin='1899-12-30', unit='D').date()
        except Exception:
            return pd.NaT

    date_formats = ["%Y-%m-%d", "%d/%m/%Y", "%m-%d-%Y", "%d-%m-%Y", "%Y.%m.%d", "%Y-%b-%d"]

    for fmt in date_formats:
        try:
            return pd.to_datetime(date, format=fmt).date()
        except (ValueError, TypeError):
            continue

    try:
        return parser.parse(str(date), fuzzy=True, ignoretz=True).date()
    except (parser.ParserError, ValueError, TypeError):
        return pd.NaT

    
# lamis and nmrs facility names
emrData = {
    'STATE': ['Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra', 'Anambra'],
    'LGA': ['Aguata', 'Aguata', 'Aguata', 'Anambra East', 'Anambra East', 'Anambra East', 'Anambra East', 'Anambra East', 'Anambra East', 'Anambra East', 'Anambra East', 'Anambra East', 'Anambra West', 'Anaocha', 'Anaocha', 'Anaocha', 'Anaocha', 'Anaocha', 'Anaocha', 'Awka North', 'Awka South', 'Awka South', 'Awka South', 'Awka South', 'Awka South', 'Awka South', 'Awka South', 'Awka South', 'Awka South', 'Ayamelum', 'Dunukofia', 'Ekwusigo', 'Ekwusigo', 'Ekwusigo', 'Ekwusigo', 'Ekwusigo', 'Idemili North', 'Idemili North', 'Idemili North', 'Idemili North', 'Idemili North', 'Idemili North', 'Idemili North', 'Idemili South', 'Idemili South', 'Ihiala', 'Ihiala', 'Njikoka', 'Nnewi North', 'Nnewi North', 'Nnewi North', 'Nnewi South', 'Nnewi South', 'Nnewi South', 'Nnewi North', 'Nnewi North', 'Ogbaru', 'Ogbaru', 'Onitsha North', 'Onitsha North', 'Onitsha North', 'Onitsha South', 'Orumba North', 'Orumba South', 'Orumba South', 'Oyi', 'Oyi', 'Oyi', 'Anambra East', 'Ayamelum'],
    'Name on Lamis': ['Catholic Visitation Hospital and Maternity, Umuchu', 'Ekwulobia General Hospital', 'Igboukwu Diocesan Hospital', 'Aguleri Immaculate Heart Hospital', 'an Enugwu Aguleri Primary Health Center', 'an Ikem Nando Primary Health Center', 'an Mama Eliza Maternity Home', 'an Otuocha Maternity and Child Hospital', 'an Umuoba-Anam Primary Health Center', 'Umueri Diocesan Hospital and Maternity', 'Umueri General Hospital', 'an Abube Agu Nando Primary Health Centre', 'an Rex Inversorium Hospital, Mmiata (Referral Health center Oroma-Etiti )', 'Adazi St Joseph Hospital', 'an Akwaeze Primary Health Centre', 'an Ichida Primary Health Centre', 'an neni 1 Primary Health Centre', 'an Neni Comprehensive Health Centre NAUTH', 'An Rise Clinic, Adazi Ani', 'Lancet Hospital LTD', 'an God is Able Maternity Home', 'an Nibo Primary Health Center', 'an Nise Primary Health Center', 'an Okpuno Primary Health Center', 'Anambra State One Stop Shop (OSS), Awka', 'Chuwuemeka Odumegwu Ojukwu University Teaching Hospital', 'Faith Hospital and Maternity, Awka', 'Isiagu Primary Health Centre', 'Regina Caeli Hospital', 'an Maranatha Caring Mission Hospital', 'Ukpo Comprehensive Health Centre', 'an Ichi Referral Health Centre', 'an Joint Hospital Ozubulu', 'an Ozubulu Referral Health Center', 'Evans Specialist Hospital', 'Orafite General Hospital', 'an Eziowelle Primary Health Center', 'an Madueke Memorial Hospital and Maternity', 'Immaculate Heart Hospital (Nkpor)', 'Iyi-Enu Hospital', 'Nkpor Crown Hospital and Maternity', 'Obosi St Martins Hospital', 'an Obinwane Hospital and Maternity', '''an Ojoto Uno Primary Health Centre
''', 'Oba Comprehensive Health Centre (Trauma Centre)', 'an Eziani Health Centre', 'Our Lady of Lourdes Hospital', 'General Hospital Enugu-Ukwu', 'an Otolo Umuenem Primary Health Centre', 'Nnamdi Azikiwe University Teaching Hospital', 'Nnewi Diocesan Hospital', 'Amichi Diocesan Hospital', 'Osumenyi Visitation Hospital', 'Ukpor General Hospital', 'an Accucare Foundation Hospital', 'an St Felix Hospital and Maternity', 'Okpoko St Lukes Diocesan Hospital', 'Patricia Memorial Hospital and Maternity Aka baby', 'Holy Rosary Hospital', 'Onitsha General Hospital', 'St Charles Borromeo Hospital', 'Onitsha Pieta Hospital', 'Community Hospital Oko', 'an Trinitas International Hospital Umuchukwu', 'an Umunze Immaculate Heart Hospital', 'an Awkuzu Primary Health Centre', 'an Ogbunike Primary Health Centre', 'Umunya Comprehensive Health Centre (NAUTH)', 'an Ogbu Primary Health Centre', 'an St. Joseph Hospital and Maternity, Ifite Ogwari'],
    'Name on NMRS': ['', 'Ekwulobia General Hospital', 'Igboukwu DiocesHospital','','','','','','', 'Umueri DiocesHospital and Maternity','','','Oroma-Etiti Referral Health Centre', 'Adazi St Joseph Hospital','','','', 'Neni Comprehensive Health Centre NAUTH', 'Rise Clinic, Adazi Ani', 'Lancet Hospital LTD','','','','', 'Anambra State One Stop Shop (OSS) Site, Awka', 'Chuwuemeka Odumegwu Ojukwu University Teaching Hospital', 'Faith Hospital and Maternity, Awka','', 'Regina Caeli Hospital -  Awka', 'Maranatha Caring Mission Hospital', 'Ukpo Comprehensive Health Centre','','','', 'Evans Specialist Hospital', 'Orafite General Hospital','','', 'Immaculate Heart Hospital (Nkpor)', 'Iyi-Enu Hospital', 'Nkpor Crown Hospital and Maternity', "Obosi St Martin's Hospital",'', 'Ojotu Uno Primary Health Centre', 'Oba Comprehensive Health Centre (Trauma Centre)','', 'Our Lady of Lourdes Hospital', 'Enugu-Ukwu General Hospital', 'Otolo Umuenem Primary Health Centre','', 'Nnewi Diocesan Hospital', 'Amichi Diocesan Hospital', 'Osumenyi Visitation Hospital', 'Ukpor General Hospital','','', "Okpoko St Luke's Diocesan Hospital", 'Patricia Memorial Hospital and Maternity', 'Holy Rosary Hospital','','', 'Onitsha Pieta Hospital', 'Community Hospital Oko','', 'Umunze Immaculate Heart Hospital','','', 'Umunya Comprehensive Health Centre (NAUTH)','', 'St. Joseph Hospital and Maternity, Ifite Ogwari'],
    'NMRS Use Status': ['NO', 'Yes', 'Yes', 'NO', 'NO', 'NO', 'NO', 'NO', 'NO', 'Yes', 'NO', 'NO', 'Yes', 'Yes', 'NO', 'NO', 'NO', 'Yes', 'Yes', 'Yes', 'NO', 'NO', 'NO', 'NO', 'Yes', 'Yes', 'Yes', 'NO', 'Yes', 'Yes', 'Yes', 'NO', 'NO', 'NO', 'Yes', 'Yes', 'NO', 'NO', 'Yes', 'Yes', 'Yes', 'Yes', 'NO', 'Yes', 'Yes', 'NO', 'Yes', 'Yes', 'Yes', 'NO', 'Yes', 'Yes', 'Yes', 'Yes', 'NO', 'NO', 'Yes', 'Yes', 'Yes', 'NO', 'NO', 'Yes', 'Yes', 'NO', 'Yes', 'NO', 'NO', 'Yes', 'NO', 'Yes']
}
dflamisnmrs = pd.DataFrame(emrData)

def cleandflamisnmrs():
    global dflamisnmrs
    dflamisnmrs = dflamisnmrs[(dflamisnmrs != '').all(axis=1)]
    #dflamisnmrs.to_excel('dflamisnmrs.xlsx')
    return dflamisnmrs

def third95Thrimed():
        startDate, endDate = get_selected_date()
        start_time = time.time()  # Start time measurement
        status_label.config(text=f"Attempting to read file...")
        input_file_path = filedialog.askopenfilename(initialdir = '/Desktop', 
                                                        title = 'Select a excel file', 
                                                        filetypes = (('excel file','*.xls'), 
                                                                     ('excel file','*.csv'),
                                                                     ('excel file','*.xlsx')))
        if not input_file_path:  # Check if the file selection was cancelled
            status_label.config(text="No file selected. Please select a file to convert.")
            return  # Exit the function
        
        # Determine file type and read accordingly
        file_ext = os.path.splitext(input_file_path)[1].lower()

        if file_ext == '.csv':
                
            df = pd.read_csv(input_file_path, 
                   encoding='utf-8', 
                   lineterminator='\n', 
                   quotechar='"', 
                   escapechar='\\', 
                   skip_blank_lines=True)
            
            #Convert Date Objects to Date using dateutil
            dfDates = ['DateConfirmedHIV+', 'ARTStartDate', 'Pharmacy_LastPickupdate', 'Pharmacy_LastPickupdate_PreviousQuarter', 'DateofCurrentViralLoad', 'DateResultReceivedFacility', 'LastDateOfSampleCollection', 'Outcomes_Date', 'IIT_Date', 'DOB', 'Date_Transfered_In', 'DateofFirstTLD_Pickup', 'EstimatedNextAppointmentPharmacy', 'Next_Ap_by_careCard', 'IPT_Screening_Date', 'First_TPT_Pickupdate', 'Last_TPT_Pickupdate', 'Current_TPT_Received', 'Date_of_TPT_Outcome', 'DateofCurrent_TBStatus', 'TB_Treatment_Start_Date', 'TB_Treatment_Stop_Date', 'Date_Enrolled_Into_OTZ', 'Date_Enrolled_Into_OTZ_Plus', 'PBS_Capture_Date', 'Date_Generated', 'PBS_Recapture_Date']
            for col in dfDates:
                df[col] = df[col].apply(parse_date)
                
            #Clean numbers
            dfnumeric = ['AgeAtStartofART', 'AgeinMonths', 'DaysOnART', 'DaysOfARVRefill', 'CurrentViralLoad', 'Current_Age', 'Weight', 'Height', 'BMI', 'Whostage', 'CurrentCD4', 'Days_To_Schedule']
            for col in dfnumeric:
                df[col] = pd.to_numeric(df[col],errors='coerce')
            
        elif file_ext in ['.xls', '.xlsx']:
            df = pd.read_excel(input_file_path,sheet_name=0, dtype=object)
            
            #Convert Date Objects to Date using dateutil
            dfDates = ['DateConfirmedHIV+', 'ARTStartDate', 'Pharmacy_LastPickupdate', 'Pharmacy_LastPickupdate_PreviousQuarter', 'DateofCurrentViralLoad', 'DateResultReceivedFacility', 'LastDateOfSampleCollection', 'Outcomes_Date', 'IIT_Date', 'DOB', 'Date_Transfered_In', 'DateofFirstTLD_Pickup', 'EstimatedNextAppointmentPharmacy', 'Next_Ap_by_careCard', 'IPT_Screening_Date', 'First_TPT_Pickupdate', 'Last_TPT_Pickupdate', 'Current_TPT_Received', 'Date_of_TPT_Outcome', 'DateofCurrent_TBStatus', 'TB_Treatment_Start_Date', 'TB_Treatment_Stop_Date', 'Date_Enrolled_Into_OTZ', 'Date_Enrolled_Into_OTZ_Plus', 'PBS_Capture_Date', 'PBS_Recapture_Date']
            for col in dfDates:
                df[col] = df[col].apply(parse_date)
                
            #Clean numbers
            dfnumeric = ['AgeAtStartofART', 'AgeinMonths', 'DaysOnART', 'DaysOfARVRefill', 'CurrentViralLoad', 'Current_Age', 'Weight', 'Height', 'BMI', 'Whostage', 'CurrentCD4', 'Days_To_Schedule']
            for col in dfnumeric:
                df[col] = pd.to_numeric(df[col],errors='coerce')
         
        # Ensure the 'Pharmacy_LastPickupdate' column is in datetime format and fill NaNs with a specific date
        df['Pharmacy_LastPickupdate2'] = pd.to_datetime(df['Pharmacy_LastPickupdate'], errors='coerce').fillna(pd.to_datetime('1900'))

        #Fill zero if the column contains number greater than 180
        df['DaysOfARVRefill2'] = df['DaysOfARVRefill'].apply(lambda x: 0 if x > 180 else x)

        # Calculate the 'NextAppt' column by adding the 'DaysOfARVRefill2' to 'Pharmacy_LastPickupdate2'
        df['NextAppt'] = df['Pharmacy_LastPickupdate2'] + pd.to_timedelta(df['DaysOfARVRefill2'], unit='D') 
         
        dflamisnmrs = cleandflamisnmrs() 
        print(dflamisnmrs)  
        df['LGA2']=df['FacilityName'].map(dflamisnmrs.set_index('Name on NMRS')['LGA'])
        df['STATE2']=df['FacilityName'].map(dflamisnmrs.set_index('Name on NMRS')['STATE'])
        df.loc[df['LGA'].isna(), 'LGA'] = df['LGA2']
        df.loc[df['State'].isna(), 'State'] = df['STATE2']
        
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
        #df['CurrentViralLoad2'] = df['CurrentViralLoad'].fillna(0)

        #3rd 95 columns integration
        df['validVlResult'] = df.apply(lambda row: 'Valid Result' if row['DateResultReceivedFacility'] > first_quarter_last_year else 'Invalid Result', axis=1)
        df['validVlSampleCollection'] = df.apply(lambda row: 'Valid SC' if row['LastDateOfSampleCollection'] > first_quarter_last_year else 'Invalid SC', axis=1)
        df['vlSCGap'] = df.apply(lambda row: 'SC Gap' if row['validVlSampleCollection'] == 'Invalid SC' else 'Not SC Gap', axis=1)
        df['PendingResult'] = df.apply(lambda row: 'Pending' if ((row['validVlSampleCollection'] == 'Valid SC') & (row['validVlResult'] == 'Invalid Result')) else 'Not pending', axis=1)  
        df['last30daysmissedSC'] = df.apply(lambda row: 'Missed SC' if ((row['vlSCGap'] == 'SC Gap') & (row['Pharmacy_LastPickupdate'] >= last_30_days) & (row['Pharmacy_LastPickupdate'] < today)) else 'Not Missed SC', axis=1)  
        df['expNext30daysdueforSC'] = df.apply(lambda row: 'due for SC' if ((row['vlSCGap'] == 'SC Gap') & (row['NextAppt'] >= today) & (row['NextAppt'] < next_30_days)) else 'Not due for SC', axis=1)  
        #df['Suppression'] = df.apply(lambda row: 'Suppressed' if ((row['validVlResult'] == 'Valid Result') & (row['CurrentViralLoad2'] < 1000)) else 'Not Suppressed', axis=1)
        df['Suppression'] = df.apply(lambda row: 'Suppressed' if ((row['validVlResult'] == 'Valid Result') & (row['CurrentViralLoad'] < 1000)) else ('Unsuppressed' if ((row['validVlResult'] == 'Valid Result') & (row['CurrentViralLoad'] > 1000)) else 'Invalid Result'), axis=1)
        
        #normalise VL datetime back to date
        df['DateResultReceivedFacility'] = df['DateResultReceivedFacility'].apply(parse_date)
        df['LastDateOfSampleCollection'] = df['LastDateOfSampleCollection'].apply(parse_date)
        df['NextAppt'] = df['NextAppt'].apply(parse_date)
        df['Pharmacy_LastPickupdate'] = df['Pharmacy_LastPickupdate'].apply(parse_date)
        df['NextAppt'] = df['NextAppt'].apply(parse_date)
        
        df.to_excel('3rd95.xlsx')
        
        df_sc_Gap = df[df['vlSCGap'] == 'SC Gap'].reset_index()
        df_sc_Gap.insert(0, 'S/N', df_sc_Gap.index + 1)
        
        df_pending_results = df[df['PendingResult'] == 'Pending'].reset_index()
        df_pending_results.insert(0, 'S/N', df_pending_results.index + 1)
        
        df_last30daysmissedSC = df[df['last30daysmissedSC'] == 'Missed SC'].reset_index()
        df_last30daysmissedSC.insert(0, 'S/N', df_last30daysmissedSC.index + 1)
        
        df_expNext30daysdueforSC = df[df['expNext30daysdueforSC'] == 'due for SC'].reset_index()
        df_expNext30daysdueforSC.insert(0, 'S/N', df_expNext30daysdueforSC.index + 1)
        
        df_Suppression = df[df['Suppression'] == 'Unsuppressed'].reset_index()
        df_Suppression.insert(0, 'S/N', df_Suppression.index + 1)
        
        #Trime Line lists
        df_sc_Gap = df_sc_Gap[["S/N", "State", "LGA", "FacilityName", "PEPID", "PatientHospitalNo", "uuid", "Sex", "Current_Age", "Surname", "Firstname", "MaritalStatus", "PhoneNo", "Address", "State_of_Residence", "LGA_of_Residence", "DateConfirmedHIV+", "ARTStartDate", "Pharmacy_LastPickupdate", "DaysOfARVRefill", "CurrentARTRegimen", "NextAppt", "CurrentPregnancyStatus", "CurrentViralLoad", "DateResultReceivedFacility", "Alphanumeric_Viral_Load_Result", "LastDateOfSampleCollection", "Outcomes", "Outcomes_Date", "IIT_Date", "CurrentARTStatus", "First_TPT_Pickupdate", "Current_TPT_Received", "PBS_Capturee", "PBS_Capture_Date", "Date_Generated"]]
        df_pending_results = df_pending_results[["S/N", "State", "LGA", "FacilityName", "PEPID", "PatientHospitalNo", "uuid", "Sex", "Current_Age", "Surname", "Firstname", "MaritalStatus", "PhoneNo", "Address", "State_of_Residence", "LGA_of_Residence", "DateConfirmedHIV+", "ARTStartDate", "Pharmacy_LastPickupdate", "DaysOfARVRefill", "CurrentARTRegimen", "NextAppt", "CurrentPregnancyStatus", "CurrentViralLoad", "DateResultReceivedFacility", "Alphanumeric_Viral_Load_Result", "LastDateOfSampleCollection", "Outcomes", "Outcomes_Date", "IIT_Date", "CurrentARTStatus", "First_TPT_Pickupdate", "Current_TPT_Received", "PBS_Capturee", "PBS_Capture_Date", "Date_Generated"]]
        df_last30daysmissedSC = df_last30daysmissedSC[["S/N", "State", "LGA", "FacilityName", "PEPID", "PatientHospitalNo", "uuid", "Sex", "Current_Age", "Surname", "Firstname", "MaritalStatus", "PhoneNo", "Address", "State_of_Residence", "LGA_of_Residence", "DateConfirmedHIV+", "ARTStartDate", "Pharmacy_LastPickupdate", "DaysOfARVRefill", "CurrentARTRegimen", "NextAppt", "CurrentPregnancyStatus", "CurrentViralLoad", "DateResultReceivedFacility", "Alphanumeric_Viral_Load_Result", "LastDateOfSampleCollection", "Outcomes", "Outcomes_Date", "IIT_Date", "CurrentARTStatus", "First_TPT_Pickupdate", "Current_TPT_Received", "PBS_Capturee", "PBS_Capture_Date", "Date_Generated"]]
        df_expNext30daysdueforSC = df_expNext30daysdueforSC[["S/N", "State", "LGA", "FacilityName", "PEPID", "PatientHospitalNo", "uuid", "Sex", "Current_Age", "Surname", "Firstname", "MaritalStatus", "PhoneNo", "Address", "State_of_Residence", "LGA_of_Residence", "DateConfirmedHIV+", "ARTStartDate", "Pharmacy_LastPickupdate", "DaysOfARVRefill", "CurrentARTRegimen", "NextAppt", "CurrentPregnancyStatus", "CurrentViralLoad", "DateResultReceivedFacility", "Alphanumeric_Viral_Load_Result", "LastDateOfSampleCollection", "Outcomes", "Outcomes_Date", "IIT_Date", "CurrentARTStatus", "First_TPT_Pickupdate", "Current_TPT_Received", "PBS_Capturee", "PBS_Capture_Date", "Date_Generated"]]
        df_Suppression = df_Suppression[["S/N", "State", "LGA", "FacilityName", "PEPID", "PatientHospitalNo", "uuid", "Sex", "Current_Age", "Surname", "Firstname", "MaritalStatus", "PhoneNo", "Address", "State_of_Residence", "LGA_of_Residence", "DateConfirmedHIV+", "ARTStartDate", "Pharmacy_LastPickupdate", "DaysOfARVRefill", "CurrentARTRegimen", "NextAppt", "CurrentPregnancyStatus", "CurrentViralLoad", "DateResultReceivedFacility", "Alphanumeric_Viral_Load_Result", "LastDateOfSampleCollection", "Outcomes", "Outcomes_Date", "IIT_Date", "CurrentARTStatus", "First_TPT_Pickupdate", "Current_TPT_Received", "PBS_Capturee", "PBS_Capture_Date", "Date_Generated"]]
        
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
        
        

        progress_bar['maximum'] = len(df)  # Set to 100 for percentage completion
        for index, row in df.iterrows():
                
            # Update the progress bar value
            progress_bar['value'] = index + 1
            
            # Calculate the percentage of completion
            percentage = ((index + 1) / len(df)) * 100
            
            # Update the status label with the current percentage
            status_label.config(text=f"Conversion in Progress: {index + 1}/{len(df)} ({percentage:.2f}%)")
            
            # Update the GUI to reflect changes
            root.update_idletasks()
            
            # Simulate time-consuming task
            time.sleep(0.000001)
        
        #format and export
        sdate = endDate.strftime("%d-%m-%Y")

        print(sdate)
        input_file_path = f"3RD 95 ANALYSIS AS AT_{sdate}xlsx"
        output_file_name = input_file_path.split("/")[-1][:-4]
        status_label2.config(text=f"Just a moment! Formating and Saving Converted File...")
        output_file_path = filedialog.asksaveasfilename(initialdir = '/Desktop', 
                                                        title = 'Select a excel file', 
                                                        filetypes = (('excel file','*.xls'), 
                                                                     ('excel file','*.xlsx')),defaultextension=".xlsx", initialfile=output_file_name)
        if not output_file_path:  # Check if the file save was cancelled
            status_label.config(text="File conversion was cancelled. No file was saved.")
            status_label2.config(text="File Conversion Cancelled!")
            progress_bar['value'] = 0
            return  # Exit the function

        # Create a Pandas Excel writer using XlsxWriter as the engine.
        writer = pd.ExcelWriter(output_file_path, engine="xlsxwriter")

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
                
            # Add percentage format to a specific column in "3RD 95 SUMMARY" (e.g., column B which is index 1)
            if sheet_name == "3RD 95 SUMMARY":
                # Add title to the worksheet
                title_format = workbook.add_format({'bold': True, 'font_size': 14, 'align': 'center'})
                worksheet.merge_range(0, 0, 0, len(dataframe.columns) - 1, '3RD 95 SUMMARY', title_format)
                
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
            if sheet_name in ["SC GAP", "PENDING RESULT"]:
                row_band_format = workbook.add_format({'bg_color': '#F9F9F9'})
                worksheet.conditional_format(1, 0, len(dataframe), len(dataframe.columns) - 1, 
                                            {'type': 'formula', 'criteria': 'MOD(ROW(), 2) = 0', 'format': row_band_format})
                                
                # Specify column widths for a range of columns
                column_ranges = {
                    'O:AF': 10,  # Example width for columns A to C
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
        end_time = time.time()  # End time measurement
        total_time = end_time - start_time  # Calculate total time taken
        status_label.config(text=f"Conversion Complete! Time taken: {total_time:.2f} seconds")
        status_label2.config(text=f" Converted File Location: {output_file_path}")
        

def third95ThrimedCMG():
        startDate, endDate = get_selected_date()
        start_time = time.time()  # Start time measurement
        
        status_label.config(text=f"Attempting to read file...")
        input_CMG = filedialog.askopenfilename(initialdir = '/Desktop', 
                                                        title = 'Select a excel file', 
                                                        filetypes = (('excel file','*.xls'), 
                                                                     ('excel file','*.csv'),
                                                                     ('excel file','*.xlsx')))
        if not input_CMG:  # Check if the file selection was cancelled
            status_label.config(text="No file selected. Please select a file to convert.")
            return  # Exit the function
        
        # Determine file type and read accordingly
        file_ext = os.path.splitext(input_CMG)[1].lower()

        if file_ext == '.csv':
            dfCMG = pd.read_csv(input_CMG, 
                   encoding='utf-8', 
                   lineterminator='\n', 
                   quotechar='"', 
                   escapechar='\\', 
                   skip_blank_lines=True)
        
        elif file_ext in ['.xls', '.xlsx']:
            dfCMG = pd.read_excel(input_CMG,sheet_name=0, dtype=object)
        
        status_label.config(text=f"Attempting to read file...")
        input_file_path = filedialog.askopenfilename(initialdir = '/Desktop', 
                                                        title = 'Select a excel file', 
                                                        filetypes = (('excel file','*.xls'), 
                                                                     ('excel file','*.csv'),
                                                                     ('excel file','*.xlsx')))
        if not input_file_path:  # Check if the file selection was cancelled
            status_label.config(text="No file selected. Please select a file to convert.")
            return  # Exit the function
        
        # Determine file type and read accordingly
        file_ext = os.path.splitext(input_file_path)[1].lower()

        if file_ext == '.csv':
                
            df = pd.read_csv(input_file_path, 
                   encoding='utf-8', 
                   lineterminator='\n', 
                   quotechar='"', 
                   escapechar='\\', 
                   skip_blank_lines=True)
            
            #Convert Date Objects to Date using dateutil
            dfDates = ['DateConfirmedHIV+', 'ARTStartDate', 'Pharmacy_LastPickupdate', 'Pharmacy_LastPickupdate_PreviousQuarter', 'DateofCurrentViralLoad', 'DateResultReceivedFacility', 'LastDateOfSampleCollection', 'Outcomes_Date', 'IIT_Date', 'DOB', 'Date_Transfered_In', 'DateofFirstTLD_Pickup', 'EstimatedNextAppointmentPharmacy', 'Next_Ap_by_careCard', 'IPT_Screening_Date', 'First_TPT_Pickupdate', 'Last_TPT_Pickupdate', 'Current_TPT_Received', 'Date_of_TPT_Outcome', 'DateofCurrent_TBStatus', 'TB_Treatment_Start_Date', 'TB_Treatment_Stop_Date', 'Date_Enrolled_Into_OTZ', 'Date_Enrolled_Into_OTZ_Plus', 'PBS_Capture_Date', 'Date_Generated', 'PBS_Recapture_Date']
            for col in dfDates:
                df[col] = df[col].apply(parse_date)
                
            #Clean numbers
            dfnumeric = ['AgeAtStartofART', 'AgeinMonths', 'DaysOnART', 'DaysOfARVRefill', 'CurrentViralLoad', 'Current_Age', 'Weight', 'Height', 'BMI', 'Whostage', 'CurrentCD4', 'Days_To_Schedule']
            for col in dfnumeric:
                df[col] = pd.to_numeric(df[col],errors='coerce')
            
        elif file_ext in ['.xls', '.xlsx']:
            df = pd.read_excel(input_file_path,sheet_name=0, dtype=object)
            
            #Convert Date Objects to Date using dateutil
            dfDates = ['DateConfirmedHIV+', 'ARTStartDate', 'Pharmacy_LastPickupdate', 'Pharmacy_LastPickupdate_PreviousQuarter', 'DateofCurrentViralLoad', 'DateResultReceivedFacility', 'LastDateOfSampleCollection', 'Outcomes_Date', 'IIT_Date', 'DOB', 'Date_Transfered_In', 'DateofFirstTLD_Pickup', 'EstimatedNextAppointmentPharmacy', 'Next_Ap_by_careCard', 'IPT_Screening_Date', 'First_TPT_Pickupdate', 'Last_TPT_Pickupdate', 'Current_TPT_Received', 'Date_of_TPT_Outcome', 'DateofCurrent_TBStatus', 'TB_Treatment_Start_Date', 'TB_Treatment_Stop_Date', 'Date_Enrolled_Into_OTZ', 'Date_Enrolled_Into_OTZ_Plus', 'PBS_Capture_Date', 'PBS_Recapture_Date']
            for col in dfDates:
                df[col] = df[col].apply(parse_date)
                
            #Clean numbers
            dfnumeric = ['AgeAtStartofART', 'AgeinMonths', 'DaysOnART', 'DaysOfARVRefill', 'CurrentViralLoad', 'Current_Age', 'Weight', 'Height', 'BMI', 'Whostage', 'CurrentCD4', 'Days_To_Schedule']
            for col in dfnumeric:
                df[col] = pd.to_numeric(df[col],errors='coerce')
                
        #add case managers and next appt to the dataframe
        df['CaseManager']=df['uuid'].map(dfCMG.set_index('uuid')['CASE MANAGER'])
        df['CaseManager'] = df['CaseManager'].fillna('UNASSIGNED')
              
        # Ensure the 'Pharmacy_LastPickupdate' column is in datetime format and fill NaNs with a specific date
        df['Pharmacy_LastPickupdate2'] = pd.to_datetime(df['Pharmacy_LastPickupdate'], errors='coerce').fillna(pd.to_datetime('1900'))

        #Fill zero if the column contains number greater than 180
        df['DaysOfARVRefill2'] = df['DaysOfARVRefill'].apply(lambda x: 0 if x > 180 else x)

        # Calculate the 'NextAppt' column by adding the 'DaysOfARVRefill2' to 'Pharmacy_LastPickupdate2'
        df['NextAppt'] = df['Pharmacy_LastPickupdate2'] + pd.to_timedelta(df['DaysOfARVRefill2'], unit='D')     
        
        dflamisnmrs = cleandflamisnmrs() 
        print(dflamisnmrs)  
        df['LGA2']=df['FacilityName'].map(dflamisnmrs.set_index('Name on NMRS')['LGA'])
        df['STATE2']=df['FacilityName'].map(dflamisnmrs.set_index('Name on NMRS')['STATE'])
        df.loc[df['LGA'].isna(), 'LGA'] = df['LGA2']
        df.loc[df['State'].isna(), 'State'] = df['STATE2']
        
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
        df.to_excel("3rd95.xlsx")
        
        #normalise VL datetime back to date
        df['DateResultReceivedFacility'] = df['DateResultReceivedFacility'].apply(parse_date)
        df['LastDateOfSampleCollection'] = df['LastDateOfSampleCollection'].apply(parse_date)
        df['NextAppt'] = df['NextAppt'].apply(parse_date)
        df['Pharmacy_LastPickupdate'] = df['Pharmacy_LastPickupdate'].apply(parse_date)
        
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


        progress_bar['maximum'] = len(df)  # Set to 100 for percentage completion
        for index, row in df.iterrows():
                
            # Update the progress bar value
            progress_bar['value'] = index + 1
            
            # Calculate the percentage of completion
            percentage = ((index + 1) / len(df)) * 100
            
            # Update the status label with the current percentage
            status_label.config(text=f"Conversion in Progress: {index + 1}/{len(df)} ({percentage:.2f}%)")
            
            # Update the GUI to reflect changes
            root.update_idletasks()
            
            # Simulate time-consuming task
            time.sleep(0.000001)
        
        #format and export
        sdate = endDate.strftime("%d-%m-%Y")

        print(sdate)
        input_file_path = f"3RD 95 ANALYSIS BY CMG AS AT_{sdate}xlsx"
        output_file_name = input_file_path.split("/")[-1][:-4]
        status_label2.config(text=f"Just a moment! Formating and Saving Converted File...")
        output_file_path = filedialog.asksaveasfilename(initialdir = '/Desktop', 
                                                        title = 'Select a excel file', 
                                                        filetypes = (('excel file','*.xls'), 
                                                                     ('excel file','*.xlsx')),defaultextension=".xlsx", initialfile=output_file_name)
        if not output_file_path:  # Check if the file save was cancelled
            status_label.config(text="File conversion was cancelled. No file was saved.")
            status_label2.config(text="File Conversion Cancelled!")
            progress_bar['value'] = 0
            return  # Exit the function

        # Create a Pandas Excel writer using XlsxWriter as the engine.
        writer = pd.ExcelWriter(output_file_path, engine="xlsxwriter")

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
                worksheet.merge_range(0, 0, 0, len(dataframe.columns) - 1, '3RD 95 SUMMARY', title_format)
                
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
        end_time = time.time()  # End time measurement
        total_time = end_time - start_time  # Calculate total time taken
        status_label.config(text=f"Conversion Complete! Time taken: {total_time:.2f} seconds")
        status_label2.config(text=f" Converted File Location: {output_file_path}")
        

def third95():
        startDate, endDate = get_selected_date()
        start_time = time.time()  # Start time measurement
        status_label.config(text=f"Attempting to read file...")
        input_file_path = filedialog.askopenfilename(initialdir = '/Desktop', 
                                                        title = 'Select a excel file', 
                                                        filetypes = (('excel file','*.xls'), 
                                                                     ('excel file','*.csv'),
                                                                     ('excel file','*.xlsx')))
        if not input_file_path:  # Check if the file selection was cancelled
            status_label.config(text="No file selected. Please select a file to convert.")
            return  # Exit the function
        
        # Determine file type and read accordingly
        file_ext = os.path.splitext(input_file_path)[1].lower()

        if file_ext == '.csv':
                
            df = pd.read_csv(input_file_path, 
                   encoding='utf-8', 
                   lineterminator='\n', 
                   quotechar='"', 
                   escapechar='\\', 
                   skip_blank_lines=True)
            
            #Convert Date Objects to Date using dateutil
            dfDates = ['DateConfirmedHIV+', 'ARTStartDate', 'Pharmacy_LastPickupdate', 'Pharmacy_LastPickupdate_PreviousQuarter', 'DateofCurrentViralLoad', 'DateResultReceivedFacility', 'LastDateOfSampleCollection', 'Outcomes_Date', 'IIT_Date', 'DOB', 'Date_Transfered_In', 'DateofFirstTLD_Pickup', 'EstimatedNextAppointmentPharmacy', 'Next_Ap_by_careCard', 'IPT_Screening_Date', 'First_TPT_Pickupdate', 'Last_TPT_Pickupdate', 'Current_TPT_Received', 'Date_of_TPT_Outcome', 'DateofCurrent_TBStatus', 'TB_Treatment_Start_Date', 'TB_Treatment_Stop_Date', 'Date_Enrolled_Into_OTZ', 'Date_Enrolled_Into_OTZ_Plus', 'PBS_Capture_Date', 'Date_Generated', 'PBS_Recapture_Date']
            for col in dfDates:
                df[col] = df[col].apply(parse_date)
                
            #Clean numbers
            dfnumeric = ['AgeAtStartofART', 'AgeinMonths', 'DaysOnART', 'DaysOfARVRefill', 'CurrentViralLoad', 'Current_Age', 'Weight', 'Height', 'BMI', 'Whostage', 'CurrentCD4', 'Days_To_Schedule']
            for col in dfnumeric:
                df[col] = pd.to_numeric(df[col],errors='coerce')
            
        elif file_ext in ['.xls', '.xlsx']:
            df = pd.read_excel(input_file_path,sheet_name=0, dtype=object)
            
            #Convert Date Objects to Date using dateutil
            dfDates = ['DateConfirmedHIV+', 'ARTStartDate', 'Pharmacy_LastPickupdate', 'Pharmacy_LastPickupdate_PreviousQuarter', 'DateofCurrentViralLoad', 'DateResultReceivedFacility', 'LastDateOfSampleCollection', 'Outcomes_Date', 'IIT_Date', 'DOB', 'Date_Transfered_In', 'DateofFirstTLD_Pickup', 'EstimatedNextAppointmentPharmacy', 'Next_Ap_by_careCard', 'IPT_Screening_Date', 'First_TPT_Pickupdate', 'Last_TPT_Pickupdate', 'Current_TPT_Received', 'Date_of_TPT_Outcome', 'DateofCurrent_TBStatus', 'TB_Treatment_Start_Date', 'TB_Treatment_Stop_Date', 'Date_Enrolled_Into_OTZ', 'Date_Enrolled_Into_OTZ_Plus', 'PBS_Capture_Date', 'PBS_Recapture_Date']
            for col in dfDates:
                df[col] = df[col].apply(parse_date)
                
            #Clean numbers
            dfnumeric = ['AgeAtStartofART', 'AgeinMonths', 'DaysOnART', 'DaysOfARVRefill', 'CurrentViralLoad', 'Current_Age', 'Weight', 'Height', 'BMI', 'Whostage', 'CurrentCD4', 'Days_To_Schedule']
            for col in dfnumeric:
                df[col] = pd.to_numeric(df[col],errors='coerce')
        
        # Ensure the 'Pharmacy_LastPickupdate' column is in datetime format and fill NaNs with a specific date
        df['Pharmacy_LastPickupdate2'] = pd.to_datetime(df['Pharmacy_LastPickupdate'], errors='coerce').fillna(pd.to_datetime('1900'))

        #Fill zero if the column contains number greater than 180
        df['DaysOfARVRefill2'] = df['DaysOfARVRefill'].apply(lambda x: 0 if x > 180 else x)

        # Calculate the 'NextAppt' column by adding the 'DaysOfARVRefill2' to 'Pharmacy_LastPickupdate2'
        df['NextAppt'] = df['Pharmacy_LastPickupdate2'] + pd.to_timedelta(df['DaysOfARVRefill2'], unit='D') 
         
        dflamisnmrs = cleandflamisnmrs() 
        print(dflamisnmrs)  
        df['LGA2']=df['FacilityName'].map(dflamisnmrs.set_index('Name on NMRS')['LGA'])
        df['STATE2']=df['FacilityName'].map(dflamisnmrs.set_index('Name on NMRS')['STATE'])
        df.loc[df['LGA'].isna(), 'LGA'] = df['LGA2']
        df.loc[df['State'].isna(), 'State'] = df['STATE2']
        
        df = df.drop(['LGA2','STATE2'], axis=1)
            
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
        
        #normalise VL datetime back to date
        df['DateResultReceivedFacility'] = df['DateResultReceivedFacility'].apply(parse_date)
        df['LastDateOfSampleCollection'] = df['LastDateOfSampleCollection'].apply(parse_date)
        df['NextAppt'] = df['NextAppt'].apply(parse_date)
        df['Pharmacy_LastPickupdate'] = df['Pharmacy_LastPickupdate'].apply(parse_date)
        
        df_sc_Gap = df[df['vlSCGap'] == 'SC Gap'].reset_index()
        df_sc_Gap.insert(0, 'S/N', df_sc_Gap.index + 1)
        
        df_pending_results = df[df['PendingResult'] == 'Pending'].reset_index()
        df_pending_results.insert(0, 'S/N', df_pending_results.index + 1)
        
        df_last30daysmissedSC = df[df['last30daysmissedSC'] == 'Missed SC'].reset_index()
        df_last30daysmissedSC.insert(0, 'S/N', df_last30daysmissedSC.index + 1)
        
        df_expNext30daysdueforSC = df[df['expNext30daysdueforSC'] == 'due for SC'].reset_index()
        df_expNext30daysdueforSC.insert(0, 'S/N', df_expNext30daysdueforSC.index + 1)
        
        df_Suppression = df[df['Suppression'] == 'Unsuppressed'].reset_index()
        df_Suppression.insert(0, 'S/N', df_Suppression.index + 1)
        
        
        df_active_Eligible = df[df['CurrentARTStatus']=="Active"]

        df_active_Eligible['vlEligible'] = df_active_Eligible['CurrentARTStatus'].apply(lambda x: 1 if x == "Active" else 0)
        df_active_Eligible['validVlResult_valid'] = df_active_Eligible['validVlResult'].apply(lambda x: 1 if x == "Valid Result" else 0)
        df_active_Eligible['validVlSampleCollection'] = df_active_Eligible['validVlSampleCollection'].apply(lambda x: 1 if x == "Valid SC" else 0)
        df_active_Eligible['vlSCGap'] = df_active_Eligible['vlSCGap'].apply(lambda x: 1 if x == "SC Gap" else 0)
        df_active_Eligible['PendingResult'] = df_active_Eligible['PendingResult'].apply(lambda x: 1 if x == "Pending" else 0)
        df_active_Eligible['last30daysmissedSC'] = df_active_Eligible['last30daysmissedSC'].apply(lambda x: 1 if x == "Missed SC" else 0)
        df_active_Eligible['expNext30daysdueforSC'] = df_active_Eligible['expNext30daysdueforSC'].apply(lambda x: 1 if x == "due for SC" else 0)
        df_active_Eligible['Suppressed'] = df_active_Eligible['Suppression'].apply(lambda x: 1 if x == "Suppressed" else 0)

        #result = df_active_Eligible.groupby(['LGA','FacilityName'])[['vlEligible', 'validVlResult_valid', 'validVlSampleCollection','vlSCGap', 'PendingResult']].sum().reset_index()
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

        progress_bar['maximum'] = len(df)  # Set to 100 for percentage completion
        for index, row in df.iterrows():
                
            # Update the progress bar value
            progress_bar['value'] = index + 1
            
            # Calculate the percentage of completion
            percentage = ((index + 1) / len(df)) * 100
            
            # Update the status label with the current percentage
            status_label.config(text=f"Conversion in Progress: {index + 1}/{len(df)} ({percentage:.2f}%)")
            
            # Update the GUI to reflect changes
            root.update_idletasks()
            
            # Simulate time-consuming task
            time.sleep(0.000001)
        
        #format and export
        sdate = endDate.strftime("%d-%m-%Y")

        print(sdate)
        input_file_path = f"3RD 95 ANALYSIS AS AT_{sdate}xlsx"
        output_file_name = input_file_path.split("/")[-1][:-4]
        status_label2.config(text=f"Just a moment! Formating and Saving Converted File...")
        output_file_path = filedialog.asksaveasfilename(initialdir = '/Desktop', 
                                                        title = 'Select a excel file', 
                                                        filetypes = (('excel file','*.xls'), 
                                                                     ('excel file','*.xlsx')),defaultextension=".xlsx", initialfile=output_file_name)
        if not output_file_path:  # Check if the file save was cancelled
            status_label.config(text="File conversion was cancelled. No file was saved.")
            status_label2.config(text="File Conversion Cancelled!")
            progress_bar['value'] = 0
            return  # Exit the function

        # Create a Pandas Excel writer using XlsxWriter as the engine.
        writer = pd.ExcelWriter(output_file_path, engine="xlsxwriter")

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
                           
            # Add percentage format to a specific column in "3RD 95 SUMMARY" (e.g., column B which is index 1)
            if sheet_name == "3RD 95 SUMMARY":
                # Add title to the worksheet
                title_format = workbook.add_format({'bold': True, 'font_size': 14, 'align': 'center'})
                worksheet.merge_range(0, 0, 0, len(dataframe.columns) - 1, '3RD 95 SUMMARY', title_format)
                
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
            if sheet_name in ["SC GAP", "PENDING RESULT"]:
                row_band_format = workbook.add_format({'bg_color': '#F9F9F9'})
                worksheet.conditional_format(1, 0, len(dataframe), len(dataframe.columns) - 1, 
                                            {'type': 'formula', 'criteria': 'MOD(ROW(), 2) = 0', 'format': row_band_format})
                                
                # Specify column widths for a range of columns
                column_ranges = {
                    'G:DO': 12,  # Example width for columns A to C
                    #'J:M': 12,  # Example width for columns D to E
                    #'N:N': 40,  # Example width for columns G to H
                    #'E:F': 15,  # Example width for columns G to H
                    # Add more column ranges and their widths as needed
                }
                for col_range, width in column_ranges.items():
                    worksheet.set_column(col_range, width)
                
        # Close the Pandas Excel writer and output the Excel file.
        writer.close()
        end_time = time.time()  # End time measurement
        total_time = end_time - start_time  # Calculate total time taken
        status_label.config(text=f"Conversion Complete! Time taken: {total_time:.2f} seconds")
        status_label2.config(text=f" Converted File Location: {output_file_path}")

##

def Second95Thrimed():
        startDate, endDate = get_selected_date()
        start_time = time.time()  # Start time measurement
        status_label.config(text=f"Attempting to read file...")
        input_file_path = filedialog.askopenfilename(initialdir = '/Desktop', 
                                                        title = 'Select a excel file', 
                                                        filetypes = (('excel file','*.xls'), 
                                                                     ('excel file','*.csv'),
                                                                     ('excel file','*.xlsx')))
        if not input_file_path:  # Check if the file selection was cancelled
            status_label.config(text="No file selected. Please select a file to convert.")
            return  # Exit the function
        
        # Determine file type and read accordingly
        file_ext = os.path.splitext(input_file_path)[1].lower()

        if file_ext == '.csv':
                
            df = pd.read_csv(input_file_path, 
                   encoding='utf-8', 
                   lineterminator='\n', 
                   quotechar='"', 
                   escapechar='\\', 
                   skip_blank_lines=True)
            
            #Convert Date Objects to Date using dateutil
            dfDates = ['DateConfirmedHIV+', 'ARTStartDate', 'Pharmacy_LastPickupdate', 'Pharmacy_LastPickupdate_PreviousQuarter', 'DateofCurrentViralLoad', 'DateResultReceivedFacility', 'LastDateOfSampleCollection', 'Outcomes_Date', 'IIT_Date', 'DOB', 'Date_Transfered_In', 'DateofFirstTLD_Pickup', 'EstimatedNextAppointmentPharmacy', 'Next_Ap_by_careCard', 'IPT_Screening_Date', 'First_TPT_Pickupdate', 'Last_TPT_Pickupdate', 'Current_TPT_Received', 'Date_of_TPT_Outcome', 'DateofCurrent_TBStatus', 'TB_Treatment_Start_Date', 'TB_Treatment_Stop_Date', 'Date_Enrolled_Into_OTZ', 'Date_Enrolled_Into_OTZ_Plus', 'PBS_Capture_Date', 'Date_Generated', 'PBS_Recapture_Date']
            for col in dfDates:
                df[col] = df[col].apply(parse_date)
                
            #Clean numbers
            dfnumeric = ['AgeAtStartofART', 'AgeinMonths', 'DaysOnART', 'DaysOfARVRefill', 'CurrentViralLoad', 'Current_Age', 'Weight', 'Height', 'BMI', 'Whostage', 'CurrentCD4', 'Days_To_Schedule']
            for col in dfnumeric:
                df[col] = pd.to_numeric(df[col],errors='coerce')
            
        elif file_ext in ['.xls', '.xlsx']:
            df = pd.read_excel(input_file_path,sheet_name=0, dtype=object)
            
            #Convert Date Objects to Date using dateutil
            dfDates = ['DateConfirmedHIV+', 'ARTStartDate', 'Pharmacy_LastPickupdate', 'Pharmacy_LastPickupdate_PreviousQuarter', 'DateofCurrentViralLoad', 'DateResultReceivedFacility', 'LastDateOfSampleCollection', 'Outcomes_Date', 'IIT_Date', 'DOB', 'Date_Transfered_In', 'DateofFirstTLD_Pickup', 'EstimatedNextAppointmentPharmacy', 'Next_Ap_by_careCard', 'IPT_Screening_Date', 'First_TPT_Pickupdate', 'Last_TPT_Pickupdate', 'Current_TPT_Received', 'Date_of_TPT_Outcome', 'DateofCurrent_TBStatus', 'TB_Treatment_Start_Date', 'TB_Treatment_Stop_Date', 'Date_Enrolled_Into_OTZ', 'Date_Enrolled_Into_OTZ_Plus', 'PBS_Capture_Date', 'PBS_Recapture_Date']
            for col in dfDates:
                df[col] = df[col].apply(parse_date)
                
            #Clean numbers
            dfnumeric = ['AgeAtStartofART', 'AgeinMonths', 'DaysOnART', 'DaysOfARVRefill', 'CurrentViralLoad', 'Current_Age', 'Weight', 'Height', 'BMI', 'Whostage', 'CurrentCD4', 'Days_To_Schedule']
            for col in dfnumeric:
                df[col] = pd.to_numeric(df[col],errors='coerce')
         
        # Ensure the 'Pharmacy_LastPickupdate' column is in datetime format and fill NaNs with a specific date
        df['Pharmacy_LastPickupdate2'] = pd.to_datetime(df['Pharmacy_LastPickupdate'], errors='coerce').fillna(pd.to_datetime('1900'))

        #Fill zero if the column contains number greater than 180
        df['DaysOfARVRefill2'] = df['DaysOfARVRefill'].apply(lambda x: 0 if x > 180 else x)

        # Calculate the 'NextAppt' column by adding the 'DaysOfARVRefill2' to 'Pharmacy_LastPickupdate2'
        df['NextAppt'] = df['Pharmacy_LastPickupdate2'] + pd.to_timedelta(df['DaysOfARVRefill2'], unit='D') 
        df['IITDate2'] = (df['NextAppt'] + pd.Timedelta(days=29)).fillna('1900')
         
        dflamisnmrs = cleandflamisnmrs() 
        print(dflamisnmrs)  
        df['LGA2']=df['FacilityName'].map(dflamisnmrs.set_index('Name on NMRS')['LGA'])
        df['STATE2']=df['FacilityName'].map(dflamisnmrs.set_index('Name on NMRS')['STATE'])
        df.loc[df['LGA'].isna(), 'LGA'] = df['LGA2']
        df.loc[df['State'].isna(), 'State'] = df['STATE2']
        
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
        
        #normalise VL datetime back to date
        df['NextAppt'] = df['NextAppt'].apply(parse_date)
        
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

        progress_bar['maximum'] = len(df)  # Set to 100 for percentage completion
        for index, row in df.iterrows():
                
            # Update the progress bar value
            progress_bar['value'] = index + 1
            
            # Calculate the percentage of completion
            percentage = ((index + 1) / len(df)) * 100
            
            # Update the status label with the current percentage
            status_label.config(text=f"Conversion in Progress: {index + 1}/{len(df)} ({percentage:.2f}%)")
            
            # Update the GUI to reflect changes
            root.update_idletasks()
            
            # Simulate time-consuming task
            time.sleep(0.000001)
        
        #format and export
        sdate = endDate.strftime("%d-%m-%Y")

        print(sdate)
        input_file_path = f"2ND 95 ANALYSIS AS AT_{sdate}xlsx"
        output_file_name = input_file_path.split("/")[-1][:-4]
        status_label2.config(text=f"Just a moment! Formating and Saving Converted File...")
        output_file_path = filedialog.asksaveasfilename(initialdir = '/Desktop', 
                                                        title = 'Select a excel file', 
                                                        filetypes = (('excel file','*.xls'), 
                                                                     ('excel file','*.xlsx')),defaultextension=".xlsx", initialfile=output_file_name)
        if not output_file_path:  # Check if the file save was cancelled
            status_label.config(text="File conversion was cancelled. No file was saved.")
            status_label2.config(text="File Conversion Cancelled!")
            progress_bar['value'] = 0
            return  # Exit the function

        # Create a Pandas Excel writer using XlsxWriter as the engine.
        writer = pd.ExcelWriter(output_file_path, engine="xlsxwriter")

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

        # Close the Pandas Excel writer and output the Excel file.
        writer.close()
        end_time = time.time()  # End time measurement
        total_time = end_time - start_time  # Calculate total time taken
        status_label.config(text=f"Conversion Complete! Time taken: {total_time:.2f} seconds")
        status_label2.config(text=f"Converted File Location: {output_file_path}")


def Second95ThrimedRefillRate():
        startDate, endDate = get_selected_date()
        start_time = time.time()  # Start time measurement
        
        status_label.config(text=f"Attempting to read baseline ART line list...")
        input_file_path2 = filedialog.askopenfilename(initialdir = '/Desktop', 
                                                        title = 'Select a excel file', 
                                                        filetypes = (('excel file','*.xls'), 
                                                                     ('excel file','*.csv'),
                                                                     ('excel file','*.xlsx')))
        if not input_file_path2:  # Check if the file selection was cancelled
            status_label.config(text="No file selected. Please select a file to convert.")
            return  # Exit the function
        
        # Determine file type and read accordingly
        file_ext2 = os.path.splitext(input_file_path2)[1].lower()

        if file_ext2 == '.csv':
                
            dfbaseline = pd.read_csv(input_file_path, 
                        encoding='utf-8', 
                        lineterminator='\n', 
                        quotechar='"', 
                        escapechar='\\', 
                        skip_blank_lines=True)
            
            #Convert Date Objects to Date using dateutil
            dfDates = ['DateConfirmedHIV+', 'ARTStartDate', 'Pharmacy_LastPickupdate', 'Pharmacy_LastPickupdate_PreviousQuarter', 'DateofCurrentViralLoad', 'DateResultReceivedFacility', 'LastDateOfSampleCollection', 'Outcomes_Date', 'IIT_Date', 'DOB', 'Date_Transfered_In', 'DateofFirstTLD_Pickup', 'EstimatedNextAppointmentPharmacy', 'Next_Ap_by_careCard', 'IPT_Screening_Date', 'First_TPT_Pickupdate', 'Last_TPT_Pickupdate', 'Current_TPT_Received', 'Date_of_TPT_Outcome', 'DateofCurrent_TBStatus', 'TB_Treatment_Start_Date', 'TB_Treatment_Stop_Date', 'Date_Enrolled_Into_OTZ', 'Date_Enrolled_Into_OTZ_Plus', 'PBS_Capture_Date', 'Date_Generated', 'PBS_Recapture_Date']
            for col in dfDates:
                dfbaseline[col] = dfbaseline[col].apply(parse_date)
                
            #Clean numbers
            dfnumeric = ['AgeAtStartofART', 'AgeinMonths', 'DaysOnART', 'DaysOfARVRefill', 'CurrentViralLoad', 'Current_Age', 'Weight', 'Height', 'BMI', 'Whostage', 'CurrentCD4', 'Days_To_Schedule']
            for col in dfnumeric:
                dfbaseline[col] = pd.to_numeric(dfbaseline[col],errors='coerce')
            
        elif file_ext2 in ['.xls', '.xlsx']:
            dfbaseline = pd.read_excel(input_file_path2,sheet_name=0, dtype=object)
            
            #Convert Date Objects to Date using dateutil
            dfDates = ['DateConfirmedHIV+', 'ARTStartDate', 'Pharmacy_LastPickupdate', 'Pharmacy_LastPickupdate_PreviousQuarter', 'DateofCurrentViralLoad', 'DateResultReceivedFacility', 'LastDateOfSampleCollection', 'Outcomes_Date', 'IIT_Date', 'DOB', 'Date_Transfered_In', 'DateofFirstTLD_Pickup', 'EstimatedNextAppointmentPharmacy', 'Next_Ap_by_careCard', 'IPT_Screening_Date', 'First_TPT_Pickupdate', 'Last_TPT_Pickupdate', 'Current_TPT_Received', 'Date_of_TPT_Outcome', 'DateofCurrent_TBStatus', 'TB_Treatment_Start_Date', 'TB_Treatment_Stop_Date', 'Date_Enrolled_Into_OTZ', 'Date_Enrolled_Into_OTZ_Plus', 'PBS_Capture_Date', 'PBS_Recapture_Date']
            for col in dfDates:
                dfbaseline[col] = dfbaseline[col].apply(parse_date)
                
            #Clean numbers
            dfnumeric = ['AgeAtStartofART', 'AgeinMonths', 'DaysOnART', 'DaysOfARVRefill', 'CurrentViralLoad', 'Current_Age', 'Weight', 'Height', 'BMI', 'Whostage', 'CurrentCD4', 'Days_To_Schedule']
            for col in dfnumeric:
                dfbaseline[col] = pd.to_numeric(dfbaseline[col],errors='coerce')
        print(dfbaseline)

        
        status_label.config(text=f"Attempting to read file...")
        input_file_path = filedialog.askopenfilename(initialdir = '/Desktop', 
                                                        title = 'Select a excel file', 
                                                        filetypes = (('excel file','*.xls'), 
                                                                     ('excel file','*.csv'),
                                                                     ('excel file','*.xlsx')))
        if not input_file_path:  # Check if the file selection was cancelled
            status_label.config(text="No file selected. Please select a file to convert.")
            return  # Exit the function
        
        # Determine file type and read accordingly
        file_ext = os.path.splitext(input_file_path)[1].lower()

        if file_ext == '.csv':
                
            df = pd.read_csv(input_file_path, 
                   encoding='utf-8', 
                   lineterminator='\n', 
                   quotechar='"', 
                   escapechar='\\', 
                   skip_blank_lines=True)
            
            #Convert Date Objects to Date using dateutil
            dfDates = ['DateConfirmedHIV+', 'ARTStartDate', 'Pharmacy_LastPickupdate', 'Pharmacy_LastPickupdate_PreviousQuarter', 'DateofCurrentViralLoad', 'DateResultReceivedFacility', 'LastDateOfSampleCollection', 'Outcomes_Date', 'IIT_Date', 'DOB', 'Date_Transfered_In', 'DateofFirstTLD_Pickup', 'EstimatedNextAppointmentPharmacy', 'Next_Ap_by_careCard', 'IPT_Screening_Date', 'First_TPT_Pickupdate', 'Last_TPT_Pickupdate', 'Current_TPT_Received', 'Date_of_TPT_Outcome', 'DateofCurrent_TBStatus', 'TB_Treatment_Start_Date', 'TB_Treatment_Stop_Date', 'Date_Enrolled_Into_OTZ', 'Date_Enrolled_Into_OTZ_Plus', 'PBS_Capture_Date', 'Date_Generated', 'PBS_Recapture_Date']
            for col in dfDates:
                df[col] = df[col].apply(parse_date)
                
            #Clean numbers
            dfnumeric = ['AgeAtStartofART', 'AgeinMonths', 'DaysOnART', 'DaysOfARVRefill', 'CurrentViralLoad', 'Current_Age', 'Weight', 'Height', 'BMI', 'Whostage', 'CurrentCD4', 'Days_To_Schedule']
            for col in dfnumeric:
                df[col] = pd.to_numeric(df[col],errors='coerce')
            
        elif file_ext in ['.xls', '.xlsx']:
            df = pd.read_excel(input_file_path,sheet_name=0, dtype=object)
            
            #Convert Date Objects to Date using dateutil
            dfDates = ['DateConfirmedHIV+', 'ARTStartDate', 'Pharmacy_LastPickupdate', 'Pharmacy_LastPickupdate_PreviousQuarter', 'DateofCurrentViralLoad', 'DateResultReceivedFacility', 'LastDateOfSampleCollection', 'Outcomes_Date', 'IIT_Date', 'DOB', 'Date_Transfered_In', 'DateofFirstTLD_Pickup', 'EstimatedNextAppointmentPharmacy', 'Next_Ap_by_careCard', 'IPT_Screening_Date', 'First_TPT_Pickupdate', 'Last_TPT_Pickupdate', 'Current_TPT_Received', 'Date_of_TPT_Outcome', 'DateofCurrent_TBStatus', 'TB_Treatment_Start_Date', 'TB_Treatment_Stop_Date', 'Date_Enrolled_Into_OTZ', 'Date_Enrolled_Into_OTZ_Plus', 'PBS_Capture_Date', 'PBS_Recapture_Date']
            for col in dfDates:
                df[col] = df[col].apply(parse_date)
                
            #Clean numbers
            dfnumeric = ['AgeAtStartofART', 'AgeinMonths', 'DaysOnART', 'DaysOfARVRefill', 'CurrentViralLoad', 'Current_Age', 'Weight', 'Height', 'BMI', 'Whostage', 'CurrentCD4', 'Days_To_Schedule']
            for col in dfnumeric:
                df[col] = pd.to_numeric(df[col],errors='coerce')
                
        
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
        df['BaselineDaysOfARVRefill'] = df['BaselineDaysOfARVRefill'].apply(lambda x: 0 if x > 180 else x)
        df['BaselineNextAppt'] = df['BaselinePharmacy_LastPickupdate'] + pd.to_timedelta(df['BaselineDaysOfARVRefill'], unit='D') 
        
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
         
        dflamisnmrs = cleandflamisnmrs() 
        print(dflamisnmrs)  
        df['LGA2']=df['FacilityName'].map(dflamisnmrs.set_index('Name on NMRS')['LGA'])
        df['STATE2']=df['FacilityName'].map(dflamisnmrs.set_index('Name on NMRS')['STATE'])
        df.loc[df['LGA'].isna(), 'LGA'] = df['LGA2']
        df.loc[df['State'].isna(), 'State'] = df['STATE2']
        
        df = df.drop(['LGA2','STATE2'], axis=1)
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
        
        #normalise VL datetime back to date
        df['NextAppt'] = df['NextAppt'].apply(parse_date)
        
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

        progress_bar['maximum'] = len(df)  # Set to 100 for percentage completion
        for index, row in df.iterrows():
                
            # Update the progress bar value
            progress_bar['value'] = index + 1
            
            # Calculate the percentage of completion
            percentage = ((index + 1) / len(df)) * 100
            
            # Update the status label with the current percentage
            status_label.config(text=f"Conversion in Progress: {index + 1}/{len(df)} ({percentage:.2f}%)")
            
            # Update the GUI to reflect changes
            root.update_idletasks()
            
            # Simulate time-consuming task
            time.sleep(0.000001)
        
        #format and export
        sdate = endDate.strftime("%d-%m-%Y")

        print(sdate)
        input_file_path = f"2ND 95 ANALYSIS AS AT_{sdate}xlsx"
        output_file_name = input_file_path.split("/")[-1][:-4]
        status_label2.config(text=f"Just a moment! Formating and Saving Converted File...")
        output_file_path = filedialog.asksaveasfilename(initialdir = '/Desktop', 
                                                        title = 'Select a excel file', 
                                                        filetypes = (('excel file','*.xls'), 
                                                                     ('excel file','*.xlsx')),defaultextension=".xlsx", initialfile=output_file_name)
        if not output_file_path:  # Check if the file save was cancelled
            status_label.config(text="File conversion was cancelled. No file was saved.")
            status_label2.config(text="File Conversion Cancelled!")
            progress_bar['value'] = 0
            return  # Exit the function

        # Create a Pandas Excel writer using XlsxWriter as the engine.
        writer = pd.ExcelWriter(output_file_path, engine="xlsxwriter")

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
                
                
                #header_format.set_align('center')
                #total_row = len(dataframe) + 2
                #worksheet.merge_range(total_row, 0, total_row, 1, 'Total', header_format)
                #for col_num in range(1, len(dataframe.columns)):
                    #col_letter = chr(65 + col_num)  # Convert column number to letter
                    #worksheet.write_formula(total_row, col_num, f'SUM({col_letter}2:{col_letter}{total_row})', header_format)
            
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
        end_time = time.time()  # End time measurement
        total_time = end_time - start_time  # Calculate total time taken
        status_label.config(text=f"Conversion Complete! Time taken: {total_time:.2f} seconds")
        status_label2.config(text=f"Converted File Location: {output_file_path}")


def Second95ThrimedCMG():
        startDate, endDate = get_selected_date()
        start_time = time.time()  # Start time measurement
        
        status_label.config(text=f"Attempting to read file...")
        input_CMG = filedialog.askopenfilename(initialdir = '/Desktop', 
                                                        title = 'Select a excel file', 
                                                        filetypes = (('excel file','*.xls'), 
                                                                     ('excel file','*.csv'),
                                                                     ('excel file','*.xlsx')))
        if not input_CMG:  # Check if the file selection was cancelled
            status_label.config(text="No file selected. Please select a file to convert.")
            return  # Exit the function
        
        # Determine file type and read accordingly
        file_ext = os.path.splitext(input_CMG)[1].lower()

        if file_ext == '.csv':
            dfCMG = pd.read_csv(input_CMG, 
                   encoding='utf-8', 
                   lineterminator='\n', 
                   quotechar='"', 
                   escapechar='\\', 
                   skip_blank_lines=True)
        
        elif file_ext in ['.xls', '.xlsx']:
            dfCMG = pd.read_excel(input_CMG,sheet_name=0, dtype=object)
        
        status_label.config(text=f"Attempting to read file...")
        input_file_path = filedialog.askopenfilename(initialdir = '/Desktop', 
                                                        title = 'Select a excel file', 
                                                        filetypes = (('excel file','*.xls'), 
                                                                     ('excel file','*.csv'),
                                                                     ('excel file','*.xlsx')))
        if not input_file_path:  # Check if the file selection was cancelled
            status_label.config(text="No file selected. Please select a file to convert.")
            return  # Exit the function
        
        # Determine file type and read accordingly
        file_ext = os.path.splitext(input_file_path)[1].lower()

        if file_ext == '.csv':
                
            df = pd.read_csv(input_file_path, 
                   encoding='utf-8', 
                   lineterminator='\n', 
                   quotechar='"', 
                   escapechar='\\', 
                   skip_blank_lines=True)
            
            #Convert Date Objects to Date using dateutil
            dfDates = ['DateConfirmedHIV+', 'ARTStartDate', 'Pharmacy_LastPickupdate', 'Pharmacy_LastPickupdate_PreviousQuarter', 'DateofCurrentViralLoad', 'DateResultReceivedFacility', 'LastDateOfSampleCollection', 'Outcomes_Date', 'IIT_Date', 'DOB', 'Date_Transfered_In', 'DateofFirstTLD_Pickup', 'EstimatedNextAppointmentPharmacy', 'Next_Ap_by_careCard', 'IPT_Screening_Date', 'First_TPT_Pickupdate', 'Last_TPT_Pickupdate', 'Current_TPT_Received', 'Date_of_TPT_Outcome', 'DateofCurrent_TBStatus', 'TB_Treatment_Start_Date', 'TB_Treatment_Stop_Date', 'Date_Enrolled_Into_OTZ', 'Date_Enrolled_Into_OTZ_Plus', 'PBS_Capture_Date', 'Date_Generated', 'PBS_Recapture_Date']
            for col in dfDates:
                df[col] = df[col].apply(parse_date)
                
            #Clean numbers
            dfnumeric = ['AgeAtStartofART', 'AgeinMonths', 'DaysOnART', 'DaysOfARVRefill', 'CurrentViralLoad', 'Current_Age', 'Weight', 'Height', 'BMI', 'Whostage', 'CurrentCD4', 'Days_To_Schedule']
            for col in dfnumeric:
                df[col] = pd.to_numeric(df[col],errors='coerce')
            
        elif file_ext in ['.xls', '.xlsx']:
            df = pd.read_excel(input_file_path,sheet_name=0, dtype=object)
            
            #Convert Date Objects to Date using dateutil
            dfDates = ['DateConfirmedHIV+', 'ARTStartDate', 'Pharmacy_LastPickupdate', 'Pharmacy_LastPickupdate_PreviousQuarter', 'DateofCurrentViralLoad', 'DateResultReceivedFacility', 'LastDateOfSampleCollection', 'Outcomes_Date', 'IIT_Date', 'DOB', 'Date_Transfered_In', 'DateofFirstTLD_Pickup', 'EstimatedNextAppointmentPharmacy', 'Next_Ap_by_careCard', 'IPT_Screening_Date', 'First_TPT_Pickupdate', 'Last_TPT_Pickupdate', 'Current_TPT_Received', 'Date_of_TPT_Outcome', 'DateofCurrent_TBStatus', 'TB_Treatment_Start_Date', 'TB_Treatment_Stop_Date', 'Date_Enrolled_Into_OTZ', 'Date_Enrolled_Into_OTZ_Plus', 'PBS_Capture_Date', 'PBS_Recapture_Date']
            for col in dfDates:
                df[col] = df[col].apply(parse_date)
                
            #Clean numbers
            dfnumeric = ['AgeAtStartofART', 'AgeinMonths', 'DaysOnART', 'DaysOfARVRefill', 'CurrentViralLoad', 'Current_Age', 'Weight', 'Height', 'BMI', 'Whostage', 'CurrentCD4', 'Days_To_Schedule']
            for col in dfnumeric:
                df[col] = pd.to_numeric(df[col],errors='coerce')
                
        #add case managers and next appt to the dataframe
        df['CaseManager']=df['uuid'].map(dfCMG.set_index('uuid')['CASE MANAGER'])
        df['CaseManager'] = df['CaseManager'].fillna('UNASSIGNED')
         
        # Ensure the 'Pharmacy_LastPickupdate' column is in datetime format and fill NaNs with a specific date
        df['Pharmacy_LastPickupdate2'] = pd.to_datetime(df['Pharmacy_LastPickupdate'], errors='coerce').fillna(pd.to_datetime('1900'))

        #Fill zero if the column contains number greater than 180
        df['DaysOfARVRefill2'] = df['DaysOfARVRefill'].apply(lambda x: 0 if x > 180 else x)

        # Calculate the 'NextAppt' column by adding the 'DaysOfARVRefill2' to 'Pharmacy_LastPickupdate2'
        df['NextAppt'] = df['Pharmacy_LastPickupdate2'] + pd.to_timedelta(df['DaysOfARVRefill2'], unit='D') 
        df['IITDate2'] = (df['NextAppt'] + pd.Timedelta(days=29)).fillna('1900')
         
        dflamisnmrs = cleandflamisnmrs() 
        print(dflamisnmrs)  
        df['LGA2']=df['FacilityName'].map(dflamisnmrs.set_index('Name on NMRS')['LGA'])
        df['STATE2']=df['FacilityName'].map(dflamisnmrs.set_index('Name on NMRS')['STATE'])
        df.loc[df['LGA'].isna(), 'LGA'] = df['LGA2']
        df.loc[df['State'].isna(), 'State'] = df['STATE2']
        
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
        
        #normalise VL datetime back to date
        df['NextAppt'] = df['NextAppt'].apply(parse_date)
        
        dfCurrentYearIIT = df[df['CurrentYearIIT'] == 'CurrentYearIIT']
        dfCurrentYearIIT = dfCurrentYearIIT.sort_values(by=['CaseManager'], ascending=True).reset_index()
        dfCurrentYearIIT.insert(0, 'S/N', dfCurrentYearIIT.index + 1)
        
        dfpreviousyearIIT = df[df['previousyearIIT'] == 'previousyearIIT']
        dfpreviousyearIIT = dfpreviousyearIIT.sort_values(by=['CaseManager'], ascending=True).reset_index()
        dfpreviousyearIIT.insert(0, 'S/N', dfpreviousyearIIT.index + 1)
        
        dfImminentIIT = df[df['ImminentIIT'] == 'ImminentIIT']
        dfImminentIIT = dfImminentIIT.sort_values(by=['CaseManager'], ascending=True).reset_index()
        dfImminentIIT.insert(0, 'S/N', dfImminentIIT.index + 1)
        
        dfsevendaysIIT = df[df['sevendaysIIT'] == 'sevendaysIIT']
        dfsevendaysIIT = dfsevendaysIIT.sort_values(by=['CaseManager'], ascending=True).reset_index()
        dfsevendaysIIT.insert(0, 'S/N', dfsevendaysIIT.index + 1)
        
        dfcurrentmonthexpected = df[df['currentmonthexpected'] == 'currentmonthexpected']
        dfcurrentmonthexpected = dfcurrentmonthexpected.sort_values(by=['CaseManager'], ascending=True).reset_index()
        dfcurrentmonthexpected.insert(0, 'S/N', dfcurrentmonthexpected.index + 1)
        
        #Trime Line lists
        dfCurrentYearIIT = dfCurrentYearIIT[["S/N", "State", "LGA", "FacilityName", "PEPID", "PatientHospitalNo", "uuid", "Sex", "Current_Age", "Surname", "Firstname", "MaritalStatus", "PhoneNo", "Address", "State_of_Residence", "LGA_of_Residence", "DateConfirmedHIV+", "ARTStartDate", "Pharmacy_LastPickupdate", "DaysOfARVRefill", "CurrentARTRegimen", "NextAppt", "CurrentPregnancyStatus", "CurrentViralLoad", "DateResultReceivedFacility", "Alphanumeric_Viral_Load_Result", "LastDateOfSampleCollection", "Outcomes", "Outcomes_Date", "IIT_Date", "CurrentARTStatus", "First_TPT_Pickupdate", "Current_TPT_Received", "PBS_Capturee", "PBS_Capture_Date", "Date_Generated","CaseManager"]]
        dfpreviousyearIIT = dfpreviousyearIIT[["S/N", "State", "LGA", "FacilityName", "PEPID", "PatientHospitalNo", "uuid", "Sex", "Current_Age", "Surname", "Firstname", "MaritalStatus", "PhoneNo", "Address", "State_of_Residence", "LGA_of_Residence", "DateConfirmedHIV+", "ARTStartDate", "Pharmacy_LastPickupdate", "DaysOfARVRefill", "CurrentARTRegimen", "NextAppt", "CurrentPregnancyStatus", "CurrentViralLoad", "DateResultReceivedFacility", "Alphanumeric_Viral_Load_Result", "LastDateOfSampleCollection", "Outcomes", "Outcomes_Date", "IIT_Date", "CurrentARTStatus", "First_TPT_Pickupdate", "Current_TPT_Received", "PBS_Capturee", "PBS_Capture_Date", "Date_Generated","CaseManager"]]
        dfImminentIIT = dfImminentIIT[["S/N", "State", "LGA", "FacilityName", "PEPID", "PatientHospitalNo", "uuid", "Sex", "Current_Age", "Surname", "Firstname", "MaritalStatus", "PhoneNo", "Address", "State_of_Residence", "LGA_of_Residence", "DateConfirmedHIV+", "ARTStartDate", "Pharmacy_LastPickupdate", "DaysOfARVRefill", "CurrentARTRegimen", "NextAppt", "CurrentPregnancyStatus", "CurrentViralLoad", "DateResultReceivedFacility", "Alphanumeric_Viral_Load_Result", "LastDateOfSampleCollection", "Outcomes", "Outcomes_Date", "IIT_Date", "CurrentARTStatus", "First_TPT_Pickupdate", "Current_TPT_Received", "PBS_Capturee", "PBS_Capture_Date", "Date_Generated","CaseManager"]]
        dfsevendaysIIT = dfsevendaysIIT[["S/N", "State", "LGA", "FacilityName", "PEPID", "PatientHospitalNo", "uuid", "Sex", "Current_Age", "Surname", "Firstname", "MaritalStatus", "PhoneNo", "Address", "State_of_Residence", "LGA_of_Residence", "DateConfirmedHIV+", "ARTStartDate", "Pharmacy_LastPickupdate", "DaysOfARVRefill", "CurrentARTRegimen", "NextAppt", "CurrentPregnancyStatus", "CurrentViralLoad", "DateResultReceivedFacility", "Alphanumeric_Viral_Load_Result", "LastDateOfSampleCollection", "Outcomes", "Outcomes_Date", "IIT_Date", "CurrentARTStatus", "First_TPT_Pickupdate", "Current_TPT_Received", "PBS_Capturee", "PBS_Capture_Date", "Date_Generated","CaseManager"]]
        dfcurrentmonthexpected = dfcurrentmonthexpected[["S/N", "State", "LGA", "FacilityName", "PEPID", "PatientHospitalNo", "uuid", "Sex", "Current_Age", "Surname", "Firstname", "MaritalStatus", "PhoneNo", "Address", "State_of_Residence", "LGA_of_Residence", "DateConfirmedHIV+", "ARTStartDate", "Pharmacy_LastPickupdate", "DaysOfARVRefill", "CurrentARTRegimen", "NextAppt", "CurrentPregnancyStatus", "CurrentViralLoad", "DateResultReceivedFacility", "Alphanumeric_Viral_Load_Result", "LastDateOfSampleCollection", "Outcomes", "Outcomes_Date", "IIT_Date", "CurrentARTStatus", "First_TPT_Pickupdate", "Current_TPT_Received", "PBS_Capturee", "PBS_Capture_Date", "Date_Generated","CaseManager"]]
        
        df2nd95Summary = df

        df2nd95Summary['ActiveClients'] = df2nd95Summary['CurrentARTStatus'].apply(lambda x: 1 if x == "Active" else 0)
        df2nd95Summary['currentYearIIT'] = df2nd95Summary['CurrentYearIIT'].apply(lambda x: 1 if x == "CurrentYearIIT" else 0)
        df2nd95Summary['previousyearIIT'] = df2nd95Summary['previousyearIIT'].apply(lambda x: 1 if x == "previousyearIIT" else 0)
        df2nd95Summary['ImminentIIT'] = df2nd95Summary['ImminentIIT'].apply(lambda x: 1 if x == "ImminentIIT" else 0)
        df2nd95Summary['sevendaysIIT'] = df2nd95Summary['sevendaysIIT'].apply(lambda x: 1 if x == "sevendaysIIT" else 0)
        df2nd95Summary['currentmonthexpected'] = df2nd95Summary['currentmonthexpected'].apply(lambda x: 1 if x == "currentmonthexpected" else 0)

        result = df2nd95Summary.groupby(['LGA','FacilityName','CaseManager'])[['ActiveClients', 'currentYearIIT', 'previousyearIIT','ImminentIIT','sevendaysIIT','currentmonthexpected']].sum().reset_index()
        result_check = ['ActiveClients', 'currentYearIIT', 'previousyearIIT', 'ImminentIIT', 'sevendaysIIT', 'currentmonthexpected']
        result = result[(result[result_check] != 0).any(axis=1)]


        result = result.rename(columns={'ActiveClients': 'Active Clients',
                                        'currentYearIIT': 'current FY IIT',
                                        'previousyearIIT': 'previous FY IIT',
                                        'ImminentIIT': 'Imminent IIT',
                                        'sevendaysIIT': 'Next 7 days IIT',
                                        'currentmonthexpected': 'EXPECTED THIS MONTH',})
        result   

        progress_bar['maximum'] = len(df)  # Set to 100 for percentage completion
        for index, row in df.iterrows():
                
            # Update the progress bar value
            progress_bar['value'] = index + 1
            
            # Calculate the percentage of completion
            percentage = ((index + 1) / len(df)) * 100
            
            # Update the status label with the current percentage
            status_label.config(text=f"Conversion in Progress: {index + 1}/{len(df)} ({percentage:.2f}%)")
            
            # Update the GUI to reflect changes
            root.update_idletasks()
            
            # Simulate time-consuming task
            time.sleep(0.000001)
        
        #format and export
        sdate = endDate.strftime("%d-%m-%Y")

        print(sdate)
        input_file_path = f"2ND 95 ANALYSIS BY CMG AS AT_{sdate}xlsx"
        output_file_name = input_file_path.split("/")[-1][:-4]
        status_label2.config(text=f"Just a moment! Formating and Saving Converted File...")
        output_file_path = filedialog.asksaveasfilename(initialdir = '/Desktop', 
                                                        title = 'Select a excel file', 
                                                        filetypes = (('excel file','*.xls'), 
                                                                     ('excel file','*.xlsx')),defaultextension=".xlsx", initialfile=output_file_name)
        if not output_file_path:  # Check if the file save was cancelled
            status_label.config(text="File conversion was cancelled. No file was saved.")
            status_label2.config(text="File Conversion Cancelled!")
            progress_bar['value'] = 0
            return  # Exit the function

        # Create a Pandas Excel writer using XlsxWriter as the engine.
        writer = pd.ExcelWriter(output_file_path, engine="xlsxwriter")

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

        # Close the Pandas Excel writer and output the Excel file.
        writer.close()
        end_time = time.time()  # End time measurement
        total_time = end_time - start_time  # Calculate total time taken
        status_label.config(text=f"Conversion Complete! Time taken: {total_time:.2f} seconds")
        status_label2.config(text=f"Converted File Location: {output_file_path}")
        
def Second95ThrimedRefillRateCMG():
        startDate, endDate = get_selected_date()
        start_time = time.time()  # Start time measurement
        
        status_label.config(text=f"Attempting to read file...")
        input_CMG = filedialog.askopenfilename(initialdir = '/Desktop', 
                                                        title = 'Select a excel file', 
                                                        filetypes = (('excel file','*.xls'), 
                                                                     ('excel file','*.csv'),
                                                                     ('excel file','*.xlsx')))
        if not input_CMG:  # Check if the file selection was cancelled
            status_label.config(text="No file selected. Please select a file to convert.")
            return  # Exit the function
        
        # Determine file type and read accordingly
        file_ext = os.path.splitext(input_CMG)[1].lower()

        if file_ext == '.csv':
            dfCMG = pd.read_csv(input_CMG, 
                   encoding='utf-8', 
                   lineterminator='\n', 
                   quotechar='"', 
                   escapechar='\\', 
                   skip_blank_lines=True)
        
        elif file_ext in ['.xls', '.xlsx']:
            dfCMG = pd.read_excel(input_CMG,sheet_name=0, dtype=object)
        
        status_label.config(text=f"Attempting to read baseline ART line list...")
        input_file_path2 = filedialog.askopenfilename(initialdir = '/Desktop', 
                                                        title = 'Select a excel file', 
                                                        filetypes = (('excel file','*.xls'), 
                                                                     ('excel file','*.csv'),
                                                                     ('excel file','*.xlsx')))
        if not input_file_path2:  # Check if the file selection was cancelled
            status_label.config(text="No file selected. Please select a file to convert.")
            return  # Exit the function
        
        # Determine file type and read accordingly
        file_ext2 = os.path.splitext(input_file_path2)[1].lower()

        if file_ext2 == '.csv':
                
            dfbaseline = pd.read_csv(input_file_path, 
                        encoding='utf-8', 
                        lineterminator='\n', 
                        quotechar='"', 
                        escapechar='\\', 
                        skip_blank_lines=True)
            
            #Convert Date Objects to Date using dateutil
            dfDates = ['DateConfirmedHIV+', 'ARTStartDate', 'Pharmacy_LastPickupdate', 'Pharmacy_LastPickupdate_PreviousQuarter', 'DateofCurrentViralLoad', 'DateResultReceivedFacility', 'LastDateOfSampleCollection', 'Outcomes_Date', 'IIT_Date', 'DOB', 'Date_Transfered_In', 'DateofFirstTLD_Pickup', 'EstimatedNextAppointmentPharmacy', 'Next_Ap_by_careCard', 'IPT_Screening_Date', 'First_TPT_Pickupdate', 'Last_TPT_Pickupdate', 'Current_TPT_Received', 'Date_of_TPT_Outcome', 'DateofCurrent_TBStatus', 'TB_Treatment_Start_Date', 'TB_Treatment_Stop_Date', 'Date_Enrolled_Into_OTZ', 'Date_Enrolled_Into_OTZ_Plus', 'PBS_Capture_Date', 'Date_Generated', 'PBS_Recapture_Date']
            for col in dfDates:
                dfbaseline[col] = dfbaseline[col].apply(parse_date)
                
            #Clean numbers
            dfnumeric = ['AgeAtStartofART', 'AgeinMonths', 'DaysOnART', 'DaysOfARVRefill', 'CurrentViralLoad', 'Current_Age', 'Weight', 'Height', 'BMI', 'Whostage', 'CurrentCD4', 'Days_To_Schedule']
            for col in dfnumeric:
                dfbaseline[col] = pd.to_numeric(dfbaseline[col],errors='coerce')
            
        elif file_ext2 in ['.xls', '.xlsx']:
            dfbaseline = pd.read_excel(input_file_path2,sheet_name=0, dtype=object)
            
            #Convert Date Objects to Date using dateutil
            dfDates = ['DateConfirmedHIV+', 'ARTStartDate', 'Pharmacy_LastPickupdate', 'Pharmacy_LastPickupdate_PreviousQuarter', 'DateofCurrentViralLoad', 'DateResultReceivedFacility', 'LastDateOfSampleCollection', 'Outcomes_Date', 'IIT_Date', 'DOB', 'Date_Transfered_In', 'DateofFirstTLD_Pickup', 'EstimatedNextAppointmentPharmacy', 'Next_Ap_by_careCard', 'IPT_Screening_Date', 'First_TPT_Pickupdate', 'Last_TPT_Pickupdate', 'Current_TPT_Received', 'Date_of_TPT_Outcome', 'DateofCurrent_TBStatus', 'TB_Treatment_Start_Date', 'TB_Treatment_Stop_Date', 'Date_Enrolled_Into_OTZ', 'Date_Enrolled_Into_OTZ_Plus', 'PBS_Capture_Date', 'PBS_Recapture_Date']
            for col in dfDates:
                dfbaseline[col] = dfbaseline[col].apply(parse_date)
                
            #Clean numbers
            dfnumeric = ['AgeAtStartofART', 'AgeinMonths', 'DaysOnART', 'DaysOfARVRefill', 'CurrentViralLoad', 'Current_Age', 'Weight', 'Height', 'BMI', 'Whostage', 'CurrentCD4', 'Days_To_Schedule']
            for col in dfnumeric:
                dfbaseline[col] = pd.to_numeric(dfbaseline[col],errors='coerce')
        print(dfbaseline)

        
        status_label.config(text=f"Attempting to read file...")
        input_file_path = filedialog.askopenfilename(initialdir = '/Desktop', 
                                                        title = 'Select a excel file', 
                                                        filetypes = (('excel file','*.xls'), 
                                                                     ('excel file','*.csv'),
                                                                     ('excel file','*.xlsx')))
        if not input_file_path:  # Check if the file selection was cancelled
            status_label.config(text="No file selected. Please select a file to convert.")
            return  # Exit the function
        
        # Determine file type and read accordingly
        file_ext = os.path.splitext(input_file_path)[1].lower()

        if file_ext == '.csv':
                
            df = pd.read_csv(input_file_path, 
                   encoding='utf-8', 
                   lineterminator='\n', 
                   quotechar='"', 
                   escapechar='\\', 
                   skip_blank_lines=True)
            
            #Convert Date Objects to Date using dateutil
            dfDates = ['DateConfirmedHIV+', 'ARTStartDate', 'Pharmacy_LastPickupdate', 'Pharmacy_LastPickupdate_PreviousQuarter', 'DateofCurrentViralLoad', 'DateResultReceivedFacility', 'LastDateOfSampleCollection', 'Outcomes_Date', 'IIT_Date', 'DOB', 'Date_Transfered_In', 'DateofFirstTLD_Pickup', 'EstimatedNextAppointmentPharmacy', 'Next_Ap_by_careCard', 'IPT_Screening_Date', 'First_TPT_Pickupdate', 'Last_TPT_Pickupdate', 'Current_TPT_Received', 'Date_of_TPT_Outcome', 'DateofCurrent_TBStatus', 'TB_Treatment_Start_Date', 'TB_Treatment_Stop_Date', 'Date_Enrolled_Into_OTZ', 'Date_Enrolled_Into_OTZ_Plus', 'PBS_Capture_Date', 'Date_Generated', 'PBS_Recapture_Date']
            for col in dfDates:
                df[col] = df[col].apply(parse_date)
                
            #Clean numbers
            dfnumeric = ['AgeAtStartofART', 'AgeinMonths', 'DaysOnART', 'DaysOfARVRefill', 'CurrentViralLoad', 'Current_Age', 'Weight', 'Height', 'BMI', 'Whostage', 'CurrentCD4', 'Days_To_Schedule']
            for col in dfnumeric:
                df[col] = pd.to_numeric(df[col],errors='coerce')
            
        elif file_ext in ['.xls', '.xlsx']:
            df = pd.read_excel(input_file_path,sheet_name=0, dtype=object)
            
            #Convert Date Objects to Date using dateutil
            dfDates = ['DateConfirmedHIV+', 'ARTStartDate', 'Pharmacy_LastPickupdate', 'Pharmacy_LastPickupdate_PreviousQuarter', 'DateofCurrentViralLoad', 'DateResultReceivedFacility', 'LastDateOfSampleCollection', 'Outcomes_Date', 'IIT_Date', 'DOB', 'Date_Transfered_In', 'DateofFirstTLD_Pickup', 'EstimatedNextAppointmentPharmacy', 'Next_Ap_by_careCard', 'IPT_Screening_Date', 'First_TPT_Pickupdate', 'Last_TPT_Pickupdate', 'Current_TPT_Received', 'Date_of_TPT_Outcome', 'DateofCurrent_TBStatus', 'TB_Treatment_Start_Date', 'TB_Treatment_Stop_Date', 'Date_Enrolled_Into_OTZ', 'Date_Enrolled_Into_OTZ_Plus', 'PBS_Capture_Date', 'PBS_Recapture_Date']
            for col in dfDates:
                df[col] = df[col].apply(parse_date)
                
            #Clean numbers
            dfnumeric = ['AgeAtStartofART', 'AgeinMonths', 'DaysOnART', 'DaysOfARVRefill', 'CurrentViralLoad', 'Current_Age', 'Weight', 'Height', 'BMI', 'Whostage', 'CurrentCD4', 'Days_To_Schedule']
            for col in dfnumeric:
                df[col] = pd.to_numeric(df[col],errors='coerce')
                
        #add case managers and next appt to the dataframe
        df['CaseManager']=df['uuid'].map(dfCMG.set_index('uuid')['CASE MANAGER'])
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
        df['BaselinePharmacy_LastPickupdate'] = pd.to_datetime(df['BaselinePharmacy_LastPickupdate'], errors='coerce').fillna(pd.to_datetime('1900'))
        df['BaselineDaysOfARVRefill'] = df['BaselineDaysOfARVRefill'].apply(lambda x: 0 if x > 180 else x)
        df['BaselineNextAppt'] = df['BaselinePharmacy_LastPickupdate'] + pd.to_timedelta(df['BaselineDaysOfARVRefill'], unit='D') 
        
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
         
        dflamisnmrs = cleandflamisnmrs() 
        print(dflamisnmrs)  
        df['LGA2']=df['FacilityName'].map(dflamisnmrs.set_index('Name on NMRS')['LGA'])
        df['STATE2']=df['FacilityName'].map(dflamisnmrs.set_index('Name on NMRS')['STATE'])
        df.loc[df['LGA'].isna(), 'LGA'] = df['LGA2']
        df.loc[df['State'].isna(), 'State'] = df['STATE2']
        
        df = df.drop(['LGA2','STATE2'], axis=1)
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
        
        #normalise VL datetime back to date
        df['NextAppt'] = df['NextAppt'].apply(parse_date)
        
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

        progress_bar['maximum'] = len(df)  # Set to 100 for percentage completion
        for index, row in df.iterrows():
                
            # Update the progress bar value
            progress_bar['value'] = index + 1
            
            # Calculate the percentage of completion
            percentage = ((index + 1) / len(df)) * 100
            
            # Update the status label with the current percentage
            status_label.config(text=f"Conversion in Progress: {index + 1}/{len(df)} ({percentage:.2f}%)")
            
            # Update the GUI to reflect changes
            root.update_idletasks()
            
            # Simulate time-consuming task
            time.sleep(0.000001)
        
        #format and export
        sdate = endDate.strftime("%d-%m-%Y")

        print(sdate)
        input_file_path = f"2ND 95 ANALYSIS WITH CMG AS AT_{sdate}xlsx"
        output_file_name = input_file_path.split("/")[-1][:-4]
        status_label2.config(text=f"Just a moment! Formating and Saving Converted File...")
        output_file_path = filedialog.asksaveasfilename(initialdir = '/Desktop', 
                                                        title = 'Select a excel file', 
                                                        filetypes = (('excel file','*.xls'), 
                                                                     ('excel file','*.xlsx')),defaultextension=".xlsx", initialfile=output_file_name)
        if not output_file_path:  # Check if the file save was cancelled
            status_label.config(text="File conversion was cancelled. No file was saved.")
            status_label2.config(text="File Conversion Cancelled!")
            progress_bar['value'] = 0
            return  # Exit the function

        # Create a Pandas Excel writer using XlsxWriter as the engine.
        writer = pd.ExcelWriter(output_file_path, engine="xlsxwriter")

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
                
                
                #header_format.set_align('center')
                #total_row = len(dataframe) + 2
                #worksheet.merge_range(total_row, 0, total_row, 1, 'Total', header_format)
                #for col_num in range(1, len(dataframe.columns)):
                    #col_letter = chr(65 + col_num)  # Convert column number to letter
                    #worksheet.write_formula(total_row, col_num, f'SUM({col_letter}2:{col_letter}{total_row})', header_format)
            
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
        end_time = time.time()  # End time measurement
        total_time = end_time - start_time  # Calculate total time taken
        status_label.config(text=f"Conversion Complete! Time taken: {total_time:.2f} seconds")
        status_label2.config(text=f"Converted File Location: {output_file_path}")
        

def ThrimedLineListCMG():
        startDate, endDate = get_selected_date()
        start_time = time.time()  # Start time measurement
        
        status_label.config(text=f"Attempting to read file...")
        input_CMG = filedialog.askopenfilename(initialdir = '/Desktop', 
                                                        title = 'Select a excel file', 
                                                        filetypes = (('excel file','*.xls'), 
                                                                     ('excel file','*.csv'),
                                                                     ('excel file','*.xlsx')))
        if not input_CMG:  # Check if the file selection was cancelled
            status_label.config(text="No file selected. Please select a file to convert.")
            return  # Exit the function
        
        # Determine file type and read accordingly
        file_ext = os.path.splitext(input_CMG)[1].lower()

        if file_ext == '.csv':
            dfCMG = pd.read_csv(input_CMG, 
                   encoding='utf-8', 
                   lineterminator='\n', 
                   quotechar='"', 
                   escapechar='\\', 
                   skip_blank_lines=True)
        
        elif file_ext in ['.xls', '.xlsx']:
            dfCMG = pd.read_excel(input_CMG,sheet_name=0, dtype=object)
        
        status_label.config(text=f"Attempting to read file...")
        input_file_path = filedialog.askopenfilename(initialdir = '/Desktop', 
                                                        title = 'Select a excel file', 
                                                        filetypes = (('excel file','*.xls'), 
                                                                     ('excel file','*.csv'),
                                                                     ('excel file','*.xlsx')))
        if not input_file_path:  # Check if the file selection was cancelled
            status_label.config(text="No file selected. Please select a file to convert.")
            return  # Exit the function
        
        # Determine file type and read accordingly
        file_ext = os.path.splitext(input_file_path)[1].lower()

        if file_ext == '.csv':
                
            df = pd.read_csv(input_file_path, 
                   encoding='utf-8', 
                   lineterminator='\n', 
                   quotechar='"', 
                   escapechar='\\', 
                   skip_blank_lines=True)
            
            #Convert Date Objects to Date using dateutil
            dfDates = ['DateConfirmedHIV+', 'ARTStartDate', 'Pharmacy_LastPickupdate', 'Pharmacy_LastPickupdate_PreviousQuarter', 'DateofCurrentViralLoad', 'DateResultReceivedFacility', 'LastDateOfSampleCollection', 'Outcomes_Date', 'IIT_Date', 'DOB', 'Date_Transfered_In', 'DateofFirstTLD_Pickup', 'EstimatedNextAppointmentPharmacy', 'Next_Ap_by_careCard', 'IPT_Screening_Date', 'First_TPT_Pickupdate', 'Last_TPT_Pickupdate', 'Current_TPT_Received', 'Date_of_TPT_Outcome', 'DateofCurrent_TBStatus', 'TB_Treatment_Start_Date', 'TB_Treatment_Stop_Date', 'Date_Enrolled_Into_OTZ', 'Date_Enrolled_Into_OTZ_Plus', 'PBS_Capture_Date', 'Date_Generated', 'PBS_Recapture_Date']
            for col in dfDates:
                df[col] = df[col].apply(parse_date)
                
            #Clean numbers
            dfnumeric = ['AgeAtStartofART', 'AgeinMonths', 'DaysOnART', 'DaysOfARVRefill', 'CurrentViralLoad', 'Current_Age', 'Weight', 'Height', 'BMI', 'Whostage', 'CurrentCD4', 'Days_To_Schedule']
            for col in dfnumeric:
                df[col] = pd.to_numeric(df[col],errors='coerce')
            
        elif file_ext in ['.xls', '.xlsx']:
            df = pd.read_excel(input_file_path,sheet_name=0, dtype=object)
            
            #Convert Date Objects to Date using dateutil
            dfDates = ['DateConfirmedHIV+', 'ARTStartDate', 'Pharmacy_LastPickupdate', 'Pharmacy_LastPickupdate_PreviousQuarter', 'DateofCurrentViralLoad', 'DateResultReceivedFacility', 'LastDateOfSampleCollection', 'Outcomes_Date', 'IIT_Date', 'DOB', 'Date_Transfered_In', 'DateofFirstTLD_Pickup', 'EstimatedNextAppointmentPharmacy', 'Next_Ap_by_careCard', 'IPT_Screening_Date', 'First_TPT_Pickupdate', 'Last_TPT_Pickupdate', 'Current_TPT_Received', 'Date_of_TPT_Outcome', 'DateofCurrent_TBStatus', 'TB_Treatment_Start_Date', 'TB_Treatment_Stop_Date', 'Date_Enrolled_Into_OTZ', 'Date_Enrolled_Into_OTZ_Plus', 'PBS_Capture_Date', 'PBS_Recapture_Date']
            for col in dfDates:
                df[col] = df[col].apply(parse_date)
                
            #Clean numbers
            dfnumeric = ['AgeAtStartofART', 'AgeinMonths', 'DaysOnART', 'DaysOfARVRefill', 'CurrentViralLoad', 'Current_Age', 'Weight', 'Height', 'BMI', 'Whostage', 'CurrentCD4', 'Days_To_Schedule']
            for col in dfnumeric:
                df[col] = pd.to_numeric(df[col],errors='coerce')
                
        #add case managers and next appt to the dataframe
        df['CaseManager']=df['uuid'].map(dfCMG.set_index('uuid')['CASE MANAGER'])
        df['CaseManager'] = df['CaseManager'].fillna('UNASSIGNED')
         
        # Ensure the 'Pharmacy_LastPickupdate' column is in datetime format and fill NaNs with a specific date
        df['Pharmacy_LastPickupdate2'] = pd.to_datetime(df['Pharmacy_LastPickupdate'], errors='coerce').fillna(pd.to_datetime('1900'))

        #Fill zero if the column contains number greater than 180
        df['DaysOfARVRefill2'] = df['DaysOfARVRefill'].apply(lambda x: 0 if x > 180 else x)

        # Calculate the 'NextAppt' column by adding the 'DaysOfARVRefill2' to 'Pharmacy_LastPickupdate2'
        df['NextAppt'] = df['Pharmacy_LastPickupdate2'] + pd.to_timedelta(df['DaysOfARVRefill2'], unit='D') 
        df['IITDate2'] = (df['NextAppt'] + pd.Timedelta(days=29)).fillna('1900')
         
        dflamisnmrs = cleandflamisnmrs() 
        print(dflamisnmrs)  
        df['LGA2']=df['FacilityName'].map(dflamisnmrs.set_index('Name on NMRS')['LGA'])
        df['STATE2']=df['FacilityName'].map(dflamisnmrs.set_index('Name on NMRS')['STATE'])
        #df.loc[df['LGA'].isna(), 'LGA'] = df['LGA2']
        #df.loc[df['State'].isna(), 'State'] = df['STATE2']
        df['LGA'] = df['LGA'].where(df['LGA'].notna(), df['LGA2'])
        df['State'] = df['State'].where(df['State'].notna(), df['STATE2'])
        
        df = df.drop(['LGA2','STATE2'], axis=1)
        df.insert(0, 'S/N', df.index + 1)
             
        #normalise VL datetime back to date
        df['NextAppt'] = df['NextAppt'].apply(parse_date)
                    
        #Trime Line lists
        df = df[["S/N", "State", "LGA", "FacilityName", "PEPID", "PatientHospitalNo", "uuid", "Sex", "Current_Age", "Surname", "Firstname", "MaritalStatus", "PhoneNo", "Address", "State_of_Residence", "LGA_of_Residence", "DateConfirmedHIV+", "ARTStartDate", "Pharmacy_LastPickupdate", "DaysOfARVRefill", "CurrentARTRegimen", "NextAppt", "CurrentPregnancyStatus", "CurrentViralLoad", "DateResultReceivedFacility", "Alphanumeric_Viral_Load_Result", "LastDateOfSampleCollection", "Outcomes", "Outcomes_Date", "IIT_Date", "CurrentARTStatus", "First_TPT_Pickupdate", "Current_TPT_Received", "PBS_Capturee", "PBS_Capture_Date", "Date_Generated","CaseManager"]]
        
        progress_bar['maximum'] = len(df)  # Set to 100 for percentage completion
        for index, row in df.iterrows():
                
            # Update the progress bar value
            progress_bar['value'] = index + 1
            
            # Calculate the percentage of completion
            percentage = ((index + 1) / len(df)) * 100
            
            # Update the status label with the current percentage
            status_label.config(text=f"Conversion in Progress: {index + 1}/{len(df)} ({percentage:.2f}%)")
            
            # Update the GUI to reflect changes
            root.update_idletasks()
            
            # Simulate time-consuming task
            time.sleep(0.000001)
        
        #format and export
        sdate = endDate.strftime("%d-%m-%Y")

        print(sdate)
        input_file_path = f"CLEANED THRIMED ART LINE LIST WITH CMG_{sdate}xlsx"
        output_file_name = input_file_path.split("/")[-1][:-4]
        status_label2.config(text=f"Just a moment! Formating and Saving Converted File...")
        output_file_path = filedialog.asksaveasfilename(initialdir = '/Desktop', 
                                                        title = 'Select a excel file', 
                                                        filetypes = (('excel file','*.xls'), 
                                                                     ('excel file','*.xlsx')),defaultextension=".xlsx", initialfile=output_file_name)
        if not output_file_path:  # Check if the file save was cancelled
            status_label.config(text="File conversion was cancelled. No file was saved.")
            status_label2.config(text="File Conversion Cancelled!")
            progress_bar['value'] = 0
            return  # Exit the function

        # Create a Pandas Excel writer using XlsxWriter as the engine.
        writer = pd.ExcelWriter(output_file_path, engine="xlsxwriter")

        # List of dataframes and their corresponding sheet names
        dataframes = {
            "ART LINE LIST": df,
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
                    worksheet.write_formula(total_row, col_num, f'SUM({col_letter}2:{col_letter}{total_row})', header_format)
            
            else:
                # Write the dataframe to the sheet, starting from row 1
                dataframe.to_excel(writer, sheet_name=sheet_name, startrow=1, header=False, index=False)
                
                # Write the column headers with the defined format, starting from row 0
                for col_num, value in enumerate(dataframe.columns.values):
                    worksheet.write(0, col_num, value, header_format)
            
            # Add row band format to specific sheets
            if sheet_name in ["ART LINE LIST"]:
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
        end_time = time.time()  # End time measurement
        total_time = end_time - start_time  # Calculate total time taken
        status_label.config(text=f"Conversion Complete! Time taken: {total_time:.2f} seconds")
        status_label2.config(text=f"Converted File Location: {output_file_path}")
        

#Creating A tooltip Class
class ToolTip(object):

    def __init__(self, widget):
        self.widget = widget
        self.tipwindow = None
        self.id = None
        self.x = self.y = 0

    def showtip(self, text):
        "Display text in tooltip window"
        self.text = text
        if self.tipwindow or not self.text:
            return
        x, y, cx, cy = self.widget.bbox("insert")
        x = x + self.widget.winfo_rootx() + 260
        y = y + cy + self.widget.winfo_rooty() +27
        self.tipwindow = tw = Toplevel(self.widget)
        tw.wm_overrideredirect(1)
        tw.wm_geometry("+%d+%d" % (x, y))
        label = Label(tw, text=self.text, justify=LEFT,
                      background="#ffffe0", relief=SOLID, borderwidth=1,
                      font=("tahoma", "8", "normal"))
        label.pack(ipadx=1)

    def hidetip(self):
        tw = self.tipwindow
        self.tipwindow = None
        if tw:
            tw.destroy()

def CreateToolTip(widget, text):
    toolTip = ToolTip(widget)
    def enter(event):
        toolTip.showtip(text)
    def leave(event):
        toolTip.hidetip()
    widget.bind('<Enter>', enter)
    widget.bind('<Leave>', leave)


def open_editor():
    editor_window = tk.Toplevel(root)
    tree = ttk.Treeview(editor_window)

    # Define columns
    tree['columns'] = list(dflamisnmrs.columns)

    # Format columns
    tree.column("#0", width=0, stretch=tk.NO)
    for col in tree['columns']:
        tree.column(col, anchor=tk.W, width=100)

    # Create headings
    tree.heading("#0", text='', anchor=tk.W)
    for col in tree['columns']:
        tree.heading(col, text=col, anchor=tk.W)

    # Insert data
    for index, row in dflamisnmrs.iterrows():
        tree.insert("", tk.END, values=list(row))

    # Pack treeview
    tree.pack(fill=tk.BOTH, expand=True)

    # Create edit button
    def edit_row():
        # Get selected row
        selected = tree.focus()
        values = tree.item(selected, 'values')

        # Create edit window
        edit_window = tk.Toplevel(editor_window)
        entries = []

        for col in tree['columns']:
            label = tk.Label(edit_window, text=col)
            label.grid(row=len(entries), column=0)
            entry = tk.Entry(edit_window)
            entry.grid(row=len(entries), column=1)
            entries.append(entry)

        # Fill entries with selected row values
        for i, value in enumerate(values):
            entries[i].insert(0, value)

        # Create save button
        def save_edits():
            # Get new values
            values = [entry.get() for entry in entries]

            # Update treeview
            tree.item(selected, values=values)

            # Close edit window
            edit_window.destroy()

        save_button = tk.Button(edit_window, text="Save", command=save_edits)
        save_button.grid(row=len(entries), column=1)

    edit_button = tk.Button(editor_window, text="Edit", command=edit_row)
    edit_button.pack()

    # Create save button
    def save_changes():
        # Get updated data from treeview
        data = []
        for child in tree.get_children():
            data.append(tree.item(child, 'values'))

        # Update DataFrame
        global dflamisnmrs
        dflamisnmrs = pd.DataFrame(data, columns=dflamisnmrs.columns)

        # Print updated DataFrame
        print(dflamisnmrs)

    save_button = tk.Button(editor_window, text="Save", command=save_changes)
    save_button.pack()

# Creating Main Window
root = tk.Tk()
root.title("NETO's NMRS ART LINE LIST 2ND AND 3RD 95 ANALYSER v1.0")
root.geometry("600x450")
root.config(bg="#f0f0f0")

frame3 = tk.Frame(root)
frame3.place(relx=1.0, rely=0, anchor='ne') # Position at the top right

#date text
selectinfo = tk.Label(root, text="Please remember to adjust start date and end date if necessary", font=("Helvetica", 10), fg="red")
selectinfo.pack(pady=1)

# Create a frame to hold the DateEntry widgets

frame = tk.Frame(root)
frame.pack(padx=10, pady=2)

frame2 = tk.Frame(root)
frame2.pack(padx=10, pady=0.5)

# Get the current date
current_date = datetime.now().date()

# Create a settings button in the top right corner
editbutton = tk.Button(frame3, text="edit fac", command=open_editor)
editbutton.pack(side=tk.RIGHT, padx=20, pady=10)

# Create DateEntry widgets with date format and select mode
start_date_label = tk.Label(frame2, text="Start Date", font=("Helvetica", 9))
start_date_label.pack(side=tk.LEFT, padx=20)
cal = DateEntry(frame, width=12, background="darkblue", foreground="white", borderwidth=2, date_pattern='dd-mm-yyyy', selectmode='day', year=2000, month=1, day=1)
cal.pack(side=tk.LEFT, padx=5)

end_date_label = tk.Label(frame2, text="End Date", font=("Helvetica", 9))
end_date_label.pack(side=tk.LEFT, padx=20)
cal2 = DateEntry(frame, width=12, background="darkblue", foreground="white", borderwidth=2, date_pattern='dd-mm-yyyy', selectmode='day')
cal2.set_date(current_date)
cal2.pack(side=tk.LEFT, padx=5)

# Create a StringVar to store the selected option
selected_option = tk.StringVar(root)
selected_option.set("3RD 95 ▼")  # Set an initial default value

# Create a frame to hold the label and OptionMenu
select_frame = tk.Frame(root)
select_frame.pack()

# Create a label for the selection text
select_label = tk.Label(select_frame, text="Select:", font=("Helvetica", 12))
select_label.pack(side="left")

# Create the OptionMenu with invisible dropdown arrow
option_menu = tk.OptionMenu(select_frame, selected_option, 
"3RD 95 ▼",
"3RD 95 Thrimed ▼",
"3RD 95 Thrimed WITH CMG ▼",
"2ND 95 Thrimed ▼",
"2ND 95 Thrimed with Refill Rate ▼",
"2ND 95 Thrimed WITH CMG ▼",
"2ND 95 Thrimed with Refill Rate With CMG ▼",
"CLEANED THRIMED ART LINE LIST WITH CMG ▼")

option_menu.config(indicatoron=0)  # Hide the arrow
option_menu.pack(side="left", padx=5, pady=10)

# Create a button to call the function based on the selected option
def on_dropdown_click():
    selected_value = selected_option.get()
    if selected_value == "3RD 95 ▼":
        third95()
    elif selected_value == "3RD 95 Thrimed ▼":    
        third95Thrimed()
    elif selected_value == "3RD 95 Thrimed WITH CMG ▼":    
        third95ThrimedCMG()
    elif selected_value == "2ND 95 Thrimed ▼":    
        Second95Thrimed()
    elif selected_value == "2ND 95 Thrimed with Refill Rate ▼":    
        Second95ThrimedRefillRate()
    elif selected_value == "2ND 95 Thrimed WITH CMG ▼":    
        Second95ThrimedCMG()
    elif selected_value == "2ND 95 Thrimed with Refill Rate With CMG ▼":    
        Second95ThrimedRefillRateCMG()
    elif selected_value == "CLEANED THRIMED ART LINE LIST WITH CMG ▼":    
        ThrimedLineListCMG()
        
        

dropdown_button = tk.Button(root, text="PROCESS SELECTION", command=on_dropdown_click, font=("Helvetica", 14), bg="green", fg="#ffffff")
dropdown_button.place(x=200, y=50)
dropdown_button.pack(pady=0.5)

convert_button_exit = tk.Button(root, text="EXIT APPLICATION.......", command=root.destroy, font=("Helvetica", 14), bg="red", fg="#ffffff")
convert_button_exit.pack(pady=0.5)

# Progress bar widget
progress_bar = ttk.Progressbar(root, orient='horizontal', length=300, mode='determinate')
progress_bar.pack(pady=20)

# Labels
status_label = tk.Label(root, text="0%", font=('Helvetica', 12))
status_label.pack(pady=0.5)

status_label2 = tk.Label(root, text='WELCOME TO NETO NMRS ART LINE LIST 2ND AND 3RD 95 ANALYSER', bg="#D3D3D3", font=('Helvetica', 9))
status_label2.pack(pady=20)

text3 = tk.Label(root, text="you will be prompted to select required files and the location you want to save the converted file \n(Please Remember to follow the file requirement procedures)")
text3.pack(pady=1)

text2 = tk.Label(root, text="Contacts: email: chinedum.pius@gmail.com, phone: +2348134453841")
text2.pack(pady=19)



# Adding File Dialog
filedialog = tk.filedialog 

# Running the GUI
root.mainloop()

#python -m nuitka --windows-console-mode=disable --onefile --enable-plugin=tk-inter nmrsAnalysis.py