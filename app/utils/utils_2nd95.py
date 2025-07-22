import pandas as pd

def compute_appointment_and_iit_dates(df):
    df['Pharmacy_LastPickupdate2'] = pd.to_datetime(df['Pharmacy_LastPickupdate'], errors='coerce')
    df['Duration'] = pd.to_numeric(df['DaysOfARVRefill'], errors='coerce')
    df['NextAppt'] = df['Pharmacy_LastPickupdate2'] + pd.to_timedelta(df['Duration'], unit='D')
    df['IITDate2'] = df['NextAppt'] + pd.Timedelta(days=29)
    return df

def classify_iit_Appt_status(df, endDate):
    today = pd.to_datetime(endDate)
    currentyear = today.year
    previousyear = today.year - 1
    currentmonthyear = f"{today.month}_{today.year}"
    sevenDaysIIT = today + pd.Timedelta(days=7)
    currentWeekYear = f"{today.isocalendar().week}_{today.year}"
    currentWeek = today.isocalendar().week

    # Ensure datetime formats
    df['IITDate2'] = pd.to_datetime(df['IITDate2'], errors='coerce')
    df['NextAppt'] = pd.to_datetime(df['NextAppt'], errors='coerce')

    # Generate derived columns
    df['IITYear'] = df['IITDate2'].dt.year
    df['NextApptMonthYear'] = df['NextAppt'].dt.month.fillna(0).astype(int).astype(str) + '_' + df['NextAppt'].dt.year.fillna(0).astype(int).astype(str)

    # Add IIT status flags
    df['CurrentYearIIT'] = df.apply(
        lambda row: 'CurrentYearIIT' if (
            (row['IITYear'] == currentyear) and 
            (row['IITDate2'] <= today) and 
            (row['CurrentARTStatus'] in ['LTFU', 'Lost to followup'])
        ) else 'notCurrentYearIIT',
        axis=1
    )

    df['previousyearIIT'] = df.apply(
        lambda row: 'previousyearIIT' if (
            (row['IITYear'] == previousyear) and 
            (row['IITDate2'] <= today) and 
            (row['CurrentARTStatus'] in ['LTFU', 'Lost to followup'])
        ) else 'notpreviousyearIIT',
        axis=1
    )

    df['ImminentIIT'] = df.apply(
        lambda row: 'ImminentIIT' if (
            pd.notna(row['NextAppt']) and
            (row['NextAppt'] <= today) and 
            (row['CurrentARTStatus'] == 'Active')
        ) else 'notImminentIIT',
        axis=1
    )

    df['sevendaysIIT'] = df.apply(
        lambda row: 'sevendaysIIT' if (
            pd.notna(row['IITDate2']) and
            (row['IITDate2'] <= sevenDaysIIT) and 
            (row['CurrentARTStatus'] == 'Active')
        ) else 'notsevenDaysIIT',
        axis=1
    )

    df['currentmonthexpected'] = df.apply(
        lambda row: 'currentmonthexpected' if (
            (row['NextApptMonthYear'] == currentmonthyear) and 
            (row['CurrentARTStatus'] == 'Active')
        ) else 'notcurrentmonthexpected',
        axis=1
    )
    
    baseline_columns = {
        'BaselineNextAppt_week_year',
        'BaselineCurrentARTStatus',
        'BaselinePharmacy_LastPickupdate',
        'BaselineNextAppt',
    }
    if baseline_columns.issubset(df.columns):
        df['NextAppt'] = pd.to_datetime(df['NextAppt'], errors='coerce').fillna(pd.to_datetime('1900'))
        df['BaselineNextAppt'] = pd.to_datetime(df['BaselineNextAppt'], errors='coerce').fillna(pd.to_datetime('1900'))
        df['currentweekexpected'] = df.apply(lambda row: 'currentweekexpected' if ((row['BaselineNextAppt_week_year'] == currentWeekYear) & ((row['CurrentARTStatus'] == 'Active')) & ((row['BaselineCurrentARTStatus'] == 'Active'))) else 'notcurrentweekexpected', axis=1)
        df['weeklyexpectedrefilled'] = df.apply(lambda row: 'weeklyexpectedrefilled' if ((row['currentweekexpected'] == 'currentweekexpected') & ((row['Pharmacy_LastPickupdate2'] > row['BaselinePharmacy_LastPickupdate']) & (row['NextAppt'] > row['BaselineNextAppt']))) else 'notweeklyexpectedrefilled', axis=1)
        df['pendingweeklyrefill'] = df.apply(lambda row: 'pendingweeklyrefill' if ((row['BaselineNextAppt_week_year'] == currentWeekYear) & ((row['CurrentARTStatus'] == 'Active')) & ((row['BaselineCurrentARTStatus'] == 'Active')) & ((row['weeklyexpectedrefilled'] == 'notweeklyexpectedrefilled'))) else 'notpendingweeklyrefill', axis=1)

    return df

def integrate_baseline_data(df, dfbaseline):
    
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
        
    return df


