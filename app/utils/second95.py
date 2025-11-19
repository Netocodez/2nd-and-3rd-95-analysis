from flask import Flask, render_template, request, jsonify, send_file
import pandas as pd
import os
from io import BytesIO
from datetime import datetime
from .emr_processor import process_Linelist, columns_to_select, columns_to_select2, export_to_excel_with_formatting, sc_gap_mask, calculate_age_vectorized
from .utils_2nd95 import (
    compute_appointment_and_iit_dates,
    classify_iit_Appt_status,
    trackBiometrics,
    integrate_baseline_data
)

def second95(df, endDate):
    
    endDate = pd.to_datetime(endDate)
    
    df = compute_appointment_and_iit_dates(df)
    df = classify_iit_Appt_status(df, endDate) #adding relevant columns for IIT and appointment status
    df = trackBiometrics(df, endDate) #adding relevant columns for biometrics tracking
    
    #df.to_excel('temp.xlsx', index=False)
    df['DOB'] = df['DOB'].astype(str).str.strip()
    df['DOB'] = pd.to_datetime(df['DOB'], errors='coerce', infer_datetime_format=True).fillna(pd.to_datetime('1900'))
    
    df['Current_Age'] = calculate_age_vectorized(df, 'DOB', ref_date=endDate)
    
    #Generate line lists
    dfCurrentYearIIT = process_Linelist(df, 'CurrentYearIIT', 'CurrentYearIIT', columns_to_select2)
    dfpreviousyearIIT = process_Linelist(df, 'previousyearIIT', 'previousyearIIT', columns_to_select2)    
    dfImminentIIT = process_Linelist(df, 'ImminentIIT', 'ImminentIIT', columns_to_select2)
    dfsevendaysIIT = process_Linelist(df, 'sevendaysIIT', 'sevendaysIIT', columns_to_select2)
    dfcurrentmonthexpected = process_Linelist(df, 'currentmonthexpected', 'currentmonthexpected', columns_to_select2)
    dfCurrentYearLosses = process_Linelist(df, 'CurrentYearLosses', 'CurrentYearLosses', columns_to_select2)
    dfBiometricsGap = process_Linelist(df, 'BiometricsGap', 'BiometricsGap', columns_to_select2)
    dfWKMissedBiometrics = process_Linelist(df, 'WKMissedBiometrics', 'WKMissedBiometrics', columns_to_select2)
        
    df2nd95Summary = df

    df2nd95Summary['ActiveClients'] = df2nd95Summary['CurrentARTStatus'].apply(lambda x: 1 if x == "Active" else 0)
    df2nd95Summary['currentYearIIT'] = df2nd95Summary['CurrentYearIIT'].apply(lambda x: 1 if x == "CurrentYearIIT" else 0)
    df2nd95Summary['CurrentWeekIIT'] = df2nd95Summary['CurrentWeekIIT'].apply(lambda x: 1 if x == "CurrentWeekIIT" else 0)
    df2nd95Summary['previousyearIIT'] = df2nd95Summary['previousyearIIT'].apply(lambda x: 1 if x == "previousyearIIT" else 0)
    df2nd95Summary['ImminentIIT'] = df2nd95Summary['ImminentIIT'].apply(lambda x: 1 if x == "ImminentIIT" else 0)
    df2nd95Summary['sevendaysIIT'] = df2nd95Summary['sevendaysIIT'].apply(lambda x: 1 if x == "sevendaysIIT" else 0)
    df2nd95Summary['currentmonthexpected'] = df2nd95Summary['currentmonthexpected'].apply(lambda x: 1 if x == "currentmonthexpected" else 0)
    df2nd95Summary['CurrentWeekLosses'] = df2nd95Summary['CurrentWeekLosses'].apply(lambda x: 1 if x == "CurrentWeekLosses" else 0)
    df2nd95Summary['CurrentYearLosses'] = df2nd95Summary['CurrentYearLosses'].apply(lambda x: 1 if x == "CurrentYearLosses" else 0)
    df2nd95Summary['CapturedBiometrics'] = df2nd95Summary['CapturedBiometrics'].apply(lambda x: 1 if x == "CapturedBiometrics" else 0)
    df2nd95Summary['BiometricsGap'] = df2nd95Summary['BiometricsGap'].apply(lambda x: 1 if x == "BiometricsGap" else 0)
    df2nd95Summary['WKMissedBiometrics'] = df2nd95Summary['WKMissedBiometrics'].apply(lambda x: 1 if x == "WKMissedBiometrics" else 0)

    result = df2nd95Summary.groupby(['LGA','FacilityName'])[['ActiveClients', 'currentYearIIT', 'previousyearIIT','ImminentIIT','sevendaysIIT','currentmonthexpected', 'CurrentWeekIIT', 'CurrentWeekLosses', 'CurrentYearLosses', 'CapturedBiometrics','BiometricsGap','WKMissedBiometrics']].sum().reset_index()
    result_check = ['ActiveClients', 'currentYearIIT', 'previousyearIIT', 'ImminentIIT', 'sevendaysIIT', 'currentmonthexpected', 'CurrentWeekIIT', 'CurrentWeekLosses', 'CurrentYearLosses', 'CapturedBiometrics','BiometricsGap','WKMissedBiometrics']
    result = result[(result[result_check] != 0).any(axis=1)]
    
    result['%Biometrics Coverage'] = ((result['CapturedBiometrics'] / result['ActiveClients'])).round(4)
    result = result[['LGA','FacilityName','ActiveClients', 'currentYearIIT', 'previousyearIIT', 'ImminentIIT', 'sevendaysIIT', 'currentmonthexpected', 'CurrentWeekIIT', 'CurrentWeekLosses', 'CurrentYearLosses', 'CapturedBiometrics','%Biometrics Coverage','BiometricsGap','WKMissedBiometrics']]



    result = result.rename(columns={'ActiveClients': 'Active Clients',
                                    'currentYearIIT': 'current FY IIT',
                                    'previousyearIIT': 'previous FY IIT',
                                    'ImminentIIT': 'Imminent IIT',
                                    'sevendaysIIT': 'Next 7 days IIT',
                                    'CurrentWeekIIT': 'Curr Week IIT',
                                    'CurrentWeekLosses': 'Curr WK Losses (Stp, Dead, TO)',
                                    'CurrentYearLosses': 'Curr FY Losses (Stp, Dead, TO)',
                                    'currentmonthexpected': 'EXPECTED THIS MONTH',
                                    'CapturedBiometrics': 'Captured Biometrics',
                                    'BiometricsGap': 'Biometrics Gap',
                                    'WKMissedBiometrics': 'Weekly Missed Biometrics',})
    
    result   
    
    #format and export
    formatted_period = endDate.strftime("%d-%m-%Y")    
    
    # List of dataframes and their corresponding sheet names
    dataframes = {
        "CURRENT FY IIT": dfCurrentYearIIT,
        "PREVIOUS FY IIT": dfpreviousyearIIT,
        "IMMINENT IIT": dfImminentIIT,
        "NEXT 7 DAYS IIT": dfsevendaysIIT,
        "EXPECTED THIS MONTH": dfcurrentmonthexpected,
        "CURRENT FY LOSSES": dfCurrentYearLosses,
        "BIOMETRICS GAP": dfBiometricsGap,
        "WEEKLY MISSED BIOMETRICS": dfWKMissedBiometrics,
        "2ND 95 SUMMARY": result,
        # Add more dataframes and sheet names as needed
    }
    
    #mask = sc_gap_mask(dfcurrentmonthexpected, endDate)
    
    # Build row_masks for all sheets except the summary
    row_masks_dict = {}
    for sheet_name, df in dataframes.items():
        if sheet_name != "2ND 95 SUMMARY":
            # Compute mask for this dataframe
            # Replace sc_gap_mask(df, endDate) with whatever logic applies per sheet
            row_masks_dict[sheet_name] = sc_gap_mask(df, endDate, age_col='Current_Age')
            
    # Dynamic numeric and percentage column configuration
    column_config = {
        "2ND 95 SUMMARY": {
            "numeric_cols": [2,3,4,5,6,7,8,9,10,11,13,14],

            "percent_formulas": {
                "%Biometrics Coverage": "=L{subtotal_row}/C{subtotal_row}"
            },

            # NEW (merge A:C)
            "merge_columns": (0, 1),

            # NEW title override
            "title": "2ND 95 PERFORMANCE REPORT – {period}"
        }
    }

    # Write each dataframe to a different sheet
    filename = export_to_excel_with_formatting(
        dataframes,
        formatted_period,
        summaryName="2ND 95 SUMMARY",
        division_columns={'M': ('L', 'C')},
        color_column=["%Biometrics Coverage"],
        column_widths={'A:A': 20, 'B:B': 35,},
        mergeNum=1, #merge first three columns for the summary total row
        row_masks=row_masks_dict,  # << pass mask here
        column_config=column_config
    )
    return filename


def second95CMG(df, endDate):
    
    endDate = pd.to_datetime(endDate)
            
    #add case managers and next appt to the dataframe
    df['CaseManager'] = df['CaseManager'].fillna('UNASSIGNED')
        
    df = compute_appointment_and_iit_dates(df)
    df = classify_iit_Appt_status(df, endDate) #adding relevant columns for IIT and appointment status
    df = trackBiometrics(df, endDate) #adding relevant columns for biometrics tracking
    
    #df.to_excel('temp.xlsx', index=False)
    df['DOB'] = df['DOB'].astype(str).str.strip()
    df['DOB'] = pd.to_datetime(df['DOB'], errors='coerce', infer_datetime_format=True).fillna(pd.to_datetime('1900'))
    
    df['Current_Age'] = calculate_age_vectorized(df, 'DOB', ref_date=endDate)
        
    dfCurrentYearIIT = process_Linelist(df, 'CurrentYearIIT', 'CurrentYearIIT', columns_to_select, sort_by='CaseManager')
    dfpreviousyearIIT = process_Linelist(df, 'previousyearIIT', 'previousyearIIT', columns_to_select, sort_by='CaseManager')
    dfImminentIIT = process_Linelist(df, 'ImminentIIT', 'ImminentIIT', columns_to_select, sort_by='CaseManager')
    dfsevendaysIIT = process_Linelist(df, 'sevendaysIIT', 'sevendaysIIT', columns_to_select, sort_by='CaseManager')
    dfcurrentmonthexpected = process_Linelist(df, 'currentmonthexpected', 'currentmonthexpected', columns_to_select, sort_by='CaseManager')
    dfCurrentYearLosses = process_Linelist(df, 'CurrentYearLosses', 'CurrentYearLosses', columns_to_select, sort_by='CaseManager')
    dfBiometricsGap = process_Linelist(df, 'BiometricsGap', 'BiometricsGap', columns_to_select, sort_by='CaseManager')
    dfWKMissedBiometrics = process_Linelist(df, 'WKMissedBiometrics', 'WKMissedBiometrics', columns_to_select, sort_by='CaseManager')
    
    df2nd95Summary = df

    df2nd95Summary['ActiveClients'] = df2nd95Summary['CurrentARTStatus'].apply(lambda x: 1 if x == "Active" else 0)
    df2nd95Summary['currentYearIIT'] = df2nd95Summary['CurrentYearIIT'].apply(lambda x: 1 if x == "CurrentYearIIT" else 0)
    df2nd95Summary['CurrentWeekIIT'] = df2nd95Summary['CurrentWeekIIT'].apply(lambda x: 1 if x == "CurrentWeekIIT" else 0)
    df2nd95Summary['previousyearIIT'] = df2nd95Summary['previousyearIIT'].apply(lambda x: 1 if x == "previousyearIIT" else 0)
    df2nd95Summary['ImminentIIT'] = df2nd95Summary['ImminentIIT'].apply(lambda x: 1 if x == "ImminentIIT" else 0)
    df2nd95Summary['sevendaysIIT'] = df2nd95Summary['sevendaysIIT'].apply(lambda x: 1 if x == "sevendaysIIT" else 0)
    df2nd95Summary['currentmonthexpected'] = df2nd95Summary['currentmonthexpected'].apply(lambda x: 1 if x == "currentmonthexpected" else 0)
    df2nd95Summary['CurrentWeekLosses'] = df2nd95Summary['CurrentWeekLosses'].apply(lambda x: 1 if x == "CurrentWeekLosses" else 0)
    df2nd95Summary['CurrentYearLosses'] = df2nd95Summary['CurrentYearLosses'].apply(lambda x: 1 if x == "CurrentYearLosses" else 0)
    df2nd95Summary['CapturedBiometrics'] = df2nd95Summary['CapturedBiometrics'].apply(lambda x: 1 if x == "CapturedBiometrics" else 0)
    df2nd95Summary['BiometricsGap'] = df2nd95Summary['BiometricsGap'].apply(lambda x: 1 if x == "BiometricsGap" else 0)
    df2nd95Summary['WKMissedBiometrics'] = df2nd95Summary['WKMissedBiometrics'].apply(lambda x: 1 if x == "WKMissedBiometrics" else 0)

    result = df2nd95Summary.groupby(['LGA','FacilityName','CaseManager'])[['ActiveClients', 'currentYearIIT', 'previousyearIIT','ImminentIIT','sevendaysIIT','currentmonthexpected', 'CurrentWeekIIT', 'CurrentWeekLosses', 'CurrentYearLosses', 'CapturedBiometrics','BiometricsGap','WKMissedBiometrics']].sum().reset_index()
    result_check = ['ActiveClients', 'currentYearIIT', 'previousyearIIT', 'ImminentIIT', 'sevendaysIIT', 'currentmonthexpected', 'CurrentWeekIIT', 'CurrentWeekLosses', 'CurrentYearLosses', 'CapturedBiometrics','BiometricsGap','WKMissedBiometrics']
    result = result[(result[result_check] != 0).any(axis=1)]
    
    result['%Biometrics Coverage'] = ((result['CapturedBiometrics'] / result['ActiveClients'])).round(4)
    result = result[['LGA','FacilityName','CaseManager','ActiveClients', 'currentYearIIT', 'previousyearIIT', 'ImminentIIT', 'sevendaysIIT', 'currentmonthexpected', 'CurrentWeekIIT', 'CurrentWeekLosses', 'CurrentYearLosses', 'CapturedBiometrics','%Biometrics Coverage','BiometricsGap','WKMissedBiometrics']]



    result = result.rename(columns={'ActiveClients': 'Active Clients',
                                    'currentYearIIT': 'current FY IIT',
                                    'previousyearIIT': 'previous FY IIT',
                                    'ImminentIIT': 'Imminent IIT',
                                    'sevendaysIIT': 'Next 7 days IIT',
                                    'CurrentWeekIIT': 'Curr Week IIT',
                                    'CurrentWeekLosses': 'Curr WK Losses (Stp, Dead, TO)',
                                    'CurrentYearLosses': 'Curr FY Losses (Stp, Dead, TO)',
                                    'currentmonthexpected': 'EXPECTED THIS MONTH',
                                    'CapturedBiometrics': 'Captured Biometrics',
                                    'BiometricsGap': 'Biometrics Gap',
                                    'WKMissedBiometrics': 'Weekly Missed Biometrics',})
    result   
    
    #format and export
    formatted_period = endDate.strftime("%d-%m-%Y")

    # List of dataframes and their corresponding sheet names
    dataframes = {
        "CURRENT FY IIT": dfCurrentYearIIT,
        "PREVIOUS FY IIT": dfpreviousyearIIT,
        "IMMINENT IIT": dfImminentIIT,
        "NEXT 7 DAYS IIT": dfsevendaysIIT,
        "EXPECTED THIS MONTH": dfcurrentmonthexpected,
        "CURRENT FY LOSSES": dfCurrentYearLosses,
        "BIOMETRICS GAP": dfBiometricsGap,
        "WEEKLY MISSED BIOMETRICS": dfWKMissedBiometrics,
        "2ND 95 SUMMARY": result,
        # Add more dataframes and sheet names as needed
    }
    
    # Build row_masks for all sheets except the summary
    row_masks_dict = {}
    for sheet_name, df in dataframes.items():
        if sheet_name != "2ND 95 SUMMARY":
            # Compute mask for this dataframe
            # Replace sc_gap_mask(df, endDate) with whatever logic applies per sheet
            row_masks_dict[sheet_name] = sc_gap_mask(df, endDate, age_col='Current_Age')
            
    # Dynamic numeric and percentage column configuration
    column_config = {
        "2ND 95 SUMMARY": {
            "numeric_cols": [3,4,5,6,7,8,9,10,11,12,14,15],

            "percent_formulas": {
                "%Biometrics Coverage": "=M{subtotal_row}/D{subtotal_row}"
            },

            # NEW (merge A:C)
            "merge_columns": (0, 2),

            # NEW title override
            "title": "2ND 95 PERFORMANCE REPORT – {period}"
        }
    }

    # Write each dataframe to a different sheet
    filename = export_to_excel_with_formatting(
        dataframes,
        formatted_period,
        summaryName="2ND 95 SUMMARY",
        division_columns={'N': ('M', 'D')},
        color_column=["%Weekly Refill Rate","%Biometrics Coverage"],
        column_widths={'A:A': 20, 'B:B': 35, 'C:C': 35,},
        mergeNum=2, #merge first three columns for the summary total row
        row_masks=row_masks_dict,  # << pass mask here
        column_config=column_config
    )
    return filename


def Second95R(df, dfbaseline, endDate): 
    
    endDate = pd.to_datetime(endDate)
    
    df = integrate_baseline_data(df, dfbaseline)
    df = compute_appointment_and_iit_dates(df)
    df = classify_iit_Appt_status(df, endDate) #adding relevant columns for IIT and appointment status
    df = trackBiometrics(df, endDate) #adding relevant columns for biometrics tracking
    
    #df.to_excel('temp.xlsx', index=False)
    df['DOB'] = df['DOB'].astype(str).str.strip()
    df['DOB'] = pd.to_datetime(df['DOB'], errors='coerce', infer_datetime_format=True).fillna(pd.to_datetime('1900'))
    
    df['Current_Age'] = calculate_age_vectorized(df, 'DOB', ref_date=endDate)
        
    dfCurrentYearIIT = process_Linelist(df, 'CurrentYearIIT', 'CurrentYearIIT', columns_to_select2)
    dfpreviousyearIIT = process_Linelist(df, 'previousyearIIT', 'previousyearIIT', columns_to_select2)
    dfImminentIIT = process_Linelist(df, 'ImminentIIT', 'ImminentIIT', columns_to_select2)
    dfsevendaysIIT = process_Linelist(df, 'sevendaysIIT', 'sevendaysIIT', columns_to_select2)
    dfcurrentmonthexpected = process_Linelist(df, 'currentmonthexpected', 'currentmonthexpected', columns_to_select2)
    dfpendingweeklyrefill = process_Linelist(df, 'pendingweeklyrefill', 'pendingweeklyrefill', columns_to_select2)
    dfCurrentYearLosses = process_Linelist(df, 'CurrentYearLosses', 'CurrentYearLosses', columns_to_select2)
    dfBiometricsGap = process_Linelist(df, 'BiometricsGap', 'BiometricsGap', columns_to_select2)
    dfWKMissedBiometrics = process_Linelist(df, 'WKMissedBiometrics', 'WKMissedBiometrics', columns_to_select2)
    
    df2nd95Summary = df

    df2nd95Summary['ActiveClients'] = df2nd95Summary['CurrentARTStatus'].apply(lambda x: 1 if x == "Active" else 0)
    df2nd95Summary['currentYearIIT'] = df2nd95Summary['CurrentYearIIT'].apply(lambda x: 1 if x == "CurrentYearIIT" else 0)
    df2nd95Summary['CurrentWeekIIT'] = df2nd95Summary['CurrentWeekIIT'].apply(lambda x: 1 if x == "CurrentWeekIIT" else 0)
    df2nd95Summary['previousyearIIT'] = df2nd95Summary['previousyearIIT'].apply(lambda x: 1 if x == "previousyearIIT" else 0)
    df2nd95Summary['ImminentIIT'] = df2nd95Summary['ImminentIIT'].apply(lambda x: 1 if x == "ImminentIIT" else 0)
    df2nd95Summary['sevendaysIIT'] = df2nd95Summary['sevendaysIIT'].apply(lambda x: 1 if x == "sevendaysIIT" else 0)
    df2nd95Summary['currentmonthexpected'] = df2nd95Summary['currentmonthexpected'].apply(lambda x: 1 if x == "currentmonthexpected" else 0)
    df2nd95Summary['currentweekexpected'] = df2nd95Summary['currentweekexpected'].apply(lambda x: 1 if x == "currentweekexpected" else 0)
    df2nd95Summary['weeklyexpectedrefilled'] = df2nd95Summary['weeklyexpectedrefilled'].apply(lambda x: 1 if x == "weeklyexpectedrefilled" else 0)
    df2nd95Summary['CurrentWeekLosses'] = df2nd95Summary['CurrentWeekLosses'].apply(lambda x: 1 if x == "CurrentWeekLosses" else 0)
    df2nd95Summary['CurrentYearLosses'] = df2nd95Summary['CurrentYearLosses'].apply(lambda x: 1 if x == "CurrentYearLosses" else 0)
    df2nd95Summary['CapturedBiometrics'] = df2nd95Summary['CapturedBiometrics'].apply(lambda x: 1 if x == "CapturedBiometrics" else 0)
    df2nd95Summary['BiometricsGap'] = df2nd95Summary['BiometricsGap'].apply(lambda x: 1 if x == "BiometricsGap" else 0)
    df2nd95Summary['WKMissedBiometrics'] = df2nd95Summary['WKMissedBiometrics'].apply(lambda x: 1 if x == "WKMissedBiometrics" else 0)

    result = df2nd95Summary.groupby(['LGA','FacilityName'])[['ActiveClients', 'currentYearIIT', 'previousyearIIT','ImminentIIT','sevendaysIIT','currentmonthexpected', 'currentweekexpected', 'weeklyexpectedrefilled', 'CurrentWeekIIT', 'CurrentWeekLosses', 'CurrentYearLosses', 'CapturedBiometrics','BiometricsGap','WKMissedBiometrics']].sum().reset_index()
    result_check = ['ActiveClients', 'currentYearIIT', 'previousyearIIT', 'ImminentIIT', 'sevendaysIIT', 'currentmonthexpected', 'currentweekexpected', 'weeklyexpectedrefilled', 'CurrentWeekIIT', 'CurrentWeekLosses', 'CurrentYearLosses', 'CapturedBiometrics','BiometricsGap','WKMissedBiometrics']
    result = result[(result[result_check] != 0).any(axis=1)]
    
    result['%Biometrics Coverage'] = ((result['CapturedBiometrics'] / result['ActiveClients'])).round(4)
    result['Weekly_Refill_Rate'] = ((result['weeklyexpectedrefilled'] / result['currentweekexpected'])).round(4)
    result = result[['LGA','FacilityName','ActiveClients','currentmonthexpected','currentweekexpected','weeklyexpectedrefilled','Weekly_Refill_Rate', 'currentYearIIT', 'previousyearIIT', 'ImminentIIT', 'sevendaysIIT', 'CurrentWeekIIT', 'CurrentWeekLosses', 'CurrentYearLosses', 'CapturedBiometrics','%Biometrics Coverage','BiometricsGap','WKMissedBiometrics']]


    result = result.rename(columns={'ActiveClients': 'Active Clients',
                                    'currentYearIIT': 'current FY IIT',
                                    'previousyearIIT': 'previous FY IIT',
                                    'ImminentIIT': 'Imminent IIT',
                                    'sevendaysIIT': 'Next 7 days IIT',
                                    'currentmonthexpected': 'EXPECTED THIS MONTH',
                                    'currentweekexpected': 'EXPECTED THIS WEEK',
                                    'weeklyexpectedrefilled': 'REFILLED FROM WEEKLY EXPECTED',
                                    'CurrentWeekIIT': 'Curr Week IIT',
                                    'CurrentWeekLosses': 'Curr WK Losses (Stp, Dead, TO)',
                                    'CurrentYearLosses': 'Curr FY Losses (Stp, Dead, TO)',
                                    'Weekly_Refill_Rate': '%Weekly Refill Rate',
                                    'CapturedBiometrics': 'Captured Biometrics',
                                    'BiometricsGap': 'Biometrics Gap',
                                    'WKMissedBiometrics': 'Weekly Missed Biometrics',})
    result   
    
    #format and export
    formatted_period = endDate.strftime("%d-%m-%Y")

    # List of dataframes and their corresponding sheet names
    dataframes = {
        "CURRENT FY IIT": dfCurrentYearIIT,
        "PREVIOUS FY IIT": dfpreviousyearIIT,
        "IMMINENT IIT": dfImminentIIT,
        "NEXT 7 DAYS IIT": dfsevendaysIIT,
        "EXPECTED THIS MONTH": dfcurrentmonthexpected,
        "EXPECTED THIS WEEK": dfpendingweeklyrefill,
        "CURRENT FY LOSSES": dfCurrentYearLosses,
        "BIOMETRICS GAP": dfBiometricsGap,
        "WEEKLY MISSED BIOMETRICS": dfWKMissedBiometrics,
        "2ND 95 SUMMARY": result,
        # Add more dataframes and sheet names as needed
    }
    
    # Build row_masks for all sheets except the summary
    row_masks_dict = {}
    for sheet_name, df in dataframes.items():
        if sheet_name != "2ND 95 SUMMARY":
            # Compute mask for this dataframe
            # Replace sc_gap_mask(df, endDate) with whatever logic applies per sheet
            row_masks_dict[sheet_name] = sc_gap_mask(df, endDate, age_col='Current_Age')
    
    # Dynamic numeric and percentage column configuration
    column_config = {
        "2ND 95 SUMMARY": {
            "numeric_cols": [2,3,4,5,7,8,9,10,11,12,13,14,16,17],

            "percent_formulas": {
                "%Weekly Refill Rate": "=F{subtotal_row}/E{subtotal_row}",
                "%Biometrics Coverage": "=O{subtotal_row}/C{subtotal_row}"
            },

            # NEW (merge A:C)
            "merge_columns": (0, 1),

            # NEW title override
            "title": "2ND 95 PERFORMANCE REPORT – {period}"
        }
    }

    # Write each dataframe to a different sheet
    filename = export_to_excel_with_formatting(
        dataframes,
        formatted_period,
        summaryName="2ND 95 SUMMARY",
        division_columns={"G": ("F", "E"), 'P': ('O', 'C')},
        color_column=["%Weekly Refill Rate","%Biometrics Coverage"],
        column_widths={'A:A': 20, 'B:B': 35,},
        mergeNum=1, #merge first three columns for the summary total row
        row_masks=row_masks_dict,  # << pass mask here
        column_config=column_config
    )
    return filename  


def Second95RCMG(df, dfbaseline, endDate): 
    
    endDate = pd.to_datetime(endDate)
            
    print(dfbaseline)
            
    #add case managers and next appt to the dataframe
    df['CaseManager'] = df['CaseManager'].fillna('UNASSIGNED')
    
    df = integrate_baseline_data(df, dfbaseline)
    df = compute_appointment_and_iit_dates(df)
    df = classify_iit_Appt_status(df, endDate) #adding relevant columns for IIT and appointment status
    df = trackBiometrics(df, endDate) #adding relevant columns for biometrics tracking
    
    #df.to_excel('temp.xlsx', index=False)
    df['DOB'] = df['DOB'].astype(str).str.strip()
    df['DOB'] = pd.to_datetime(df['DOB'], errors='coerce', infer_datetime_format=True).fillna(pd.to_datetime('1900'))
    
    df['Current_Age'] = calculate_age_vectorized(df, 'DOB', ref_date=endDate)
    
    #apply function to process line list
    dfCurrentYearIIT = process_Linelist(df, 'CurrentYearIIT', 'CurrentYearIIT', columns_to_select, sort_by='CaseManager')
    dfpreviousyearIIT = process_Linelist(df, 'previousyearIIT', 'previousyearIIT', columns_to_select, sort_by='CaseManager')
    dfImminentIIT = process_Linelist(df, 'ImminentIIT', 'ImminentIIT', columns_to_select, sort_by='CaseManager')
    dfsevendaysIIT = process_Linelist(df, 'sevendaysIIT', 'sevendaysIIT', columns_to_select, sort_by='CaseManager')
    dfcurrentmonthexpected = process_Linelist(df, 'currentmonthexpected', 'currentmonthexpected', columns_to_select, sort_by='CaseManager')
    dfpendingweeklyrefill = process_Linelist(df, 'pendingweeklyrefill', 'pendingweeklyrefill', columns_to_select, sort_by='CaseManager') 
    dfCurrentYearLosses = process_Linelist(df, 'CurrentYearLosses', 'CurrentYearLosses', columns_to_select, sort_by='CaseManager')
    dfBiometricsGap = process_Linelist(df, 'BiometricsGap', 'BiometricsGap', columns_to_select, sort_by='CaseManager')
    dfWKMissedBiometrics = process_Linelist(df, 'WKMissedBiometrics', 'WKMissedBiometrics', columns_to_select, sort_by='CaseManager')
    
    df2nd95Summary = df

    df2nd95Summary['ActiveClients'] = df2nd95Summary['CurrentARTStatus'].apply(lambda x: 1 if x == "Active" else 0)
    df2nd95Summary['currentYearIIT'] = df2nd95Summary['CurrentYearIIT'].apply(lambda x: 1 if x == "CurrentYearIIT" else 0)
    df2nd95Summary['CurrentWeekIIT'] = df2nd95Summary['CurrentWeekIIT'].apply(lambda x: 1 if x == "CurrentWeekIIT" else 0)
    df2nd95Summary['previousyearIIT'] = df2nd95Summary['previousyearIIT'].apply(lambda x: 1 if x == "previousyearIIT" else 0)
    df2nd95Summary['ImminentIIT'] = df2nd95Summary['ImminentIIT'].apply(lambda x: 1 if x == "ImminentIIT" else 0)
    df2nd95Summary['sevendaysIIT'] = df2nd95Summary['sevendaysIIT'].apply(lambda x: 1 if x == "sevendaysIIT" else 0)
    df2nd95Summary['currentmonthexpected'] = df2nd95Summary['currentmonthexpected'].apply(lambda x: 1 if x == "currentmonthexpected" else 0)
    df2nd95Summary['currentweekexpected'] = df2nd95Summary['currentweekexpected'].apply(lambda x: 1 if x == "currentweekexpected" else 0)
    df2nd95Summary['weeklyexpectedrefilled'] = df2nd95Summary['weeklyexpectedrefilled'].apply(lambda x: 1 if x == "weeklyexpectedrefilled" else 0)
    df2nd95Summary['CurrentWeekLosses'] = df2nd95Summary['CurrentWeekLosses'].apply(lambda x: 1 if x == "CurrentWeekLosses" else 0)
    df2nd95Summary['CurrentYearLosses'] = df2nd95Summary['CurrentYearLosses'].apply(lambda x: 1 if x == "CurrentYearLosses" else 0)
    df2nd95Summary['CapturedBiometrics'] = df2nd95Summary['CapturedBiometrics'].apply(lambda x: 1 if x == "CapturedBiometrics" else 0)
    df2nd95Summary['BiometricsGap'] = df2nd95Summary['BiometricsGap'].apply(lambda x: 1 if x == "BiometricsGap" else 0)
    df2nd95Summary['WKMissedBiometrics'] = df2nd95Summary['WKMissedBiometrics'].apply(lambda x: 1 if x == "WKMissedBiometrics" else 0)

    result = df2nd95Summary.groupby(['LGA','FacilityName','CaseManager'])[['ActiveClients', 'currentYearIIT', 'previousyearIIT','ImminentIIT','sevendaysIIT','currentmonthexpected', 'currentweekexpected', 'weeklyexpectedrefilled', 'CurrentWeekIIT', 'CurrentWeekLosses', 'CurrentYearLosses', 'CapturedBiometrics','BiometricsGap','WKMissedBiometrics']].sum().reset_index()
    result_check = ['ActiveClients', 'currentYearIIT', 'previousyearIIT', 'ImminentIIT', 'sevendaysIIT', 'currentmonthexpected', 'currentweekexpected', 'weeklyexpectedrefilled', 'CurrentWeekIIT', 'CurrentWeekLosses', 'CurrentYearLosses', 'CapturedBiometrics','BiometricsGap','WKMissedBiometrics']
    result = result[(result[result_check] != 0).any(axis=1)]
    
    result['%Biometrics Coverage'] = ((result['CapturedBiometrics'] / result['ActiveClients'])).round(4)    
    result['Weekly_Refill_Rate'] = ((result['weeklyexpectedrefilled'] / result['currentweekexpected'])).round(4)
    result = result[['LGA','FacilityName','CaseManager','ActiveClients','currentmonthexpected','currentweekexpected','weeklyexpectedrefilled','Weekly_Refill_Rate', 'currentYearIIT', 'previousyearIIT', 'ImminentIIT', 'sevendaysIIT', 'CurrentWeekIIT', 'CurrentWeekLosses', 'CurrentYearLosses', 'CapturedBiometrics','%Biometrics Coverage','BiometricsGap','WKMissedBiometrics']]


    result = result.rename(columns={'ActiveClients': 'Active Clients',
                                    'currentYearIIT': 'current FY IIT',
                                    'previousyearIIT': 'previous FY IIT',
                                    'ImminentIIT': 'Imminent IIT',
                                    'sevendaysIIT': 'Next 7 days IIT',
                                    'currentmonthexpected': 'EXPECTED THIS MONTH',
                                    'currentweekexpected': 'EXPECTED THIS WEEK',
                                    'weeklyexpectedrefilled': 'REFILLED FROM WEEKLY EXPECTED',
                                    'CurrentWeekIIT': 'Curr Week IIT',
                                    'CurrentWeekLosses': 'Curr WK Losses (Stp, Dead, TO)',
                                    'CurrentYearLosses': 'Curr FY Losses (Stp, Dead, TO)',
                                    'Weekly_Refill_Rate': '%Weekly Refill Rate',
                                    'CapturedBiometrics': 'Captured Biometrics',
                                    'BiometricsGap': 'Biometrics Gap',
                                    'WKMissedBiometrics': 'Weekly Missed Biometrics',})
    result   
    
    #format and export
    formatted_period = endDate.strftime("%d-%m-%Y")

    # List of dataframes and their corresponding sheet names
    dataframes = {
        "CURRENT FY IIT": dfCurrentYearIIT,
        "PREVIOUS FY IIT": dfpreviousyearIIT,
        "IMMINENT IIT": dfImminentIIT,
        "NEXT 7 DAYS IIT": dfsevendaysIIT,
        "EXPECTED THIS MONTH": dfcurrentmonthexpected,
        "EXPECTED THIS WEEK": dfpendingweeklyrefill,
        "CURRENT FY LOSSES": dfCurrentYearLosses,
        "BIOMETRICS GAP": dfBiometricsGap,
        "WEEKLY MISSED BIOMETRICS": dfWKMissedBiometrics,
        "2ND 95 SUMMARY": result,
        # Add more dataframes and sheet names as needed
    }
    
    # Build row_masks for all sheets except the summary
    row_masks_dict = {}
    for sheet_name, df in dataframes.items():
        if sheet_name != "2ND 95 SUMMARY":
            # Compute mask for this dataframe
            # Replace sc_gap_mask(df, endDate) with whatever logic applies per sheet
            row_masks_dict[sheet_name] = sc_gap_mask(df, endDate, age_col='Current_Age')
    
     # Dynamic numeric and percentage column configuration
    column_config = {
        "2ND 95 SUMMARY": {
            "numeric_cols": [3,4,5,6,8,9,10,11,12,13,14,15,17,18],

            "percent_formulas": {
                "%Weekly Refill Rate": "=G{subtotal_row}/F{subtotal_row}",
                "%Biometrics Coverage": "=P{subtotal_row}/D{subtotal_row}"
            },

            # NEW (merge A:C)
            "merge_columns": (0, 2),

            # NEW title override
            "title": "2ND 95 PERFORMANCE REPORT – {period}"
        }
    }

    # Write each dataframe to a different sheet
    filename = export_to_excel_with_formatting(
        dataframes,
        formatted_period,
        summaryName="2ND 95 SUMMARY",
        division_columns={"H": ("G", "F"), 'Q': ('P', 'D')},
        color_column=["%Weekly Refill Rate","%Biometrics Coverage"],
        column_widths={'A:A': 20, 'B:B': 35, 'C:C': 35,},
        mergeNum=2, #merge first three columns for the summary total row
        row_masks=row_masks_dict,  # << pass mask here
        column_config=column_config
    )
    return filename