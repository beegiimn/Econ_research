    import pandas as pd
    import jupyter as pn
    import numpy as np
    import matplotlib.pyplot as plt
    import eventstudy as es
    sheet1 = pd.read_excel('C:/Users/beegi/OneDrive/Spring_courses/Capital_flow/Eventstudy_0312.xlsx', sheet_name='rating_m')
    sheet2 = pd.read_excel('C:/Users/beegi/OneDrive/Spring_courses/Capital_flow/Eventstudy_0312.xlsx', sheet_name='debt')
    # Finding difference between subsequent values in sheet1
    sheet1diff = sheet1.diff()
    # creating list consists of only positive numbers after finding 
    # diff in sheet1 
    upgrade = sheet1.loc[sheet1diff['Ukraine'] > 0] 
    Ukraine_up = pd.DataFrame({'Count': [len(upgrade)]})
    # creating list consists of only negative numbers after finding 
    # diff in sheet1 
    downgrade = sheet1.loc[sheet1diff['Ukraine'] < 0] 
    Ukraine_down = pd.DataFrame({'Count': [len(downgrade)]})

    ################
    # Get unique dates from 'downgrade'
    dates_to_match = downgrade['Date'].unique()
    # Create a boolean mask for matching dates in 'sheet2'
    match_mask = sheet2['Date'].isin(dates_to_match)
    # Filter 'sheet2' using the boolean mask
    matched_rows = sheet2.loc[match_mask, ['Date', 'Ukraine']]
    # Get the index of the first matching row in matched_rows
    matching_row_idx = matched_rows.index[0]
    # Create empty lists to store the data
    prev_dates = []
    prev_Ukraine = []
    matched_dates = []
    matched_Ukraine = []
    next_dates = []
    next_Ukraine = []

    # Define the window size
    window_size = 11
    # Loop through the rows of matched_rows and extract the corresponding rows from sheet2
    for idx, row in matched_rows.iterrows():
        matching_date = row['Date']
        # Find the matching row in sheet2 and add the previous and next rows
        matching_row_idx = sheet2.loc[sheet2['Date'] == matching_date].index[0]
        prev_rows = sheet2.loc[matching_row_idx-window_size : matching_row_idx-1, ['Date', 'Ukraine']]
        matched_rows = sheet2.loc[matching_row_idx : matching_row_idx, ['Date', 'Ukraine']]
        next_rows = sheet2.loc[matching_row_idx+1 : matching_row_idx+window_size, ['Date', 'Ukraine']]
        # Append the columns to the corresponding lists
        prev_dates.append(prev_rows['Date'].reset_index(drop=True))
        prev_Ukraine.append(prev_rows['Ukraine'].reset_index(drop=True))
            
        matched_dates.append(matched_rows['Date'].reset_index(drop=True))
        matched_Ukraine.append(matched_rows['Ukraine'].reset_index(drop=True))
            
        next_dates.append(next_rows['Date'].reset_index(drop=True))
        next_Ukraine.append(next_rows['Ukraine'].reset_index(drop=True))
            
    # Concatenate the lists into dataframes
    prev_df = pd.concat(prev_dates + prev_Ukraine, axis=1)
    matched_df = pd.concat(matched_dates + matched_Ukraine, axis=1)
    next_df = pd.concat(next_dates + next_Ukraine, axis=1)
    Ukraine = pd.concat([prev_df, matched_df, next_df], axis=0, ignore_index=True)
    Ukraine

with pd.ExcelWriter('C:/Users/beegi/OneDrive/Spring_courses/Capital_flow/countries_downgrade_debt.xlsx', engine='openpyxl', mode='a') as writer:
    Ukraine.to_excel(writer, sheet_name='merged', index=False)

 