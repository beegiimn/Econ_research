import pandas as pd
import jupyter as pn
import numpy as np
import matplotlib.pyplot as plt
import eventstudy as es
sheet1 = pd.read_excel('C:/Users/beegi/OneDrive/Spring_courses/Capital_flow/Eventstudy_0312.xlsx', sheet_name='rating_m')
sheet2 = pd.read_excel('C:/Users/beegi/OneDrive/Spring_courses/Capital_flow/Eventstudy_0312.xlsx', sheet_name='equity')
# Finding difference between subsequent values in sheet1
sheet1diff = sheet1.diff()
# creating list consists of only positive numbers after finding 
# diff in sheet1 
upgrade = sheet1.loc[sheet1diff['Brazil'] > 0] 
Brazil_up = pd.DataFrame({'Count': [len(upgrade)]})
# creating list consists of only negative numbers after finding 
# diff in sheet1 
downgrade = sheet1.loc[sheet1diff['Brazil'] < 0] 
Brazil_down = pd.DataFrame({'Count': [len(downgrade)]})

################
# Get unique dates from 'downgrade'
dates_to_match = upgrade['Date'].unique()
# Create a boolean mask for matching dates in 'sheet2'
match_mask = sheet2['Date'].isin(dates_to_match)
# Filter 'sheet2' using the boolean mask
matched_rows = sheet2.loc[match_mask, ['Date', 'Brazil']]
# Get the index of the first matching row in matched_rows
matching_row_idx = matched_rows.index[0]
# Create empty lists to store the data
prev_dates = []
prev_Brazil = []
matched_dates = []
matched_Brazil = []
next_dates = []
next_Brazil = []

# Define the window size
window_size = 11
# Loop through the rows of matched_rows and extract the corresponding rows from sheet2
for idx, row in matched_rows.iterrows():
    matching_date = row['Date']
    # Find the matching row in sheet2 and add the previous and next rows
    matching_row_idx = sheet2.loc[sheet2['Date'] == matching_date].index[0]
    prev_rows = sheet2.loc[matching_row_idx-window_size : matching_row_idx-1, ['Date', 'Brazil']]
    matched_rows = sheet2.loc[matching_row_idx : matching_row_idx, ['Date', 'Brazil']]
    next_rows = sheet2.loc[matching_row_idx+1 : matching_row_idx+window_size, ['Date', 'Brazil']]
    # Append the columns to the corresponding lists
    prev_dates.append(prev_rows['Date'].reset_index(drop=True))
    prev_Brazil.append(prev_rows['Brazil'].reset_index(drop=True))
                
    matched_dates.append(matched_rows['Date'].reset_index(drop=True))
    matched_Brazil.append(matched_rows['Brazil'].reset_index(drop=True))
                
    next_dates.append(next_rows['Date'].reset_index(drop=True))
    next_Brazil.append(next_rows['Brazil'].reset_index(drop=True))
                
# Concatenate the lists into dataframes
prev_df = pd.concat(prev_dates + prev_Brazil, axis=1)
matched_df = pd.concat(matched_dates + matched_Brazil, axis=1)
next_df = pd.concat(next_dates + next_Brazil, axis=1)
Brazil = pd.concat([prev_df, matched_df, next_df], axis=0, ignore_index=True)
Brazil

with pd.ExcelWriter('C:/Users/beegi/OneDrive/Spring_courses/Capital_flow/countries_upgrade_equity.xlsx', engine='openpyxl', mode='a') as writer:
    Brazil.to_excel(writer, sheet_name='Brazil', index=False)


    
# concatenate the two dataframes by column
merged_df = pd.concat([Ukraine, Ukraine, Thailand, SriLanka, SouthAfrica, Slovenia, Slovakia, Romania, Poland, Philippines, Pakistan, Mongolia, Mexico, Malaysia, Lithuania, Lebanon, Latvia, Korea, India, Hungary, Estonia, Czech, China, Chile, Bulgaria, Ukraine], axis=1)
# write the merged dataframe to an Excel file
# create a list of all column names that contain the string "Date"
date_columns = [col for col in merged_df.columns if 'Date' in col]
# drop all the columns containing the string "Date"
merged_df.drop(columns=date_columns, inplace=True)
merged_df
 

with pd.ExcelWriter('C:/Users/beegi/OneDrive/Spring_courses/Capital_flow/Eventstudy_0312.xlsx', engine='openpyxl', mode='a') as writer:
    merged_df.to_excel(writer, sheet_name='merged_debt_downgrade', index=False)
