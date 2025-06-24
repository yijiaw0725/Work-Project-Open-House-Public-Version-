# %%
import pandas as pd
import seaborn as sns
# from rapidfuzz import process, fuzz
import re
# %%
# Step1: Extract only the street name
# Extract thing from we first see digit, until the first ,
# So everything before and after street name should not be there
def extract_street_name(job):
    if pd.isna(job):
        return None
    # Use regex to remove everything before the first digit and stop at the first comma
    match = re.search(r'\d.*?(?=,|$)', job)
    return match.group(0).strip() if match else job.strip()

# Step 2: Standardize street names in both sheet using a dictionary of replacements
def standardize_street_names_regex(name):
    if pd.isna(name):
        return None
    replacements = {
        r"\bDr\b\.?": "Drive",
        r"\bLn\b\.?": "Lane",
        r"\bSt\b\.?": "Street",
        r"\bAve\b\.?": "Avenue",
        r"\bBlvd\b\.?": "Boulevard",
        r"\bCt\b\.?": "Court",
        r"\bRd\b\.?": "Road",
        r"\bPl\b\.?": "Place",
    }
    for short, full in replacements.items():
        name = pd.Series([name]).str.replace(short, full, regex=True).iloc[0]
    return name.lower().strip()

# %%
################################## STEP 1: REPLACE AND LOAD ALL SALESFORCE MOVE IN DATA ##################################
salesforce_path = r"C:\Users\Yijia Wang\Desktop\Open-House-Analysis\Data\May_2025\raw_data\All Move Ins May 2025.xlsx"
salesforce = pd.read_excel(salesforce_path)
# %%
salesforce['Street Name'] = salesforce['Address'].apply(extract_street_name)
salesforce['Standardized Street Name'] = salesforce['Street Name'].apply(standardize_street_names_regex)
# %%
################################### STEP 2: REPLACE AND LOAD ALL POS FOR ATL FROM BT ###########################################
purchase_orders_path_atl_all = r"C:\Users\Yijia Wang\Desktop\Open-House-Analysis\Data\May_2025\raw_data\PurchaseOrders_May_2025_atl.xlsx"
purchase_orders_atl_all = pd.read_excel(purchase_orders_path_atl_all, skiprows=1)
# %%
purchase_orders_atl_all = purchase_orders_atl_all.dropna(subset=['Cost'])
purchase_orders_atl_all['Street Name'] = purchase_orders_atl_all['Job'].apply(extract_street_name)
purchase_orders_atl_all = purchase_orders_atl_all.dropna(subset=['Title'])
purchase_orders_atl_all['Job'] = purchase_orders_atl_all['Job'].str.strip().str.lower()

# %%
# Step 2: Initialize a DataFrame to store results
final_results_atl = pd.DataFrame(columns=purchase_orders_atl_all.columns)

# Step 3: Loop through unique Job names
for job_name in purchase_orders_atl_all['Job'].unique():
    # Filter rows for the current job name
    job_rows = purchase_orders_atl_all[purchase_orders_atl_all['Job'] == job_name]
    
    # Sort rows by "Created Date" from earliest to latest
    job_rows = job_rows.sort_values(by='Created Date')

    # Search for "touch" in the "Title" column (case-insensitive)
    touch_rows = job_rows[job_rows['Title'].str.contains(r'touch', case=False, na=False)]
    
    if not touch_rows.empty:
        # Get the date of the first "touch" match
        first_touch_date = touch_rows.iloc[0]['Created Date']
        
        # Return all rows below the "touch" row (Created Date later)
        later_rows = job_rows[
            (job_rows['Created Date'] > first_touch_date) &  # Strictly later
            (~job_rows.index.isin(touch_rows.index))  # Exclude the actual "touch" row
            ].copy()
        
        if not later_rows.empty:
            # Separate positive and negative costs
            positive_rows = later_rows[later_rows['Cost'] > 0].copy()
            negative_rows = later_rows[later_rows['Cost'] < 0].copy()
            
            # Create a count dictionary for the absolute values of negative costs
            from collections import Counter
            negative_counts = Counter(negative_rows['Cost'].abs().tolist())
            
            # Iterate through each positive row and adjust if a match is found
            adjusted_positive_rows = []
            for _, pos_row in positive_rows.iterrows():
                pos_cost = pos_row['Cost']
                # If there's a matching negative cost, set this positive cost to 0
                if negative_counts[pos_cost] > 0:
                    negative_counts[pos_cost] -= 1
                    pos_row['Cost'] = 0  # set cost to zero instead of removing the row
                # Add the (possibly adjusted) positive row to the final list
                adjusted_positive_rows.append(pos_row)
            
            # Convert the adjusted positives back to a DataFrame
            adjusted_positive_df = pd.DataFrame(adjusted_positive_rows, columns=positive_rows.columns)
            
            # Since we are not including negative rows in the final results,
            # we only append the adjusted positive rows.
            final_results_atl = pd.concat([final_results_atl, adjusted_positive_df], ignore_index=True)

# final_results_atl now contains the rows after offsetting negative costs by zeroing out the matching positive costs.

# %%
final_results_atl['Standardized Street Name'] = final_results_atl['Street Name'].apply(standardize_street_names_regex)

# Format Created Date to MM/YYYY
final_results_atl['Created Date'] = pd.to_datetime(final_results_atl['Created Date'], errors='coerce')
final_results_atl['Month/Year'] = final_results_atl['Created Date'].dt.strftime('%m/%Y')

# Perform the merge with standardized street names
merged_data_standardized_atl = final_results_atl.merge(
    salesforce[['Standardized Street Name', 'Cleaned Name', 'Area Picklist']],
    left_on='Standardized Street Name',
    right_on='Standardized Street Name',
    how='inner'
)
names_ATL = ['Clifford Senter',  'Jason Bishop', 'Jimmy Knox', 'Kirsten Davis', 'Nicholas Beuoy', 'Nicole Quiles', 'Paris Paggett', 'Ryan Worrell', 'Shaamar Moore', 'Trevor Stevens']
merged_data_standardized_atl = merged_data_standardized_atl[merged_data_standardized_atl['Cleaned Name'].isin(names_ATL)]

# %%
# Export the data for further use
export_columns = [
    "Title",
    "Street Name", 
    "Standardized Street Name", 
    "Created Date", 
    "Month/Year", 
    "Cleaned Name", 
    "Area Picklist", 
    "Cost"
]
export_data_atl = merged_data_standardized_atl[export_columns]

# %%
export_data_atl['Month/Year'] = pd.to_datetime(export_data_atl['Month/Year'], format='%m/%Y', errors='coerce')
########################################## STEP 3: CHANGE THE MONTH HERE FOR ATL #############################################
export_data_atl_Jan = export_data_atl[export_data_atl['Month/Year'].isin(['2025-01', '2025-02', '2025-03','2025-04','2025-05'])]

# %%
### EXPORT FOR ONLY ATL (no need to run unless needed) ###
# path = r"C:\Users\Yijia Wang\Desktop\Open-House-Analysis\Data\Feb_2025\archive\WO_Cost_Jan2025_2_25_atl.xlsx"
# export_data_atl_Jan.to_excel(path, index=False)

# %%
########################################### STEP 4: REPLACE AND LOAD ALL POS FRO TX FROM BT ########################################
purchase_orders_path_tx_all = r"C:\Users\Yijia Wang\Desktop\Open-House-Analysis\Data\May_2025\raw_data\PurchaseOrders_May_2025_tx.xlsx"
purchase_orders_tx_all = pd.read_excel(purchase_orders_path_tx_all, skiprows=1)

# %%
purchase_orders_tx_all['Street Name'] = purchase_orders_tx_all['Job'].apply(extract_street_name)
purchase_orders_atl_all['Job'] = purchase_orders_atl_all['Job'].str.strip().str.lower()

# %%
# Step 2: Initialize a DataFrame to store results
# Updated the loop to not only for each Job, show rows created after first see touch
# but also remove all negatives and also set positive value with same amount of negative cost found to zero
final_results_tx = pd.DataFrame(columns=purchase_orders_tx_all.columns)

# Step 3: Loop through unique Job names
for job_name_tx in purchase_orders_tx_all['Job'].unique():
    # Filter rows for the current job name
    job_rows_tx = purchase_orders_tx_all[purchase_orders_tx_all['Job'] == job_name_tx]
    
    # Sort rows by "Created Date" from earliest to latest
    job_rows_tx = job_rows_tx.sort_values(by='Created Date')
    
    # Search for "touch" in the "Title" column (case-insensitive)
    touch_rows_tx = job_rows_tx[job_rows_tx['Title'].str.contains(r'touch', case=False, na=False)]
    
    if not touch_rows_tx.empty:
        # Get the date of the first "touch" match
        first_touch_date_tx = touch_rows_tx.iloc[0]['Created Date']
        
        # Return all rows below the "touch" row (Created Date later)
        # Return all rows below the "touch" row (Created Date later)
        later_rows_tx = job_rows_tx[
            (job_rows_tx['Created Date'] > first_touch_date_tx) &  # Strictly later
            (~job_rows_tx.index.isin(touch_rows_tx.index))  # Exclude the actual "touch" row
            ].copy()
        
        if not later_rows_tx.empty:
            # Separate into positive and negative rows
            positive_rows = later_rows_tx[later_rows_tx['Cost'] > 0].copy()
            negative_rows = later_rows_tx[later_rows_tx['Cost'] < 0].copy()
            
            # Get a count of the absolute values of negative costs
            from collections import Counter
            negative_counts = Counter(negative_rows['Cost'].abs().tolist())

            # Keep track of final positive rows after adjustment
            adjusted_positive_rows = []
            
            # Iterate through each positive row to see if it can be offset by a negative
            for _, pos_row in positive_rows.iterrows():
                pos_cost = pos_row['Cost']
                # If there's a matching negative cost available, set the positive cost to 0
                if negative_counts[pos_cost] > 0:
                    negative_counts[pos_cost] -= 1
                    pos_row['Cost'] = 0
                # Add the (possibly adjusted) positive row to the final list
                adjusted_positive_rows.append(pos_row)
            
            # Convert the adjusted positives back to a DataFrame
            adjusted_positive_df = pd.DataFrame(adjusted_positive_rows, columns=positive_rows.columns)
            
            # Append these rows to the final results
            final_results_tx = pd.concat([final_results_tx, adjusted_positive_df], ignore_index=True)

# %%
final_results_tx['Standardized Street Name'] = final_results_tx['Street Name'].apply(standardize_street_names_regex)

# Format Created Date to MM/YYYY
final_results_tx['Created Date'] = pd.to_datetime(final_results_tx['Created Date'], errors='coerce')
final_results_tx['Month/Year'] = final_results_tx['Created Date'].dt.strftime('%m/%Y')

# Perform the merge with standardized street names
merged_data_standardized_tx = final_results_tx.merge(
    salesforce[['Standardized Street Name', 'Cleaned Name', 'Area Picklist']],
    left_on='Standardized Street Name',
    right_on='Standardized Street Name',
    how='inner'
)

# %%
merged_data_standardized_dfw= merged_data_standardized_tx[merged_data_standardized_tx['Area Picklist'] == 'DFW']
names_DFW = ['Chase Wilson',  'Christopher Poujol','Christopher Silbaugh', 'Gilbert Sifuentes', 'Ricardo Martinez', 'Michael Woodson', 'Oscar Flores', 'William Goodson', 'William MacQueenette', 'Damon Nash']
merged_data_standardized_dfw = merged_data_standardized_dfw[merged_data_standardized_dfw['Cleaned Name'].isin(names_DFW)]

# %%
# Export the data for further use
export_columns = [
    "Title",
    "Street Name", 
    "Standardized Street Name", 
    "Created Date", 
    "Month/Year", 
    "Cleaned Name", 
    "Area Picklist", 
    "Cost"
]
export_data_dfw = merged_data_standardized_dfw[export_columns]
# %%
# Ensure 'Month/Year' column is in datetime format
export_data_dfw['Month/Year'] = pd.to_datetime(export_data_dfw['Month/Year'], format='%m/%Y', errors='coerce')
################################################### STEP 5: CHANGE THE MONTH HERE FOR DFW #############################################
export_data_dfw_Jan = export_data_dfw[export_data_dfw['Month/Year'].isin(['2025-01', '2025-02', '2025-03','2025-04','2025-05'])]

### EXPORT FOR ONLY DFW (no need to run unless needed) ####
# path = r"C:\Users\Yijia Wang\Desktop\Open-House-Analysis\Data\Feb_2025\archive\WO_Cost_Jan2025_2_25_dfw.xlsx"
# export_data_dfw_Jan.to_excel(path, index=False)

# %%
merged_data_standardized_hou= merged_data_standardized_tx[merged_data_standardized_tx['Area Picklist'] == 'Houston']
names_HOU = ['Angel Rosas', 'Tony Chavez', 'Bryant Johnson', 'Bryce Porter', 'Kenin Vargas', 'Kenneth Lee', 'Steve Wentz']
merged_data_standardized_hou = merged_data_standardized_hou[merged_data_standardized_hou['Cleaned Name'].isin(names_HOU)]

# %%
export_data_hou = merged_data_standardized_hou[export_columns]

# %%
# Ensure 'Month/Year' column is in datetime format
export_data_hou['Month/Year'] = pd.to_datetime(export_data_hou['Month/Year'], format='%m/%Y', errors='coerce')
################################################### STEP 6: CHANGE THE MONTH HERE FOR HOU ##############################################
export_data_hou_Jan = export_data_hou[export_data_hou['Month/Year'].isin(['2025-01', '2025-02', '2025-03','2025-04','2025-05'])]

### EXPORT FOR ONLY HOU (no need to run unless needed) ###
# path = r"C:\Users\Yijia Wang\Desktop\Open-House-Analysis\Data\Feb_2025\archive\WO_Cost_Jan2025_2_25_hou.xlsx"
# export_data_hou_Jan.to_excel(path, index=False)

# %%
appended_table = pd.concat([export_data_atl_Jan, export_data_dfw_Jan, export_data_hou_Jan], ignore_index=True)
appended_table['Month/Year'] = pd.to_datetime(appended_table['Month/Year'],  format="%m/%Y")

# %%
############################ STEP 7: CHANGE NAME AND EXPORT FOR FINAL MOVE IN ISSUES ##############################
combined_path = r"C:\Users\Yijia Wang\Desktop\Open-House-Analysis\Summary\monthly_move_in_issues\May_2025\test_WO_Cost_Q22025.xlsx"
appended_table.to_excel(combined_path, index=False)