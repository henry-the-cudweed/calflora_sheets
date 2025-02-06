import pandas as pd

import requests
import json
import pandas as pd
#import shapely.geometry as s
#import folium
from IPython.display import display

#https://api.calflora.org/docs#/Observations/getMatchingObservations
#
#
#
#                      sample_api_code
#
#
#

url = 'https://api.calflora.org/observations'
api_key = 'key'
headers = {
    'Accept': 'application/json',
    'X-Api-Key': api_key
}


shapeID = your_shapeID
Cal_IPC_list_ID = "px7"
projectIds = your_projectIds
groupIds = your_groupIds

# Construct the request with specified fields
params = {
    'csetId': 291,
    #'groupIds': 
    #'dateAfter': 
    'shapeId': shapeID,
    'plantlistId': Cal_IPC_list_ID,
    #'projectIds': 

    }

response = requests.get(url, headers=headers, params=params)

if response.status_code == 200:
    data = response.json()
    
    # Print the total number of records received
    total_records = len(data)
    print(f"Total number of records received: {total_records}")
    
    if total_records > 0:
        # Convert data to a pandas DataFrame
        df = pd.DataFrame(data)
        
    else:
        print("No records received.")
else:
    print(f"Error: {response.status_code}")
    print(response.text)    

print(df)

#
#
#
#                      summary from api
#
#
#



patches_df = df 

# Convert areas to numeric values
patches_df['Infested Area'] = pd.to_numeric(patches_df['Infested Area'].str.replace(' Square Meters', ''), errors='coerce')
patches_df['Gross Area'] = pd.to_numeric(patches_df['Gross Area'].str.replace(' Square Meters', ''), errors='coerce')

# Initialize new columns
patches_df['Total Area'] = patches_df['Infested Area'].fillna(patches_df['Gross Area']).fillna(1)  # prioritize infested, then gross, then 1 sq. meter
patches_df['Gross Area Used'] = patches_df['Infested Area'].isna() & patches_df['Gross Area'].notna()  # Track where Gross Area was used
patches_df['No Area Available'] = patches_df['Infested Area'].isna() & patches_df['Gross Area'].isna()  # Track where 1 sq. meter was used by default

# Aggregate the total area per Taxon and Common Name
summary = patches_df.groupby(['Taxon', 'Common Name']).agg(
    Total_Area=('Total Area', 'sum'),
    Gross_Area_Used_Count=('Gross Area Used', 'sum'),
    No_Area_Available_Count=('No Area Available', 'sum'),
    Gross_Area_Ids=('ID', lambda x: ', '.join(x[df.loc[x.index, 'Gross Area Used']])),
    No_Area_Available_Ids=('ID', lambda x: ', '.join(x[patches_df.loc[x.index, 'No Area Available']])),
    Gross_Area_Diff=('Gross Area', lambda x: x[patches_df.loc[x.index, 'Gross Area Used']].sum()),
    No_Area_Data_Diff=('ID', lambda x: (patches_df.loc[x.index, 'No Area Available'].sum() * 1))  # 1 sq meter for each no area data
).reset_index()

# Total questionable area difference
summary['Total_Diff'] = summary['Gross_Area_Diff'] + summary['No_Area_Data_Diff']
summary['Total_Area_Uncertainty'] = ((summary['Total_Diff'] / summary['Total_Area']) * 100).round(2)

#print(summary)

# Export to Excel
#summary.to_excel('summary.xlsx', index=False)

#
#
#
#    creating patches df from api
#
#
#
#


# Create DataFrames
patches_df = df

summaries_df = summary




#
#
#
#book summary
#
#
#



# File paths (going up one directory)

# Create a new Excel writer with XlsxWriter as the engine
with pd.ExcelWriter('CF_summary_book.xlsx', engine='xlsxwriter') as writer:
    # Write the summary data to the first sheet
    summaries_df.to_excel(writer, sheet_name='Summary', index=False)

    # Access the XlsxWriter workbook and worksheet
    workbook = writer.book
    summary_sheet = writer.sheets['Summary']
    
    # Loop through each species in the summary and add hyperlinks to the sheets
    for i, species in enumerate(summaries_df['Taxon']):
        # Create a sheet for each species
        species_data = patches_df[patches_df['Taxon'] == species]
        
        # Truncate species name if it exceeds 31 characters
        sheet_name = species[:31]  # Truncate to 31 characters

        # Add the hyperlink column
        species_data['Calflora Link'] = species_data['ID'].apply(lambda x: f"https://www.calflora.org/entry/poe.html#vrid={x}")

        # Reorder columns: Insert 'Calflora Link' as the second column
        cols = list(species_data.columns)
        cols.insert(1, cols.pop(cols.index('Calflora Link')))  # Move 'Calflora Link' to second position
        species_data = species_data[cols]  # Reorder DataFrame

        species_data.to_excel(writer, sheet_name=sheet_name, index=False)
        
        # Add a hyperlink from the summary sheet to the respective species sheet
        link = f"'{sheet_name}'!A1"  # Link to cell A1 of the respective species sheet
        summary_sheet.write_url(i + 1, 0, f'internal:{link}', string=species)

    # Save the workbook (handled automatically by 'with' context)
