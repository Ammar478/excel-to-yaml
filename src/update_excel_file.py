import os
import pandas as pd
import json

# Load the Excel file
file_path = '/Users/ammar/Documents/MacFolder /Projescts/Super App/integration-doc-auto/App/integration-doc-auto/src/SuperAppQAProcessFlowV4.xlsx'
xls = pd.ExcelFile(file_path)

# Create a dictionary to hold processed data
processed_data = {}

# Function to check and return JSON from string
def extract_json(body_string):
    try:
        # Convert the string into a JSON object to verify its structure
        json_object = json.loads(body_string)
        return json.dumps(json_object, indent=2)  # Return formatted JSON string
    except ValueError:
        # If it fails to parse, return the original string
        return body_string

# Iterate through each sheet
for sheet_name in xls.sheet_names:
    df = pd.read_excel(xls, sheet_name=sheet_name)
    
    # Debug: Print the first few rows to inspect the data
    print(f"Processing sheet: {sheet_name}")
    print(df.head())

    # Initialize the new columns
    df['HTTP Header'] = ''
    df['Body'] = ''

    # Check if 'Token' and 'JSON BODY' exist, if not, initialize empty columns
    if 'Token' not in df.columns:
        print(f"'Token' column not found in sheet {sheet_name}. Adding it.")
        df['Token'] = ''
    if 'JSON BODY' not in df.columns:
        print(f"'JSON BODY' column not found in sheet {sheet_name}. Adding it.")
        df['JSON BODY'] = ''

    # Iterate through each row
    for index, row in df.iterrows():
        # Process 'Token' column
        if pd.notna(row['Token']):
            df.at[index, 'HTTP Header'] = row['Token']
            print(f"Row {index}: Set HTTP Header to {row['Token']}")  # Debug

        # Process 'JSON BODY' column
        if pd.notna(row['JSON BODY']):
            body_data = row['JSON BODY']
            df.at[index, 'Body'] = extract_json(body_data)
            print(f"Row {index}: Set Body to {df.at[index, 'Body']}")  # Debug

    # Remove the original 'JSON BODY' column
    if 'JSON BODY' in df.columns:
        df.drop(columns=['JSON BODY'], inplace=True)

    # Store the processed dataframe in the dictionary
    processed_data[sheet_name] = df

# Define the output path
output_dir = '/Users/ammar/Documents/MacFolder /Projescts/Super App/integration-doc-auto/App/integration-doc-auto/src/data'
output_path = os.path.join(output_dir, 'Processed_SuperApp_QA_Process_Flow.xlsx')

# Create the directory if it doesn't exist
os.makedirs(output_dir, exist_ok=True)

# Save the processed data to a new Excel file
with pd.ExcelWriter(output_path) as writer:
    for sheet_name, df in processed_data.items():
        # Ensure that the sheet remains visible
        df.to_excel(writer, sheet_name=sheet_name, index=False)

print(f"Processing complete. The modified file is saved as {output_path}.")