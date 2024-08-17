import os
import pandas as pd
import yaml
import re
import json

def extract_json_body(input_str):
    """Extract the JSON body from a string, removing any 'JSON Body' prefix and returning the content within `{}`."""
    json_match = re.search(r'JSON Body\s*\{(.*)\}', input_str, re.DOTALL)
    if json_match:
        json_str = "{" + json_match.group(1).strip() + "}"
        # Attempt to load and validate the JSON structure
        try:
            return json.loads(json_str)
        except json.JSONDecodeError as e:
            print(f"JSON decode error: {e}")
            return None
    return None

def create_swagger_for_sheet(sheet_name, df):
    # Define the base structure of the Swagger YAML
    swagger_base = {
        'openapi': '3.0.0',
        'info': {
            'title': f'SuperApp {sheet_name} API',
            'description': f'API documentation for {sheet_name} services in the SuperApp.',
            'version': '1.0.0'
        },
        'servers': [
            {'url': 'https://emse1-cai.dm1-emse.informaticacloud.com/caic/api/rest/v1'}
        ],
        'paths': {}
    }

    # Iterate through each row in the sheet to create paths
    for index, row in df.iterrows():
        api_name = row['API Management']
        url = row['Unnamed: 2']
        method = row['Unnamed: 3'].lower()  # Method in lowercase (get, post, etc.)
        expected_input = row['Unnamed: 4']
        test_result = row['Unnamed: 7']  # Test Result (Pass / Fail)

        # Define the structure for the endpoint
        endpoint = {
            method: {
                'summary': api_name,
                'description': row['Unnamed: 1'],  # Business Logic
                'responses': {
                    '200': {
                        'description': 'Success',
                        'content': {
                            'application/json': {
                                'schema': {
                                    'type': 'object',
                                    'example': test_result
                                }
                            }
                        }
                    },
                    '400': {'description': 'Bad request - Invalid input'},
                    '401': {'description': 'Unauthorized - Invalid credentials or token'}
                }
            }
        }

        # Handle "Token" as a header parameter
        if 'Token' in expected_input:
            endpoint[method]['parameters'] = [
                {
                    'in': 'header',
                    'name': 'Authorization',
                    'description': 'Bearer token',
                    'required': True,
                    'schema': {
                        'type': 'string',
                        'example': 'Bearer eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9'
                    }
                }
            ]

        # Handle JSON Body if present
        json_body = extract_json_body(expected_input)
        if json_body:
            endpoint[method]['requestBody'] = {
                'description': 'Request payload',
                'required': True,
                'content': {
                    'application/json': {
                        'schema': {
                            'type': 'object',
                            'example': json_body
                        }
                    }
                }
            }

        # Add the endpoint to paths
        swagger_base['paths'][url] = endpoint

    return swagger_base

def main():
    # Load the spreadsheet
    file_path = '/Users/ammar/Documents/MacFolder /Projescts/Super App/integration-doc-auto/App/integration-doc-auto/src/SuperAppQAProcessFlowV4.xlsx'
    xls = pd.ExcelFile(file_path)

    # Create directory to save YAML files
    output_dir = 'OpenApi_Swagger'
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # Iterate through each sheet and generate a Swagger YAML file
    for sheet_name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet_name)
        swagger_yaml = create_swagger_for_sheet(sheet_name, df)

        # Save the YAML file
        yaml_file_path = os.path.join(output_dir, f'{sheet_name}.yaml')
        with open(yaml_file_path, 'w') as yaml_file:
            yaml.dump(swagger_yaml, yaml_file, sort_keys=False)

        print(f'Created Swagger YAML for {sheet_name} at {yaml_file_path}')

if __name__ == "__main__":
    main()