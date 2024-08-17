import os
import pandas as pd
import yaml
import re

def clean_key(key):
    return re.sub(r'^\d+\.\s*', '', key).strip('"\'')

def clean_value(value):
    if not isinstance(value, str):
        return value
    return value.strip('"\'')

def infer_type(value):
    if isinstance(value, (int, float)):
        return "number" if isinstance(value, float) else "integer"
    
    if isinstance(value, str):
        if value.lower() in ["true", "false"]:
            return "boolean"
        try:
            int(value)
            return "integer"
        except ValueError:
            pass

        try:
            float(value)
            return "number"
        except ValueError:
            pass

    return "string"

def parse_input_parameters(expected_input, expected_input_sample):
    headers = {}
    body_content = {}

    input_lines = expected_input.split('\n')
    sample_lines = expected_input_sample.split('\n')
    
    for inp, samp in zip(input_lines, sample_lines):
        key = clean_key(inp.strip())
        value = clean_value(samp.strip())
        
        if "token" in key.lower():
            headers['Authorization'] = f"Bearer {value}"
        elif "login" in key.lower() or "username" in key.lower():
            body_content['login'] = value
        elif "password" in key.lower():
            body_content['password'] = value
        elif "json body" not in key.lower():
            body_content[key] = value
    
    return headers, body_content

def create_openapi_structure(excel_file, output_folder, server_url):
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    xls = pd.ExcelFile(excel_file)

    for sheet_name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet_name)

        title = f'SuperApp {sheet_name} API'
        description = f'API documentation for {sheet_name} services in the SuperApp.'
        version = '1.0.0'

        openapi = {
            'openapi': '3.0.0',
            'info': {
                'title': title,
                'description': description,
                'version': version
            },
            'servers': [{'url': server_url}],
            'paths': {}
        }

        for _, row in df.iterrows():
            base_url = str(row[df.columns[2]]).strip()
            if '://' in base_url:
                base_url = base_url.split('/', 3)[-1]
            if not base_url.startswith('/'):
                base_url = '/' + base_url

            http_method = str(row[df.columns[3]]).lower().strip()
            summary = str(row[df.columns[0]]).strip()
            business_logic = str(row[df.columns[1]]).strip()
            expected_input = str(row[df.columns[4]]).strip()
            expected_input_sample = str(row[df.columns[5]]).strip()
            expected_output = str(row[df.columns[6]]).strip()

            headers, body_content = parse_input_parameters(expected_input, expected_input_sample)

            request_body = None
            if http_method in ['post', 'put'] and body_content:
                request_body = {
                    'description': 'User credentials for authentication',
                    'required': True,
                    'content': {
                        'application/json': {
                            'schema': {
                                'type': 'object',
                                'properties': {
                                    key: {
                                        'type': infer_type(value),
                                        'example': value
                                    } for key, value in body_content.items()
                                }
                            }
                        }
                    }
                }

            response_properties = {}
            try:
                response_dict = yaml.safe_load(expected_output)
                for key, value in response_dict.items():
                    response_properties[clean_key(key)] = {
                        'type': infer_type(value),
                        'example': clean_value(value)
                    }
            except Exception as e:
                print(f"Error parsing expected output: {e}")

            responses = {
                '200': {
                    'description': 'Successful operation',
                    'content': {
                        'application/json': {
                            'schema': {
                                'type': 'object',
                                'properties': response_properties
                            }
                        }
                    }
                },
                '401': {
                    'description': 'Unauthorized - Invalid credentials'
                }
            }

            if base_url not in openapi['paths']:
                openapi['paths'][base_url] = {}

            method_details = {
                'summary': summary,
                'description': business_logic,
                'responses': responses
            }

            if request_body:
                method_details['requestBody'] = request_body

            if headers:
                method_details['parameters'] = [
                    {
                        'name': key,
                        'in': 'header',
                        'required': True,
                        'schema': {"type": "string", "example": value},
                        'description': key
                    } for key, value in headers.items()
                ]

            openapi['paths'][base_url][http_method] = method_details

        output_yaml_file = os.path.join(output_folder, f'{sheet_name.replace(" ", "_")}_swagger.yaml')
        with open(output_yaml_file, 'w') as yaml_file:
            yaml.dump(openapi, yaml_file, default_flow_style=False, sort_keys=False)

if __name__ == '__main__':
    excel_file = '/Users/ammar/Documents/MacFolder /Projescts/Super App/integration-doc-auto/App/integration-doc-auto/src/SuperAppQAProcessFlowV4.xlsx'
    output_folder = '/Users/ammar/Documents/MacFolder /Projescts/Super App/integration-doc-auto/App/integration-doc-auto/src/output'
    server_url = 'https://emse1-cai.dm1-emse.informaticacloud.com/caic/api/rest/v1'
    create_openapi_structure(excel_file, output_folder, server_url)
