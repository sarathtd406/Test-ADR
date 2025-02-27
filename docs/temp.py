import os
import re
import pandas as pd
from datetime import datetime
from dateutil.relativedelta import relativedelta

def remove_comments(content):
    """
    Removes markdown comments of the format:
    [comment]: <> (This is a comment)
    <!-- This is a comment -->
    """
    content = re.sub(r'\[comment\]: <> \(.*?\)', '', content, flags=re.IGNORECASE)  # Remove inline comments
    content = re.sub(r'<!--.*?-->', '', content, flags=re.DOTALL)  # Remove HTML comments
    return content

def parse_markdown(file_path, file_type='foundational'):
    """
    Parse markdown files based on the file type ('foundational', 'deprecated', or 'service').
    """
    with open(file_path, 'r', encoding='utf-8') as file:
        content = file.read()

    # Remove comments before processing
    content = remove_comments(content)
    
    parsed_data = {
        'Service Name': '',
        'Service Owner': [],
        'Service Owner Id': [],
        'Service Status': '',
        'ADR Authors': [],
        'ADR Document Status': '',
        'Latest Approval date': '',
        'Capability Mapping Hierarchy': [],
        'Data Classification': {},
        'S-ADR Service Name': '',
        'S-ADR Document Status': '',
        'S-ADR Service Status': '',
        'S-ADR Approval Date': ''
    }

    if file_type == 'foundational' or file_type == 'deprecated':
        # Extract title for both foundational and deprecated ADRs
        title_match = re.search(r'^---\ntitle:\s*(.*?)\n---', content, re.DOTALL)
        if title_match:
            parsed_data['Service Name'] = title_match.group(1).strip()
        
        # Other sections extraction (same logic as before)
        # For brevity, skipping the rest of the 'foundational' logic (similar to your previous code)
        # You can reuse your previous function for foundational/adr file parsing here.

    elif file_type == 'service':
        # For service-adr files, extract title and Document Status section
        title_match = re.search(r'^---\ntitle:\s*(.*?)\n---', content, re.DOTALL)
        if title_match:
            parsed_data['S-ADR Service Name'] = title_match.group(1).strip()

        # Extract Document Status Table
        doc_status_match = re.search(r'## Document Status\s*\n\| Document Status \| Service Status \| Forum \| Approval Date \|\s*\n\|:--\|:--\|:--\|:--\|\s*\n\|([^\|]+)\|([^\|]+)\|([^\|]+)\|([^\|]+)\|', content)
        if doc_status_match:
            parsed_data['S-ADR Document Status'] = doc_status_match.group(1).strip()
            parsed_data['S-ADR Service Status'] = doc_status_match.group(2).strip()
            parsed_data['S-ADR Approval Date'] = doc_status_match.group(4).strip()

    # Return parsed data
    return parsed_data

def calculate_rec_cert_dates(approval_date_str):
    """
    Calculate the recertification due date, due month, and upcoming recertification date
    based on the S-ADR Approval Date.
    """
    try:
        approval_date = datetime.strptime(approval_date_str, "%d-%m-%Y")
        recertify_due_date = approval_date + relativedelta(years=1)
        recertify_due_month = recertify_due_date.strftime("%m-%Y")
        upcoming_recertification = recertify_due_date.strftime("%m-%Y")
    except Exception as e:
        recertify_due_date = '01-01-2000'
        recertify_due_month = '01-2000'
        upcoming_recertification = '01-2000'
    
    return recertify_due_date, recertify_due_month, upcoming_recertification

def main():
    # List of folders containing markdown files
    folder_paths = ['C:\\Users\\sarath\\TCO_APP_DEV\\foundational-adr', 
                    'C:\\Users\\sarath\\TCO_APP_DEV\\deprecated-adr', 
                    'C:\\Users\\sarath\\TCO_APP_DEV\\service-adr']  # Update with actual folder paths
    
    # Initialize an empty list to hold dataframes from each file
    all_dfs = []
    
    # Initialize an empty dictionary to store S-ADR data
    sadr_data = {}

    for folder_path in folder_paths:
        # Identify which folder to process
        if folder_path.endswith('service-adr'):
            # From service-adr, consider files that start with "s-adr-"
            markdown_files = [f for f in os.listdir(folder_path) if f.startswith('s-adr-') and f.endswith('.md')]
            file_type = 'service'  # Set file type as 'service' for S-ADR files
        elif folder_path.endswith('deprecated-adr'):
            # From deprecated-adr, consider files that start with "do-not-use-f-adr-"
            markdown_files = [f for f in os.listdir(folder_path) if f.startswith('do-not-use-f-adr-') and f.endswith('.md')]
            file_type = 'deprecated'
        else:
            # For foundational-adr folder, include all markdown files
            markdown_files = [f for f in os.listdir(folder_path) if f.endswith('.md') and f.lower() not in ['README.md', 'foundational-adr-structure.md']]
            file_type = 'foundational'

        for file_name in markdown_files:
            file_path = os.path.join(folder_path, file_name)
            
            try:
                # Parse the markdown file and get the parsed data dictionary
                parsed_data = parse_markdown(file_path, file_type=file_type)
                
                # If it's an s-adr file, store the parsed title and document status data
                if file_type == 'service':
                    sadr_data[parsed_data['S-ADR Service Name']] = {
                        'S-ADR Document Status': parsed_data['S-ADR Document Status'],
                        'S-ADR Service Status': parsed_data['S-ADR Service Status'],
                        'S-ADR Approval Date': parsed_data['S-ADR Approval Date']
                    }
                else:
                    # Process non-s-adr files (as you did previously)
                    df = pd.DataFrame([parsed_data])
                    all_dfs.append(df)
            
            except Exception as e:
                print(f"Error while processing the file {file_name}: {e}")
    
    # Concatenate all DataFrames into one
    final_df = pd.concat(all_dfs, ignore_index=True)

    # Add S-ADR data to final_df by matching the Service Name
    for index, row in final_df.iterrows():
        service_name = row['Service Name']
        if service_name in sadr_data:
            final_df.at[index, 'S-ADR Document Status'] = sadr_data[service_name].get('S-ADR Document Status', '')
            final_df.at[index, 'S-ADR Service Status'] = sadr_data[service_name].get('S-ADR Service Status', '')
            final_df.at[index, 'S-ADR Approval Date'] = sadr_data[service_name].get('S-ADR Approval Date', '')
        else:
            # Add a new row with default values
            sadr_service_name = sadr_data.get(service_name, {}).get('S-ADR Service Name', '')
            sadr_service_status = sadr_data.get(service_name, {}).get('S-ADR Service Status', '')
            sadr_approval_date = sadr_data.get(service_name, {}).get('S-ADR Approval Date', '01-01-2000')

            recertify_due_date, recertify_due_month, upcoming_recertification = calculate_rec_cert_dates(sadr_approval_date)
            
            # Creating a new row
            new_row = {
                'Service Name': sadr_service_name,
                'Service Owner': 'NA',
                'Service Owner Id': 'NA',
                'Service Status': sadr_service_status,
                'ADR Authors': 'NA',
                'ADR Document Status': 'No f-adr',
                'Latest Approval date': '01-01-2000',
                'Capability Mapping Hierarchy': 'NA',
                'Data Classification': 'NA',
                'S-ADR Document Status': sadr_data.get(service_name, {}).get('S-ADR Document Status', 'NA'),
                'S-ADR Service Status': sadr_service_status,
                'S-ADR Approval Date': sadr_approval_date,
                'Re-certify Due Date': recertify_due_date,
                'Re-certify Due Month': recertify_due_month,
                'Upcoming Recertification': upcoming_recertification
            }
            
            # Append the new row to the final dataframe
            final_df = final_df.append(new_row, ignore_index=True)

    # Modify the DataFrame as per the requirements
    final_df.reset_index(inplace=True, drop=True)
    final_df.index += 1
    final_df.index.name = "SL No."
    
    # Save the final dataframe to an Excel sheet
    final_df.to_excel('Governance_Data.xlsx', sheet_name='F-ADR', index=True)
    print("Output saved to 'Governance_Data.xlsx' with sheet name 'F-ADR'")

# Entry point of the program
if __name__ == "__main__":
    main()
