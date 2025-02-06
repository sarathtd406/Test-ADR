import os
import re
import pandas as pd
from datetime import datetime
from dateutil.relativedelta import relativedelta

def parse_markdown(file_path):
    # Read markdown file
    with open(file_path, 'r') as file:
        content = file.read()
    
    # Define the sections we are interested in
    parsed_data = {
        'Title': '',
        'Document Owner': [],
        'Document Owner Id': [],
        'Service Status': '',
        'ADR Authors': [],
        'ADR Latest Approved Date': '',
        'Capability Mapping Hierarchy': [],
        'Data Classification': {}  # To store the classification columns
    }
    
    # Extract Title (typically appears after ---)
    title_match = re.search(r'^---\ntitle:\s*(.*?)\n---', content, re.DOTALL)
    if title_match:
        parsed_data['Title'] = title_match.group(1).strip()

    # Extract Document Owner and their emails (handling names with hyphens)
    document_owner_match = re.search(r'## Document Owner\s*\n([\s\S]*?)(?=##|$)', content)
    if document_owner_match:
        owners = document_owner_match.group(1).strip().split('\n')
        for owner in owners:
            owner = owner.lstrip('-').strip()
            match = re.match(r'(.+?)\s<(.+?)>', owner)
            if match:
                parsed_data['Document Owner'].append(match.group(1))
                parsed_data['Document Owner Id'].append(match.group(2))

    # Extract Authors/Contributors
    authors_match = re.search(r'## Author/Contributors\s*\n([\s\S]*?)(?=##|$)', content)
    if authors_match:
        authors = authors_match.group(1).strip().split('\n')
        for author in authors:
            author = author.lstrip('-').strip()
            parsed_data['ADR Authors'].append(author)

    # Extract Service Status
    service_status_match = re.search(r'## Service Status\s*\n\|Service Status \|\s*\n\|----\|(?:\n\|----\|)?\s*\n\|([^\|]+)\|', content)
    if service_status_match:
        parsed_data['Service Status'] = service_status_match.group(1).strip()

    # Extract Document Status and Latest Approved Date
    doc_status_match = re.search(r'## Document Status\s*\n\|Document Status\|Forum\| Date \|\s*\n\|:--\|:--\|:--\|\s*\n\|([^\|]+)\|([^\|]+)\|([^\|]+)\|', content)
    if doc_status_match:
        parsed_data['ADR Latest Approved Date'] = doc_status_match.group(3).strip()

    # Extract Capability Mapping Hierarchy Table
    capability_mapping_section = re.search(r'## 1\. Capability Mapping Hierarchy\s*\n([\s\S]*?)(?=##|$)', content)
    
    if capability_mapping_section:
        capability_table_rows = re.findall(r'\|([^\|]+)\|([^\|]+)\|([^\|]+)\|', capability_mapping_section.group(1))
        capability_table_rows = capability_table_rows[1:]
        capability_table_rows = [row for row in capability_table_rows if not any(":--" in col for col in row)]
        
        for row in capability_table_rows:
            parsed_data['Capability Mapping Hierarchy'].append({
                "Cap-Map Level 0": row[0].strip(),
                "Cap-Map Level 1": row[1].strip(),
                "Cap-Map Level 2": row[2].strip()
            })

    # Extract Data Classification Table
    data_classification_section = re.search(r'### 2\.2 Data Classification\s*\n\|Data Classification \| Risk Ratings\|\s*\n\|:--\|:--\|\s*\n([\s\S]*?)(?=###|$)', content)
    
    if data_classification_section:
        data_classification_rows = re.findall(r'\|([^\|]+)\|([^\|]+)\|', data_classification_section.group(1))
        
        for row in data_classification_rows:
            classification = row[0].strip()
            risk_rating = row[1].strip()
            parsed_data['Data Classification'][f"DC-{classification}"] = risk_rating

    # Create DataFrame with the parsed data
    data = {
        'Title': [parsed_data['Title']],
        'Document Owner': [', '.join(parsed_data['Document Owner'])],
        'Document Owner Id': [', '.join(parsed_data['Document Owner Id'])],
        'Service Status': [parsed_data['Service Status']],
        'ADR Authors': ['; '.join(parsed_data['ADR Authors'])],
        'ADR Latest Approved Date': [parsed_data['ADR Latest Approved Date']],
    }
    
    df = pd.DataFrame(data)

    # Capability Mapping Hierarchy
    capability_df = pd.DataFrame(parsed_data['Capability Mapping Hierarchy'])
    
    if not capability_df.empty:
        main_data = pd.DataFrame({
            'Title': [parsed_data['Title']] * len(capability_df),
            'Document Owner': [', '.join(parsed_data['Document Owner'])] * len(capability_df),
            'Document Owner Id': [', '.join(parsed_data['Document Owner Id'])] * len(capability_df),
            'Service Status': [parsed_data['Service Status']] * len(capability_df),
            'ADR Authors': ['; '.join(parsed_data['ADR Authors'])] * len(capability_df),
            'ADR Latest Approved Date': [parsed_data['ADR Latest Approved Date']] * len(capability_df)
        })
        
        df = pd.concat([main_data, capability_df], axis=1)

    # Data Classification
    classification_df = pd.DataFrame([parsed_data['Data Classification']])

    if not classification_df.empty:
        df = pd.concat([df, classification_df], axis=1)

    # Calculate Re-certify Due Date
    if parsed_data['ADR Latest Approved Date']:
        try:
            latest_approved_date = datetime.strptime(parsed_data['ADR Latest Approved Date'], '%Y-%m-%d')
            recertify_due_date = latest_approved_date + relativedelta(months=10)
            df['Re-certify Due Date'] = recertify_due_date.strftime('%Y-%m-%d')
        except Exception as e:
            print(f"Error while calculating Re-certify Due Date: {e}")

    return df

def parse_all_markdown_files(folder_path):
    # Initialize an empty dataframe to collect all results
    final_df = pd.DataFrame()
    
    # List all files in the directory
    for file_name in os.listdir(folder_path):
        if file_name.endswith('.md'):  # Only consider markdown files
            file_path = os.path.join(folder_path, file_name)
            # Parse each markdown file and get the dataframe
            df = parse_markdown(file_path)
            # Append the result to the final dataframe
            final_df = pd.concat([final_df, df], ignore_index=True)
    
    # Reset the index and adjust it to start from 1
    final_df.reset_index(inplace=True, drop=True)
    final_df.index += 1
    final_df.index.name = "SL No."

    return final_df

def save_to_excel(df, output_path):
    # Save the dataframe to an Excel file
    with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=True, sheet_name='ADR Data')

# Path to the folder containing markdown files
folder_path = 'C:\\path\\to\\foundation-adr'  # Update with your actual folder path

# Path to the output Excel file
output_file_path = 'C:\\path\\to\\output\\final_output.xlsx'  # Update with your desired output path

# Parse all markdown files and get the final dataframe
final_df = parse_all_markdown_files(folder_path)

# Save the dataframe to an Excel file
save_to_excel(final_df, output_file_path)

print(f"Data successfully saved to {output_file_path}")
