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
        # Extract title
        title_match = re.search(r'^---\ntitle:\s*(.*?)\n---', content, re.DOTALL)
        if title_match:
            parsed_data['Service Name'] = title_match.group(1).strip()

        # Extract document owner and contributors
        document_owner_match = re.search(r'## Document Owner\s*\n([\s\S]*?)(?=##|$)', content)
        if document_owner_match:
            owners = document_owner_match.group(1).strip().split('\n')
            for owner in owners:
                owner = owner.lstrip('-').strip()
                match = re.match(r'(.+?)\s<(.+?)>', owner)
                if match:
                    parsed_data['Service Owner'].append(match.group(1))
                    parsed_data['Service Owner Id'].append(match.group(2))
        
        authors_match = re.search(r'## Author/Contributors\s*\n([\s\S]*?)(?=##|$)', content)
        if authors_match:
            authors = authors_match.group(1).strip().split('\n')
            for author in authors:
                author = author.lstrip('-').strip()
                parsed_data['ADR Authors'].append(author)
        
        # Extract document status and approval date
        doc_status_match = re.search(r'## Document Status\s*\n\| Document Status \| Forum \| Date \|\s*\n\|:--\|:--\|:--\|\s*\n\|([^\|]+)\|([^\|]+)\|([^\|]+)\|', content)
        if doc_status_match:
            parsed_data['ADR Document Status'] = doc_status_match.group(1).strip()
            parsed_data['Latest Approval date'] = doc_status_match.group(3).strip()
        
        # Extract capability mapping hierarchy
        capability_mapping_section = re.search(r'## 1\. Capability Mapping Hierarchy\s*\n([\s\S]*?)(?=##|$)', content)
        if capability_mapping_section:
            capability_table_rows = re.findall(r'\|([^\|]+)\|([^\|]+)\|([^\|]+)\|', capability_mapping_section.group(1))
            capability_table_rows = capability_table_rows[1:]  # Skip the header row
            capability_table_rows = [row for row in capability_table_rows if not any(":--" in col for col in row)]
            for row in capability_table_rows:
                parsed_data['Capability Mapping Hierarchy'].append({
                    "Cap-Map Level 0": row[0].strip(),
                    "Cap-Map Level 1": row[1].strip(),
                    "Cap-Map Level 2": row[2].strip()
                })
        
        # Extract data classification
        data_classification_section = re.search(r'### 2\.2 Data Classification\s*\n\| Data Classification \| Risk Rating \|\s*\n\|:-\|:-\|\s*\n([\s\S]*?)(?=###|$)', content)
        if data_classification_section:
            data_classification_rows = re.findall(r'\|([^\|]+)\|([^\|]+)\|', data_classification_section.group(1))
            for row in data_classification_rows:
                classification = row[0].strip()
                risk_rating = row[1].strip()
                parsed_data['Data Classification'][f"DC-{classification}"] = risk_rating
        
        # Extract service status
        service_status_section = re.search(r'## Service Status\s*\n\| Service Status \|\s*\n\|---\|\s*\n\|([^\|]+)\|', content)

        if service_status_section:
            parsed_data['Service Status'] = service_status_section.group(1).strip()
    
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
        
    return parsed_data


def process_f_adr(parsed_data):

    # Prepare DataFrame from parsed data
    data = {
        'Service Name': [parsed_data['Service Name']],
        'Service Owner': [', '.join(parsed_data['Service Owner'])],
        'Service Owner Id': [', '.join(parsed_data['Service Owner Id'])],
        'Service Status': [parsed_data['Service Status']],
        'ADR Authors': ['; '.join(parsed_data['ADR Authors'])],
        'ADR Document Status': [parsed_data['ADR Document Status']],
        'Latest Approval date': [parsed_data['Latest Approval date']]
    }

    df = pd.DataFrame(data)
    capability_df = pd.DataFrame(parsed_data['Capability Mapping Hierarchy'])
    if not capability_df.empty:
        main_data = pd.DataFrame({
            'Service Name': [parsed_data['Service Name']] * len(capability_df),
            'Service Owner': [', '.join(parsed_data['Service Owner'])] * len(capability_df),
            'Service Owner Id': [', '.join(parsed_data['Service Owner Id'])] * len(capability_df),
            'Service Status': [parsed_data['Service Status']] * len(capability_df),
            'ADR Authors': ['; '.join(parsed_data['ADR Authors'])] * len(capability_df),
            'ADR Document Status': [parsed_data['ADR Document Status']] * len(capability_df),
            'Latest Approval date': [parsed_data['Latest Approval date']] * len(capability_df)
        })
        df = pd.concat([main_data, capability_df], axis=1)
    
    # Data classification
    classification_df = pd.DataFrame([parsed_data['Data Classification']])
    if not classification_df.empty:
        df = pd.concat([df, classification_df], axis=1)
    
    # Add dummy date and check for missing values
    if parsed_data['Latest Approval date']:
        try:
            if parsed_data['Latest Approval date'] in ["TBD", "dd-mm-yyyy", ""]:
                parsed_data['Latest Approval date'] = "00-00-0000"
            
            latest_approved_date = datetime.strptime(parsed_data['Latest Approval date'], '%d-%m-%Y')
            recertify_due_date = latest_approved_date + relativedelta(months=10)
            recertify_due_date_str = recertify_due_date.strftime('%d-%m-%Y')
        except Exception as e:
            print(f"Error while calculating Re-certify Due Date: {e}")
            parsed_data['Latest Approval date'] = "00-00-0000"
            recertify_due_date_str = "00-00-0000"
        
        # Update the DataFrame with dummy date if necessary
        df['Re-certify Due Date'] = recertify_due_date_str
    else:
        parsed_data['Latest Approval date'] = "00-00-0000"
        recertify_due_date_str = "00-00-0000"
        df['Re-certify Due Date'] = recertify_due_date_str
    
    # Check for empty columns and fill them with "Check with CPA team"
    df = df.applymap(lambda x: x if pd.notnull(x) and x != "" else "Check with CPA team")

    # Convert date columns to datetime format before saving
    df['Latest Approval date'] = pd.to_datetime(df['Latest Approval date'], format='%d-%m-%Y', errors='coerce')
    df['Re-certify Due Date'] = pd.to_datetime(df['Re-certify Due Date'], format='%d-%m-%Y', errors='coerce')
    
    # Replace invalid dates with dummy date
    df['Latest Approval date'] = df['Latest Approval date'].apply(lambda x: datetime(2000, 1, 1) if pd.isnull(x) or str(x) in ["Check with CPA team", "00-00-0000"] else x)
    df['Re-certify Due Date'] = df['Re-certify Due Date'].apply(lambda x: datetime(2000, 1, 1) if pd.isnull(x) or str(x) in ["Check with CPA team", "00-00-0000"] else x)
    
    # Convert back to string in required format
    df['Latest Approval date'] = df['Latest Approval date'].dt.strftime('%d-%m-%Y')
    df['Re-certify Due Date'] = df['Re-certify Due Date'].dt.strftime('%d-%m-%Y')

    # Add new columns
    df['Re-certify Due Month'] = pd.to_datetime(df['Re-certify Due Date'], format='%d-%m-%Y', errors='coerce').dt.strftime('%b')
    
    # Calculate upcoming recertifications in the next 3 months
    today = datetime.today()
    three_months_later = today + relativedelta(months=3)
    df['Upcoming Recertification'] = df['Re-certify Due Date'].apply(lambda x: 1 if today <= pd.to_datetime(x, format='%d-%m-%Y', errors='coerce') <= three_months_later else 0)

    return df

def main():
    # List of folders containing markdown files
    folder_paths = ['C:\\Users\\foundation-adr', 
                    'C:\\Users\\deprecated-adrs',
                    'C:\\Users\\service-adr']  # Update with actual folder paths
    
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
            markdown_files = [f for f in os.listdir(folder_path) if f.startswith('do-not-use-adr') and f.endswith('.md')]
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
                        'S-ADR Service Name': parsed_data['S-ADR Service Name'],
                        'S-ADR Document Status': parsed_data['S-ADR Document Status'],
                        'S-ADR Service Status': parsed_data['S-ADR Service Status'],
                        'S-ADR Approval Date': parsed_data['S-ADR Approval Date']
                    }
                else:
                    # Process non-s-adr files (as you did previously)
                    df = process_f_adr(parsed_data)
                    all_dfs.append(df)
            
            except Exception as e:
                print(f"Error while processing the file {file_name}: {e}")
    
    # Concatenate all DataFrames into one
    final_df = pd.concat(all_dfs, ignore_index=True)

    # Add S-ADR data to final_df by matching the Service Name
    for sadr_service_name, sadr_info in sadr_data.items():
        service_name = sadr_info['S-ADR Service Name']
        
        # Check if the service_name exists in final_df
        if service_name in final_df['Service Name'].values:
            # Find the row in final_df where the Service Name matches
            index = final_df[final_df['Service Name'] == service_name].index[0]
            
            # Update the S-ADR columns in final_df
            final_df.at[index, 'S-ADR Document Status'] = sadr_info.get('S-ADR Document Status', '')
            final_df.at[index, 'S-ADR Service Status'] = sadr_info.get('S-ADR Service Status', '')
            final_df.at[index, 'S-ADR Approval Date'] = sadr_info.get('S-ADR Approval Date', '')
        else:
            # If service_name does not exist, create a new row with default values
            sadr_document_status = sadr_info.get('S-ADR Document Status', '')
            sadr_service_status = sadr_info.get('S-ADR Service Status', '')
            sadr_approval_date = sadr_info.get('S-ADR Approval Date', '01-01-2000')

            # recertify_due_date, recertify_due_month, upcoming_recertification = calculate_rec_cert_dates(sadr_approval_date)
            
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
                'S-ADR Document Status': sadr_document_status,
                'S-ADR Service Status': sadr_service_status,
                'S-ADR Approval Date': sadr_approval_date,
                'Re-certify Due Date': '01-01-2000',
                'Re-certify Due Month': 'Jan',
                'Upcoming Recertification': 0
            }
            
            # Append the new row to the final dataframe
            new_row_df = pd.DataFrame([new_row])
            final_df = pd.concat([final_df, new_row_df], ignore_index=True)

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
