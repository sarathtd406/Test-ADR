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
            # Remove hyphen and space at the beginning of the line if it's a bullet point
            owner = owner.lstrip('-').strip()
            match = re.match(r'(.+?)\s<(.+?)>', owner)
            if match:
                parsed_data['Document Owner'].append(match.group(1))
                parsed_data['Document Owner Id'].append(match.group(2))

    # Extract Authors/Contributors (handling names with hyphens)
    authors_match = re.search(r'## Author/Contributors\s*\n([\s\S]*?)(?=##|$)', content)
    if authors_match:
        authors = authors_match.group(1).strip().split('\n')
        for author in authors:
            # Remove hyphen and space at the beginning of the line if it's a bullet point
            author = author.lstrip('-').strip()
            parsed_data['ADR Authors'].append(author)

    # Extract Service Status (single column table)
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
        # Now extract only the table rows within this section, ensuring that we exclude the header and separator rows
        capability_table_rows = re.findall(r'\|([^\|]+)\|([^\|]+)\|([^\|]+)\|', capability_mapping_section.group(1))

        # Skip the first row (header row) and filter out rows with the alignment markers (i.e., containing only ':--')
        capability_table_rows = capability_table_rows[1:]  # Skip the first row (headers)
        capability_table_rows = [row for row in capability_table_rows if not any(":--" in col for col in row)]
        
        # Process each row and store them as key-value pairs
        for row in capability_table_rows:
            parsed_data['Capability Mapping Hierarchy'].append({
                "Cap-Map Level 0": row[0].strip(),
                "Cap-Map Level 1": row[1].strip(),
                "Cap-Map Level 2": row[2].strip()
            })

    # Extract Data Classification Table (Data Classification section)
    data_classification_section = re.search(r'### 2\.2 Data Classification\s*\n\|Data Classification \| Risk Ratings\|\s*\n\|:--\|:--\|\s*\n([\s\S]*?)(?=###|$)', content)
    
    if data_classification_section:
        # Extract the Data Classification table rows
        data_classification_rows = re.findall(r'\|([^\|]+)\|([^\|]+)\|', data_classification_section.group(1))
        
        # Process each row, using the first column as the new column name and the second column as the value
        for row in data_classification_rows:
            classification = row[0].strip()  # Data Classification value
            risk_rating = row[1].strip()    # Risk Ratings value
            
            # Add the classification as a new column in the final dictionary with its risk rating value
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
    
    # Create the main DataFrame
    df = pd.DataFrame(data)

    # Now include Capability Mapping Hierarchy in the DataFrame as separate columns
    capability_df = pd.DataFrame(parsed_data['Capability Mapping Hierarchy'])
    
    # Check if Capability Mapping Hierarchy has content
    if not capability_df.empty:
        # Repeat the main data for each row of capability data
        main_data = pd.DataFrame({
            'Title': [parsed_data['Title']] * len(capability_df),
            'Document Owner': [', '.join(parsed_data['Document Owner'])] * len(capability_df),
            'Document Owner Id': [', '.join(parsed_data['Document Owner Id'])] * len(capability_df),
            'Service Status': [parsed_data['Service Status']] * len(capability_df),
            'ADR Authors': ['; '.join(parsed_data['ADR Authors'])] * len(capability_df),
            'ADR Latest Approved Date': [parsed_data['ADR Latest Approved Date']] * len(capability_df)
        })
        
        # Concatenate the main data with capability data
        df = pd.concat([main_data, capability_df], axis=1)

    # Now add the Data Classification columns to the DataFrame
    classification_df = pd.DataFrame([parsed_data['Data Classification']])

    # Concatenate the main data with classification data
    if not classification_df.empty:
        df = pd.concat([df, classification_df], axis=1)

    # Add a new column "Re-certify Due Date" which is 10 months ahead of the "ADR Latest Approved Date"
    if parsed_data['ADR Latest Approved Date']:
        try:
            # Parse the date with the correct format
            latest_approved_date = datetime.strptime(parsed_data['ADR Latest Approved Date'], '%d-%m-%Y')  # Change to DD-MM-YYYY format
            # Calculate the Re-certify Due Date (10 months ahead)
            recertify_due_date = latest_approved_date + relativedelta(months=10)
            # Add the new column to the DataFrame
            df['Re-certify Due Date'] = recertify_due_date.strftime('%d-%m-%Y')  # Return in the same format
        except Exception as e:
            print(f"Error while calculating Re-certify Due Date: {e}")

    return df

# Path to the markdown file
file_path = 'sample-foundation.md'

# Parse the file and get the DataFrame
df = parse_markdown(file_path)

# Adjust display settings for full output
pd.set_option('display.max_colwidth', None)  # No truncation for column values
pd.set_option('display.width', 1000)  # Expand the output width
pd.set_option('display.max_rows', None)  # Ensure all rows are displayed (if multiple)

# Print the entire DataFrame
print(df.to_string(index=False))  # Use .to_string() to show the full DataFrame without truncation
