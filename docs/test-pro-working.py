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
    content = re.sub(r'\[comment\]: <> \(.*?\)', '', content, flags=re.IGNORECASE)  # Remove inline comments, case-insensitive
    content = re.sub(r'<!--.*?-->', '', content, flags=re.DOTALL)  # Remove HTML comments
    return content

def parse_markdown(file_path):
    with open(file_path, 'r', encoding='utf-8') as file:
        content = file.read()
    
    # Remove comments before processing
    content = remove_comments(content)
    
    parsed_data = {
        'Title': '',
        'Document Owner': [],
        'Document Owner Id': [],
        'Service Status': '',
        'ADR Authors': [],
        'ADR Latest Approved Date': '',
        'Capability Mapping Hierarchy': [],
        'Data Classification': {}
    }
    
    title_match = re.search(r'^---\ntitle:\s*(.*?)\n---', content, re.DOTALL)
    if title_match:
        parsed_data['Title'] = title_match.group(1).strip()

    document_owner_match = re.search(r'## Document Owner\s*\n([\s\S]*?)(?=##|$)', content)
    if document_owner_match:
        owners = document_owner_match.group(1).strip().split('\n')
        for owner in owners:
            owner = owner.lstrip('-').strip()
            match = re.match(r'(.+?)\s<(.+?)>', owner)
            if match:
                parsed_data['Document Owner'].append(match.group(1))
                parsed_data['Document Owner Id'].append(match.group(2))
    
    authors_match = re.search(r'## Author/Contributors\s*\n([\s\S]*?)(?=##|$)', content)
    if authors_match:
        authors = authors_match.group(1).strip().split('\n')
        for author in authors:
            author = author.lstrip('-').strip()
            parsed_data['ADR Authors'].append(author)
    
    service_status_match = re.search(r'## Service Status\s*\n\|Service Status \|\s*\n\|----\|(?:\n\|----\|)?\s*\n\|([^\|]+)\|', content)
    if service_status_match:
        parsed_data['Service Status'] = service_status_match.group(1).strip()
    
    doc_status_match = re.search(r'## Document Status\s*\n\|Document Status\|Forum\| Date \|\s*\n\|:--\|:--\|:--\|\s*\n\|([^\|]+)\|([^\|]+)\|([^\|]+)\|', content)
    if doc_status_match:
        parsed_data['ADR Latest Approved Date'] = doc_status_match.group(3).strip()
    
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
    
    data_classification_section = re.search(r'### 2\.2 Data Classification\s*\n\|Data Classification \| Risk Ratings\|\s*\n\|:--\|:--\|\s*\n([\s\S]*?)(?=###|$)', content)
    if data_classification_section:
        data_classification_rows = re.findall(r'\|([^\|]+)\|([^\|]+)\|', data_classification_section.group(1))
        for row in data_classification_rows:
            classification = row[0].strip()
            risk_rating = row[1].strip()
            parsed_data['Data Classification'][f"DC-{classification}"] = risk_rating
    
    data = {
        'Title': [parsed_data['Title']],
        'Document Owner': [', '.join(parsed_data['Document Owner'])],
        'Document Owner Id': [', '.join(parsed_data['Document Owner Id'])],
        'Service Status': [parsed_data['Service Status']],
        'ADR Authors': ['; '.join(parsed_data['ADR Authors'])],
        'ADR Latest Approved Date': [parsed_data['ADR Latest Approved Date']],
    }
    
    df = pd.DataFrame(data)
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
    classification_df = pd.DataFrame([parsed_data['Data Classification']])
    if not classification_df.empty:
        df = pd.concat([df, classification_df], axis=1)
    
    # Add dummy date value if "ADR Latest Approved Date" or "Re-certify Due Date" is missing
    if not parsed_data['ADR Latest Approved Date']:
        parsed_data['ADR Latest Approved Date'] = '00-00-0000'
    
    if 'Re-certify Due Date' not in df.columns or df['Re-certify Due Date'].isnull().any():
        df['Re-certify Due Date'] = '00-00-0000'
    
    # Check for empty columns and fill with "Check with CPA team"
    df = df.applymap(lambda x: 'Check with CPA team' if pd.isna(x) or x == '' else x)
    
    return df

def main():
    # Specify the markdown file path
    file_path = 'sample_markdown_file.md'  # Update with the actual file path
    
    # Check if file exists
    if not os.path.exists(file_path):
        print(f"Error: The file {file_path} does not exist.")
        return
    
    try:
        # Parse the markdown file and get the DataFrame
        df = parse_markdown(file_path)
        
        # Show the resulting dataframe (or save it to a file)
        print(df)
        
        # Optionally, save to CSV
        df.to_csv('parsed_output.csv', index=False)
        print("Output saved to 'parsed_output.csv'")
    
    except Exception as e:
        print(f"Error while processing the markdown file: {e}")

# Entry point of the program
if __name__ == "__main__":
    main()
