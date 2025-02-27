import pandas as pd
from datetime import datetime
from dateutil.relativedelta import relativedelta

# Function to calculate recertify dates
def calculate_rec_cert_dates(approval_date):
    try:
        approval_date = datetime.strptime(approval_date, '%d-%m-%Y')
        recertify_due_date = approval_date + relativedelta(months=10)
        recertify_due_date_str = recertify_due_date.strftime('%d-%m-%Y')
        recertify_due_month = recertify_due_date.strftime('%b')
        
        # Check if upcoming recertification is within the next 3 months
        today = datetime.today()
        three_months_later = today + relativedelta(months=3)
        upcoming_recertification = 1 if today <= recertify_due_date <= three_months_later else 0
        
        return recertify_due_date_str, recertify_due_month, upcoming_recertification
    except Exception as e:
        print(f"Error while calculating recertification dates: {e}")
        return '01-01-2000', 'NA', 0

# Sample sadr_data dictionary
sadr_data = {
    'Service1': {'S-ADR Service Name': 'Service1', 'S-ADR Service Status': 'Active', 'S-ADR Approval Date': '15-05-2022', 'S-ADR Document Status': 'Approved'},
    'Service2': {'S-ADR Service Name': 'Service2', 'S-ADR Service Status': 'Inactive', 'S-ADR Approval Date': '20-07-2021', 'S-ADR Document Status': 'Not Approved'}
}

# Sample final_df DataFrame
final_df = pd.DataFrame({
    'Service Name': ['Service1', 'Service3'],
    'Service Owner': ['Owner1', 'Owner2'],
    'Service Owner Id': ['ID1', 'ID2'],
    'Service Status': ['Active', 'Inactive'],
    'ADR Authors': ['Author1', 'Author2'],
    'ADR Document Status': ['Approved', 'Pending'],
    'Latest Approval date': ['15-05-2022', '01-01-2020'],
    'S-ADR Document Status': ['NA', 'NA'],
    'S-ADR Service Status': ['NA', 'NA'],
    'S-ADR Approval Date': ['01-01-2000', '01-01-2000'],
    'Re-certify Due Date': ['15-05-2022', '01-01-2021'],
    'Re-certify Due Month': ['May', 'Jan'],
    'Upcoming Recertification': [1, 0]
})

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
        sadr_service_status = sadr_info.get('S-ADR Service Status', '')
        sadr_approval_date = sadr_info.get('S-ADR Approval Date', '01-01-2000')
        
        recertify_due_date, recertify_due_month, upcoming_recertification = calculate_rec_cert_dates(sadr_approval_date)
        
        # Creating a new row with default values for missing columns
        new_row = {
            'Service Name': sadr_info['S-ADR Service Name'],
            'Service Owner': 'NA',
            'Service Owner Id': 'NA',
            'Service Status': sadr_service_status,
            'ADR Authors': 'NA',
            'ADR Document Status': 'No f-adr',
            'Latest Approval date': '01-01-2000',
            'S-ADR Document Status': sadr_info.get('S-ADR Document Status', 'NA'),
            'S-ADR Service Status': sadr_service_status,
            'S-ADR Approval Date': sadr_approval_date,
            'Re-certify Due Date': recertify_due_date,
            'Re-certify Due Month': recertify_due_month,
            'Upcoming Recertification': upcoming_recertification
        }
        
        # Append the new row to the final dataframe
        final_df = final_df.append(new_row, ignore_index=True)

# Output the final DataFrame
print(final_df)
