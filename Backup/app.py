import pandas as pd

# Load the Current Day DataFrame
current_day_df = pd.read_excel(r'D:\Automation\doc\Current_day_by_Technician.xlsx')

# Load the Audit DataFrame (Service Desk Incident Management)
audit_df = pd.read_excel(r'D:\Automation\doc\Service_Desk_Incident_Management_Ticket_Audits.xlsx')

# Clean up any leading/trailing spaces in column names
current_day_df.columns = current_day_df.columns.str.strip()
audit_df.columns = audit_df.columns.str.strip()

# Print column names for debugging
print("Current Day DataFrame Columns:", current_day_df.columns)
print("Audit DataFrame Columns:", audit_df.columns)

# Initialize an empty list to store the updated rows for the Audit DataFrame
updated_audit_rows = []

# Loop through each row of the Current Day DataFrame
for _, current_row in current_day_df.iterrows():
    # Prepare a new row for the Audit DataFrame
    audit_row = {}
    # print(current_row)
    # Fill in the values based on your conditions
    
    # Technician - Technician column of Current_day_by_Technician excel
    audit_row['Technician'] = current_row['Technician']
    
    # Request ID - RequestID of Current_day_by_Technician excel
    audit_row['Request ID'] = current_row['RequestID']
    
    # Subject - Subject of Current_day_by_Technician excel
    audit_row['Subject'] = current_row['Subject']
    
    # Completed Date - Resolved Time of Current_day_by_Technician excel
    audit_row['Completed Date'] = current_row['Resolved Time']
    
    # Manager - "Saravanan"
    audit_row['Manager'] = 'Saravanan'

    # Section 2b: Has the "Requester Name" been updated to reflect who the ticket is for?
    requester = current_row['Requester']
    if pd.notna(requester):
        audit_row['Has the "Requester Name" been updated to reflect who the ticket is for? (Section 2b)'] = 'No Audit'
    else:
        audit_row['Has the "Requester Name" been updated to reflect who the ticket is for? (Section 2b)'] = 'Normal Audit required'
    
    # Section 5a ix: Has the ticket "Subject" been updated to leverage the naming convention?
    subject = current_row['Subject']
    if pd.notna(subject):
        audit_row['Has the ticket "Subject" been updated to leverage the naming convention "SERVICE - Brief Description of Issue or Request"? (Section 5a ix)'] = 'No Audit'
    else:
        audit_row['Has the ticket "Subject" been updated to leverage the naming convention "SERVICE - Brief Description of Issue or Request"? (Section 5a ix)'] = 'Normal Audit required'
    
    # Section 8b: Did the technician search for and note a relevant Solution article?
    resolution = current_row['Resolution']
    if any(keyword in str(resolution).lower() for keyword in ['kba', 'solution article', 'auto resolved']):
        audit_row['Did the technician search for and note a relevant Solution article, if one exists? (Section 8b)'] = 'No Audit'
    else:
        audit_row['Did the technician search for and note a relevant Solution article, if one exists? (Section 8b)'] = 'Normal Audit required'

    # Section 9: Did the technician provide clear and detailed notes, documenting all steps taken during troubleshooting?
    if any(keyword in str(resolution).lower() for keyword in ['kba', 'solution article', 'auto resolved']):
        audit_row['Did the technician provide clear and detailed notes, documenting all steps taken during troubleshooting? (Section 9)'] = 'No Audit'
    else:
        audit_row['Did the technician provide clear and detailed notes, documenting all steps taken during troubleshooting? (Section 9)'] = 'Normal Audit required'

    # Section 10a: Are the resolutions notes clearly and fully documented?
    if pd.notna(resolution) and 'steps' in str(resolution).lower():
        audit_row['Are the resolutions notes clearly and fully documented, including the exact steps taken for resolution? (Section 10a)'] = 'Yes'
    else:
        audit_row['Are the resolutions notes clearly and fully documented, including the exact steps taken for resolution? (Section 10a)'] = 'No'

    # Section 10e: If no solution article existed, did the technician submit a new solution article request?
    if any(keyword in str(resolution).lower() for keyword in ['new kba', 'new solution article']):
        audit_row['If no solution article existed, did the technician submit a new solution article request? (Section 10e)'] = 'Yes'
    else:
        audit_row['If no solution article existed, did the technician submit a new solution article request? (Section 10e)'] = 'No'

    # Notes: Explain why certain audits are required
    if 'Normal Audit required' in audit_row.values():
        audit_row['Notes'] = 'Audit required due to missing information or failure to follow conventions.'
    else:
        audit_row['Notes'] = 'No issues found.'

    # Append the updated row to the list of rows for the Audit DataFrame
    updated_audit_rows.append(audit_row)

# Create a DataFrame from the updated rows
updated_audit_df = pd.DataFrame(updated_audit_rows)

# Save the updated audit DataFrame back to Excel
updated_audit_df.to_excel(r'D:\Automation\doc\Updated_Service_Desk_Incident_Management_Ticket_Audits.xlsx', index=False)

print("Audit process completed and saved to 'D:\\Automation\\doc\\Updated_Service_Desk_Incident_Management_Ticket_Audits.xlsx'")
