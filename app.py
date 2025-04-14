import pandas as pd
import re

# Load the Current Day DataFrame
current_day_df = pd.read_excel(r'D:\Automation\doc\Current_day_by_Technician.xlsx')

# Load the Audit DataFrame (Service Desk Incident Management)
audit_df = pd.read_excel(r'D:\Automation\doc\Service_Desk_Incident_Management_Ticket_Audits.xlsx')

# Clean up any leading/trailing spaces in column names
current_day_df.columns = current_day_df.columns.str.strip()
audit_df.columns = audit_df.columns.str.strip()

# Define a regex pattern to match the correct format for Subject
subject_pattern = r'^[A-Z]+ - .+'

# Initialize an empty list to store the updated rows for the Audit DataFrame
updated_audit_rows = []

# Loop through each row of the Current Day DataFrame
for _, current_row in current_day_df.iterrows():
    # Prepare a new row for the Audit DataFrame
    audit_row = {}

    # Fill in the values based on the requirements
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
        requester_note = ''
    else:
        audit_row['Has the "Requester Name" been updated to reflect who the ticket is for? (Section 2b)'] = 'Normal Audit required'
        requester_note = 'Requester name missing'

    # Section 5a ix: Has the ticket "Subject" been updated to leverage the naming convention?
    subject = current_row['Subject']
    if pd.notna(subject) and re.match(subject_pattern, subject):
        audit_row['Has the ticket "Subject" been updated to leverage the naming convention "SERVICE - Brief Description of Issue or Request"? (Section 5a ix)'] = 'No Audit'
        subject_note = ''
    else:
        audit_row['Has the ticket "Subject" been updated to leverage the naming convention "SERVICE - Brief Description of Issue or Request"? (Section 5a ix)'] = 'Normal Audit required'
        subject_note = 'Subject is not in proper format'

    # Section 8b: Did the technician search for and note a relevant Solution article?
    resolution = current_row['Resolution']
    if any(keyword in str(resolution).lower() for keyword in ['kba', 'solution article', 'auto resolved']):
        audit_row['Did the technician search for and note a relevant Solution article, if one exists? (Section 8b)'] = 'No Audit'
        kba_note = ''
    else:
        audit_row['Did the technician search for and note a relevant Solution article, if one exists? (Section 8b)'] = 'Normal Audit required'
        kba_note = 'KBA is missing'

    # Section 9: Did the technician provide clear and detailed notes, documenting all steps taken during troubleshooting?
    if any(keyword in str(resolution).lower() for keyword in ['kba', 'auto resolved', 'user confirmation']):
        audit_row['Did the technician provide clear and detailed notes, documenting all steps taken during troubleshooting? (Section 9)'] = 'True'
    else:
        audit_row['Did the technician provide clear and detailed notes, documenting all steps taken during troubleshooting? (Section 9)'] = 'False'

    # Section 10a: Are the resolution notes clearly and fully documented, including the exact steps taken for resolution?
    if pd.notna(resolution) and 'steps' in str(resolution).lower():
        audit_row['Are the resolutions notes clearly and fully documented, including the exact steps taken for resolution? (Section 10a)'] = 'Yes'
        troubleshooting_note = ''
    else:
        audit_row['Are the resolutions notes clearly and fully documented, including the exact steps taken for resolution? (Section 10a)'] = 'No'
        troubleshooting_note = 'There is no proper steps for troubleshoot'

    # Section 10e: If no solution article existed, did the technician submit a new solution article request?
    if any(keyword in str(resolution).lower() for keyword in ['new kba', 'new solution article']):
        audit_row['If no solution article existed, did the technician submit a new solution article request? (Section 10e)'] = 'Yes'
        new_kba_note = ''
    else:
        audit_row['If no solution article existed, did the technician submit a new solution article request? (Section 10e)'] = 'No'
        new_kba_note = 'New KBA or New solution missing'

    # Combine all the notes
    all_notes = []
    if requester_note:
        all_notes.append(requester_note)
    if subject_note:
        all_notes.append(subject_note)
    if kba_note:
        all_notes.append(kba_note)
    if troubleshooting_note:
        all_notes.append(troubleshooting_note)
    if new_kba_note:
        all_notes.append(new_kba_note)
    
    # Join all notes with semicolon
    audit_row['Notes'] = '; '.join(all_notes) if all_notes else ''

    # Append the updated row to the list of rows for the Audit DataFrame
    updated_audit_rows.append(audit_row)

# Create a DataFrame from the updated rows
updated_audit_df = pd.DataFrame(updated_audit_rows)

# Save the updated audit DataFrame back to Excel
updated_audit_df.to_excel(r'D:\Automation\doc\Updated_Service_Desk_Incident_Management_Ticket_Audits.xlsx', index=False)

print("Audit process completed and saved to 'D:\\Automation\\doc\\Updated_Service_Desk_Incident_Management_Ticket_Audits.xlsx'")
