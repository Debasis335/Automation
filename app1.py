import pandas as pd
import os
from langchain_core.prompts import ChatPromptTemplate
from langchain_groq import ChatGroq
from dotenv import load_dotenv

load_dotenv() 
# Load the data from the Excel files
current_day_df = pd.read_excel(r'D:\Automation\doc\Current_day_by_Technician.xlsx')
audit_df = pd.read_excel(r'D:\Automation\doc\Service_Desk_Incident_Management_Ticket_Audits.xlsx')

# Clean up any leading/trailing spaces in column names
current_day_df.columns = current_day_df.columns.str.strip()
audit_df.columns = audit_df.columns.str.strip()

# Ensure that your Groq API Key is set as an environment variable
api_key = os.getenv("GROQ_API_KEY")
if not api_key:
    raise ValueError("API key for Groq is missing. Please set the GROQ_API_KEY environment variable.")

# Initialize ChatGroq with the Groq model you want to use
chat = ChatGroq(temperature=0, model_name="llama-3.3-70b-versatile")

# Function to check subject format with Groq API
def check_subject_format_with_groq(subject):
    system = "You are a helpful assistant."
    human = f"Check if the following subject follows the naming convention 'SERVICE - Brief Description of Issue or Request': {subject}. Reply with 'Valid Format' if it follows the convention, or 'Invalid Format' if it does not."

    prompt = ChatPromptTemplate.from_messages([("system", system), ("human", human)])
    chain = prompt | chat
    
    response = chain.invoke({"text": subject})
    
    if "Valid Format" in response:
        return "No Audit"
    else:
        return "Normal Audit required"

# Function to check if the resolution mentions a solution article
def check_resolution_for_solution_article(resolution):
    if pd.notna(resolution) and ('kba' in resolution.lower() or 'solution article' in resolution.lower() or 'auto resolved' in resolution.lower()):
        return "No Audit"
    return "Normal Audit required"

# Initialize an empty list to store the updated rows for the Audit DataFrame
updated_audit_rows = []

# Loop through each row of the Current Day DataFrame
for _, current_row in current_day_df.iterrows():
    # Prepare a new row for the Audit DataFrame
    audit_row = {}
    
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

    # Section 5a ix: Has the ticket "Subject" been updated to leverage the naming convention "SERVICE - Brief Description of Issue or Request"?
    subject = current_row['Subject']
    if pd.notna(subject):
        # Check subject format using Groq model
        audit_row['Has the ticket "Subject" been updated to leverage the naming convention "SERVICE - Brief Description of Issue or Request"? (Section 5a ix)'] = check_subject_format_with_groq(subject)
    else:
        audit_row['Has the ticket "Subject" been updated to leverage the naming convention "SERVICE - Brief Description of Issue or Request"? (Section 5a ix)'] = 'Normal Audit required'

    # Section 8b: Did the technician search for and note a relevant Solution article, if one exists?
    resolution = current_row['Resolution']
    audit_row['Did the technician search for and note a relevant Solution article, if one exists? (Section 8b)'] = check_resolution_for_solution_article(resolution)

    # Section 9: Did the technician provide clear and detailed notes, documenting all steps taken during troubleshooting?
    if any(keyword in str(resolution).lower() for keyword in ['kba', 'auto resolved', 'user confirmation']):
        audit_row['Did the technician provide clear and detailed notes, documenting all steps taking during troubleshooting? (Section 9)'] = 'True'
    else:
        audit_row['Did the technician provide clear and detailed notes, documenting all steps taking during troubleshooting? (Section 9)'] = 'False'

    # Section 10a: Are the resolutions notes clearly and fully documented, including the exact steps taken for resolution?
    if pd.notna(resolution) and 'steps' in str(resolution).lower():
        audit_row['Are the resolutions notes clearly and fully documented, including the exact steps taken for resolution? (Section 10a)'] = 'Yes'
    else:
        audit_row['Are the resolutions notes clearly and fully documented, including the exact steps taken for resolution? (Section 10a)'] = 'No'

    # Section 10e: If no solution article existed, did the technician submit a new solution article request?
    if any(keyword in str(resolution).lower() for keyword in ['new kba', 'new solution article']):
        audit_row['If no solution article existed, did the technician submit a new solution article request? (Section 10e)'] = 'Yes'
    else:
        audit_row['If no solution article existed, did the technician submit a new solution article request? (Section 10e)'] = 'No'

    # Notes: Provide reasoning based on audit outcomes
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

print("Audit process with Groq completed and saved.")
