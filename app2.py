import pandas as pd
import os
from langchain_core.prompts import ChatPromptTemplate
from langchain_groq import ChatGroq
from dotenv import load_dotenv
from fuzzywuzzy import fuzz

# Load environment variables from .env file
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

# Function to generate valid subject header based on issue description
def generate_valid_subject_header(issue_description):
    service = "SERVICE"  # You could extract this dynamically if needed, e.g., 'CLOCK', 'CASEWARE', etc.
    valid_subject = f"{service} - {issue_description.strip()}"
    return valid_subject

# Function to compare generated subject with actual subject using Levenshtein distance (fuzzy matching)
def compare_subjects(generated_subject, actual_subject):
    similarity_score = fuzz.ratio(generated_subject.lower(), actual_subject.lower())  # Case-insensitive comparison
    return similarity_score

# Function to check subject format with Groq API using dynamic subject generation and similarity comparison
def check_subject_format_with_groq(row):
    subject = row['Subject']
    issue_description = row['Issue Description']  # Assuming 'Issue Description' is a column in the DataFrame
    generated_subject = generate_valid_subject_header(issue_description)
    similarity_score = compare_subjects(generated_subject, subject)
    
    if similarity_score >= 70:
        return "No Audit"
    else:
        return "Manual Audit required"

# Function to check subject format with Groq API using few-shot learning
def check_subject_format_with_groq1(subject):
    system = "You are a helpful assistant."
    
    few_shot_examples = """
    Example 1: 'CLOCK - Adjust and Add EST clock' --> Valid Format
    Example 2: 'CASEWARE - Unable to access the file' --> Valid Format
    Example 3: 'PASSWORD - DS1 account' --> Valid Format
    Example 4: 'SPAM - Email from aol.com' --> Valid Format
    Example 5: 'PPC - Single audit missing template update' --> Valid Format
    Example 6: 'QUICKBOOKS - Unable to map R drive' --> Valid Format
    Example 7: 'Password - No access to Work' --> Invalid Format
    Example 8: 'Email may be fake' --> Invalid Format
    Example 9: 'CCH Axcess - Not able to update' --> Invalid Format
    Example 10: 'Security - Email with "Remittance" Subject' --> Invalid Format
    Example 11: 'Immediate Action Required: Return of Berdon HP Laptops' --> Invalid Format
    """
    
    human = f"""
    Check if the following subject follows the naming convention 'SERVICE - Brief Description of Issue or Request'. 
    If the subject follows the convention, respond with 'Valid Format', otherwise respond with 'Invalid Format'.
    
    {few_shot_examples}
    
    Subject to check: {subject}
    """
    
    prompt = ChatPromptTemplate.from_messages([("system", system), ("human", human)])
    chain = prompt | chat
    
    # Invoke the Groq model
    response = chain.invoke({"text": subject})
    
    if "Valid Format" in response.content:
        return "No Audit"  # If the subject follows the correct format
    else:
        return "Normal Audit required"  # If the subject does not follow the correct format

# Function to check if the resolution mentions a solution article
def check_resolution_for_solution_article(resolution):
    if pd.notna(resolution) and ('kba' in resolution.lower() or 'solution article' in resolution.lower() or 'auto resolved' in resolution.lower()):
        return "No Audit"
    return "Normal Audit required"

# Function to generate dynamic notes based on audit status and reason
def generate_dynamic_notes(audit_status, reason):
    system = "You are a helpful assistant."
    human = f"Generate a detailed audit note for the following audit status: {audit_status}. Reason: {reason}. Provide a clear and concise explanation."
    
    prompt = ChatPromptTemplate.from_messages([("system", system), ("human", human)])
    chain = prompt | chat
    
    response = chain.invoke({"text": f"Audit Status: {audit_status}. Reason: {reason}"})
    return response.content  # Access content, not ['text']

# Initialize an empty list to store the updated rows for the Audit DataFrame
updated_audit_rows = []

# Loop through each row of the Current Day DataFrame
for _, current_row in current_day_df.iterrows():
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
    # Q-1
    requester = current_row['Requester']
    if pd.notna(requester):
        audit_row['Has the "Requester Name" been updated to reflect who the ticket is for? (Section 2b)'] = 'No Audit'
    else:
        audit_row['Has the "Requester Name" been updated to reflect who the ticket is for? (Section 2b)'] = 'Normal Audit required'

    # Section 5a ix: Has the ticket "Subject" been updated to leverage the naming convention "SERVICE - Brief Description of Issue or Request"?
    subject = current_row['Subject']
    if pd.notna(subject):
        audit_row['Has the ticket "Subject" been updated to leverage the naming convention "SERVICE - Brief Description of Issue or Request"? (Section 5a ix)'] = check_subject_format_with_groq(current_row)
    else:
        audit_row['Has the ticket "Subject" been updated to leverage the naming convention "SERVICE - Brief Description of Issue or Request"? (Section 5a ix)'] = 'Normal Audit required'

    # Section 8b: Did the technician search for and note a relevant Solution article, if one exists?
    # Q-3
    resolution = current_row['Resolution']
    audit_row['Did the technician search for and note a relevant Solution article, if one exists? (Section 8b)'] = check_resolution_for_solution_article(resolution)

    # Section 9: Did the technician provide clear and detailed notes, documenting all steps taken during troubleshooting?
    # Q-4
    if any(keyword in str(resolution).lower() for keyword in ['kba', 'auto resolved', 'user confirmation']):
        audit_row['Did the technician provide clear and detailed notes, documenting all steps taken during troubleshooting? (Section 9)'] = 'True'
    else:
        audit_row['Did the technician provide clear and detailed notes, documenting all steps taken during troubleshooting? (Section 9)'] = 'False'

    # Section 10a: Are the resolutions notes clearly and fully documented, including the exact steps taken for resolution?
    # Q-5: Check if the Resolution column contains keywords like 'KBA', 'Solution article', 'Auto resolved' and user confirmation variations
    resolution = current_row['Resolution']

    # Define variations of 'user confirmation'
    user_confirmation_keywords = ['confirmed', 'confirm', 'user confirmation', 'confirmation']

    # Check if the resolution contains any of the keywords
    if pd.notna(resolution) and (
    'kba' in resolution.lower() or 
    'solution article' in resolution.lower() or 
    'auto resolved' in resolution.lower() or 
    any(keyword in resolution.lower() for keyword in user_confirmation_keywords)
        ):
        audit_row['Are the resolutions notes clearly and fully documented, including the exact steps taken for resolution? (Section 10a)'] = 'No Audit'
    else:
        audit_row['Are the resolutions notes clearly and fully documented, including the exact steps taken for resolution? (Section 10a)'] = 'Normal Audit required'


    # Section 10e: If no solution article existed, did the technician submit a new solution article request?
    # Q-6
    if any(keyword in str(resolution).lower() for keyword in ['new kba', 'new solution article']):
        audit_row['If no solution article existed, did the technician submit a new solution article request? (Section 10e)'] = 'Yes'
    else:
        audit_row['If no solution article existed, did the technician submit a new solution article request? (Section 10e)'] = 'No'

    # Append the updated row to the list of rows for the Audit DataFrame
    updated_audit_rows.append(audit_row)

# Create a DataFrame from the updated rows
updated_audit_df = pd.DataFrame(updated_audit_rows)

# Save the updated audit DataFrame back to Excel
updated_audit_df.to_excel(r'D:\Automation\doc\Updated_Service_Desk_Incident_Management_Ticket_Audits.xlsx', index=False)

print("Audit process with Groq completed and saved.")
