import pandas as pd
import os
from langchain_core.prompts import ChatPromptTemplate
from langchain_groq import ChatGroq
from dotenv import load_dotenv
from fuzzywuzzy import fuzz

# Load environment variables from .env file
load_dotenv()

# Ensure that your Groq API Key is set as an environment variable
api_key = os.getenv("GROQ_API_KEY")
if not api_key:
    raise ValueError("API key for Groq is missing. Please set the GROQ_API_KEY environment variable.")

# Initialize ChatGroq with the Groq model you want to use
chat = ChatGroq(temperature=0, model_name="llama-3.3-70b-versatile")

# Define the path for the Excel files
current_day_path = r'D:\Automation\doc\doc1\Current_day_by_Technician.xlsx'
audit_path = r'D:\Automation\doc\doc1\Service_Desk_Incident_Management_Ticket_Audits.xlsx'

# Load the data from the Excel files
try:
    current_day_df = pd.read_excel(current_day_path)
    audit_df = pd.read_excel(audit_path)
    print("Excel files loaded successfully!")
except Exception as e:
    print(f"Error loading Excel files: {e}")
    raise

# Update Service_Desk_Incident_Management_Ticket_Audits with the relevant data from Current_day_by_Technician
audit_df['Technician'] = current_day_df['Technician']
audit_df['Request ID'] = current_day_df['RequestID']
audit_df['Subject'] = current_day_df['Subject']
audit_df['Completed Date'] = current_day_df['Resolved Time']

# Set the Manager column to "Saravanan"
audit_df['Manager'] = 'Saravanan'

# Update the "Has the 'Requester Name' been updated to reflect who the ticket is for?" column based on 'Requester' column in Current_day_by_Technician
audit_df['Has the "Requester Name" been updated to reflect who the ticket is for? (Section 2b)'] = current_day_df['Requester'].apply(
    lambda x: "No Audit" if pd.notna(x) else "Normal Audit required"
)

# Update the "Notes" column based on whether the 'Requester' is missing
audit_df['Notes'] = current_day_df['Requester'].apply(
    lambda x: "Requester name missing" if pd.isna(x) else ""
)

# Function to generate the correct subject header
def generate_subject_header(issue_description):
    words = issue_description.split()
    first_word = words[0].upper()
    rest_of_description = " ".join(words[1:])
    return f"{first_word} - {rest_of_description}"

# Function to check the similarity between generated and actual subject
def check_subject_format(issue_description, actual_subject):
    generated_subject = generate_subject_header(issue_description)
    similarity = fuzz.ratio(generated_subject, actual_subject)
    if similarity >= 70:
        return "Yes"
    else:
        return "Manual audit required"
# Function to check if the generated subject matches the naming convention
def check_subject_with_model(issue_description, actual_subject):
    generated_subject = generate_subject_header(issue_description)
    
    prompt = f"""
    Check if the following generated subject follows the naming convention "One word in uppercase - Brief description of the issue":
    - Correct Format Example: CLOCK - Adjust and Add EST clock
    - Incorrect Format Example: Password - No access to Work
    
    Generated Subject: "{generated_subject}"
    Actual Subject: "{actual_subject}"

    Response:
    - If the generated subject matches the correct format with 70% similarity or more, return 'Yes'.
    - If the similarity is below 70%, return 'Manual audit required'.
    """

    try:
        # Send the prompt to the Groq model and get a response
        response = chat.invoke([prompt])  # Pass the prompt inside a list
        response_text = response.content.strip().lower()

        # If the response mentions "Yes", consider it a match, otherwise a manual audit is required
        if "yes" in response_text:
            return "Yes"
        else:
            return "Manual audit required"
    except Exception as e:
        print(f"Error calling Groq model: {e}")
        return "Manual audit required"  # Default to "Manual audit required" in case of error
    

# Apply the logic for "Has the ticket 'Subject' been updated to leverage the naming convention" column
# audit_df['Has the ticket "Subject" been updated to leverage the naming convention "SERVICE - Brief Description of Issue or Request"? (Section 5a ix)'] = current_day_df['Issue Description'].apply(
#     lambda x: check_subject_format(x, current_day_df.loc[current_day_df['Issue Description'] == x, 'Subject'].values[0])
# )
# Apply the function to check the subject format using the Groq model
audit_df['Has the ticket "Subject" been updated to leverage the naming convention "SERVICE - Brief Description of Issue or Request"? (Section 5a ix)'] = current_day_df.apply(
    lambda row: check_subject_with_model(row['Issue Description'], row['Subject']),
    axis=1
)

# Add comments to the Notes column if "Manual audit required"
def add_comments_to_notes(row):
    if row['Has the ticket "Subject" been updated to leverage the naming convention "SERVICE - Brief Description of Issue or Request"? (Section 5a ix)'] == "Manual audit required":
        # Add a comment to the Notes column
        if row['Notes']:
            row['Notes'] += ", Subject format incorrect"
        else:
            row['Notes'] = "Subject format incorrect"
    return row

# Apply the function to add comments to the Notes column
audit_df = audit_df.apply(add_comments_to_notes, axis=1)

# Define the prompt for Groq model to check if solution article or "KBA" is mentioned
def check_solution_article_with_groq(resolution):
    # prompt = f"Check the following resolution and determine if it mentions 'KBA', 'Solution article', or 'Auto resolved': \n\nResolution: {resolution}\n\nResponse:"
    prompt = f"Check the following resolution and return 'No Audit' if it mentions 'KBA', 'Solution article', or 'Auto resolved'. If none of these are mentioned, return 'Normal Audit required': \n\nResolution: {resolution}\n\nResponse:"

    try:
        response = chat.invoke([prompt])  # Pass the prompt inside a list
        
        response_text = response.content.strip().lower()  # Extract text from the first response object
        # print(response_text)
        # Check if any of the target terms are present in the response content
        # if any(keyword in response_text for keyword in ['kba', 'solution article', 'auto resolved']):
        #     return "No Audit"
        # else:
        #     return "Normal Audit required"
        if response_text == "no audit":
            return "No Audit"
        else:
            return "Normal Audit required"
    except Exception as e:
        print(f"Error calling Groq model: {e}")
        return "Normal Audit required"  # Default to "Normal Audit required" in case of error

def check_solution_article_with_groq1(resolution):
    # prompt = f"Check the following resolution and determine if it mentions 'KBA', 'Solution article', or 'Auto resolved': \n\nResolution: {resolution}\n\nResponse:"
    prompt = f"Check the following resolution and return 'No Audit' if it mentions 'KBA', 'Solution article', or 'Auto resolved' with user confirmation must needed. If none of these are mentioned, return 'Normal Audit required': \n\nResolution: {resolution}\n\nResponse:"

    try:
        response = chat.invoke([prompt])  # Pass the prompt inside a list
        
        response_text = response.content.strip().lower()  # Extract text from the first response object
        # print(response_text)
        # Check if any of the target terms are present in the response content
        # if any(keyword in response_text for keyword in ['kba', 'solution article', 'auto resolved']):
        #     return "No Audit"
        # else:
        #     return "Normal Audit required"
        if response_text == "no audit":
            return "No Audit"
        else:
            return "Normal Audit required"
    except Exception as e:
        print(f"Error calling Groq model: {e}")
        return "Normal Audit required"  # Default to "Normal Audit required" in case of error
def check_solution_article_with_groq2(resolution):
    if not resolution or pd.isna(resolution):
        return "no"  # Handle missing or empty resolutions gracefully

    prompt = f"""
        Check the following resolution:
        - If the resolution mentions 'KBA', 'Solution article', or 'Auto resolved', or user confirmation is needed, return 'n/a - Existing Article'.
        - If a new 'KBA -105' or 'KBA -106' is mentioned, or if a new solution article request has been uploaded or submitted, return 'Yes'.
        - If none of the above conditions are met, return 'No'.
    Resolution: {resolution}
    """

    try:
        response = chat.invoke([prompt])  # Pass the prompt inside a list
        response_text = response.content.strip().lower()  # Extract text from the first response object

        if response_text == "n/a - existing article":
            return "n/a - Existing Article"
        elif response_text == "no":
            return "No"
        elif response_text == "yes":
            return "Yes"
        else:
            return "No"
    except Exception as e:
        print(f"Error calling Groq model: {e}")
        return "Normal Audit required"  # Default to "Normal Audit required" in case of error

# Apply the function to check each resolution and update the audit dataframe
audit_df['Did the technician search for and note a relevant Solution article, if one exists? (Section 8b)'] = current_day_df['Resolution'].apply(
    lambda x: check_solution_article_with_groq(x)
)

audit_df['Are the resolutions notes clearly and fully documented, including the exact steps taken for resolution? (Section 10a)'] = current_day_df['Resolution'].apply(
    lambda x: check_solution_article_with_groq1(x)
)
audit_df['If no solution article existed, did the technician submit a new solution article request? (Section 10e)'] = current_day_df['Resolution'].apply(
    lambda x: check_solution_article_with_groq2(x)
)

# Function to add comments to the Notes column if necessary
def add_kba_missing_comments(row):
    if row['Did the technician search for and note a relevant Solution article, if one exists? (Section 8b)'] == "Normal Audit required":
        if row['Notes']:  # If there are already existing notes
            row['Notes'] += ", KBA is missing"
        else:
            row['Notes'] = "KBA is missing"
    return row

# Apply the function to add comments to the Notes column
audit_df = audit_df.apply(add_kba_missing_comments, axis=1)

# Reuse the result from Section 8b for Section 9
audit_df['Did the technician provide clear and detailed notes, documenting all steps taken during troubleshooting? (Section 9)'] = audit_df['Did the technician search for and note a relevant Solution article, if one exists? (Section 8b)']
# Final check before saving
# print(audit_df[['Did the technician search for and note a relevant Solution article, if one exists? (Section 8b)',
#                 'Did the technician provide clear and detailed notes, documenting all steps taken during troubleshooting? (Section 9)']].head())

# Save the updated DataFrame back to an Excel file
updated_audit_path = r'D:\Automation\doc\doc1\Updated_Service_Desk_Incident_Management_Ticket_Audits.xlsx'
audit_df.to_excel(updated_audit_path, index=False)

print("Service Desk Incident Management Ticket Audits updated successfully!")
