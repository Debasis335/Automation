import pandas as pd
import os
from langchain_core.prompts import ChatPromptTemplate
from langchain_groq import ChatGroq
from dotenv import load_dotenv
from fuzzywuzzy import fuzz
from difflib import SequenceMatcher
from sentence_transformers import SentenceTransformer, util

# Load environment variables from .env file
load_dotenv()

# Ensure that your Groq API Key is set as an environment variable
api_key = os.getenv("GROQ_API_KEY")
if not api_key:
    raise ValueError("API key for Groq is missing. Please set the GROQ_API_KEY environment variable.")

# Initialize ChatGroq with the Groq model you want to use
chat = ChatGroq(temperature=0, model_name="llama-3.3-70b-versatile")
# Load the model globally (once)
semantic_model = SentenceTransformer('all-MiniLM-L6-v2')
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
audit_df['Manager'] = ''

#Q1
# Update the "Has the 'Requester Name' been updated to reflect who the ticket is for?" column based on 'Requester' and 'On Behalf Of User'
audit_df['Has the "Requester Name" been updated to reflect who the ticket is for? (Section 2b)'] = current_day_df.apply(
    lambda row: "Yes" if (
        (str(row['Request Mode']).lower() not in ['phone', 'service portal']) or  # Skip validation if Request Mode is not phone or service portal
        pd.isna(row['On Behalf Of User']) or  # Skip validation if 'On Behalf Of User' is empty
        str(row['On Behalf Of User']).lower() == 'not assigned' or  # Skip validation if 'On Behalf Of User' is "Not assigned"
        row['Requester'] == row['On Behalf Of User']  # Pass as Yes if 'Requester' matches 'On Behalf Of User'
    ) else "Manual Audit required", axis=1
)



# # Update the "Notes" column based on whether the 'Requester' is missing
# audit_df['Notes'] = current_day_df['Requester'].apply(
#     lambda x: "Requester name missing" if pd.isna(x) else ""
# )

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

def check_subject_with_model(issue_description, actual_subject, row, audit_df):
    if pd.isna(issue_description) or pd.isna(actual_subject):
        return "Manual Audit required"

    prompt = f"""
    Generate a subject header for the following issue description. The subject should follow the naming convention "One word in uppercase - Brief description of the issue":
    - Correct Format Example: CLOCK - Adjust and Add EST clock
    - Incorrect Format Example: Password - No access to Work

    Issue Description: "{issue_description}"

    Generated Subject Header:
    """

    try:
        response = chat.invoke([prompt])
        generated_subject = response.content.strip()

        # Semantic similarity using sentence-transformers
        embedding1 = semantic_model.encode(generated_subject, convert_to_tensor=True)
        embedding2 = semantic_model.encode(actual_subject, convert_to_tensor=True)
        similarity_score = util.cos_sim(embedding1, embedding2).item()  # cosine similarity value

        if similarity_score >= 0.60:
            return "Yes"
        else:
           current_note = str(audit_df.at[row.name, 'Notes']) if pd.notna(audit_df.at[row.name, 'Notes']) else ""
           audit_df.at[row.name, 'Notes'] = current_note + " |Q2- " + generated_subject+"-"+ str(similarity_score)
           return "Manual audit required"

    except Exception as e:
        print(f"Error calling model: {e}")
        return "Manual audit required"
    

# Apply the logic for "Has the ticket 'Subject' been updated to leverage the naming convention" column
# audit_df['Has the ticket "Subject" been updated to leverage the naming convention "SERVICE - Brief Description of Issue or Request"? (Section 5a ix)'] = current_day_df['Issue Description'].apply(
#     lambda x: check_subject_format(x, current_day_df.loc[current_day_df['Issue Description'] == x, 'Subject'].values[0])
# )
# Apply the function to check the subject format using the Groq model
#Q-2
audit_df['Has the ticket "Subject" been updated to leverage the naming convention "SERVICE - Brief Description of Issue or Request"? (Section 5a ix)'] = current_day_df.apply(
    lambda row: check_subject_with_model(row['Issue Description'], row['Subject'], row, audit_df),
    axis=1
)

# Add comments to the Notes column if "Manual audit required"
def add_comments_to_notes(row):
    if row['Has the ticket "Subject" been updated to leverage the naming convention "SERVICE - Brief Description of Issue or Request"? (Section 5a ix)'] == "Manual audit required":
        # Add a comment to the Notes column
        if row['Notes']:
            row['Notes'] += ", Subject format is not as expected"
        else:
            row['Notes'] = "Subject format is not as expected"
    return row

# Apply the function to add comments to the Notes column
# audit_df = audit_df.apply(add_comments_to_notes, axis=1)

# Define the prompt for Groq model to check if solution article or "KBA" is mentioned
def check_solution_article_with_groq(resolution,row):
    
    # Prompt to send to the Groq model asking it to understand if the issue was auto-resolved or if KBA/Solution Article was followed or a new article submitted
    prompt = f"""
    Please read the following resolution and analyze it carefully. Based on the content of the resolution:
    1. If the issue was auto-resolved(for example, the user resolved the issue themselves, or the issue resolved on its own or without manual intervention), please return 'Yes'.
    2. If the technician referred to a Knowledge Base Article (KBA) or followed Solution Article steps to resolve the issue, return 'Yes'.
    3. If the resolution describes a new KBA submitted (example- 'KBA -105' or 'KBA -106') or if a new solution article has been uploaded or submitted as part of resolving the issue, return 'New article submitted'.
    4. If none of the above conditions are met, return 'Manual Audit required'.

    Resolution: {resolution}

    Response:
    """

    try:
        # Pass the prompt to the Groq model and get the response
        response = chat.invoke([prompt])  # Pass the prompt inside a list
        
        response_text = response.content.strip().lower()  # Normalize to lowercase
        
        if "yes" in response_text or "new article submitted" in response_text:
            return "Yes"
        else:
            current_note = str(audit_df.at[row.name, 'Notes']) if pd.notna(audit_df.at[row.name, 'Notes']) else ""
            audit_df.at[row.name, 'Notes'] = current_note + " |Q3- " + response_text
            return "Manual audit required"
    except Exception as e:
        print(f"Error calling Groq model: {e}")
        return "Manual Audit required"

# def update_notes_based_on_resolution(row):
#     # Check the response from the Groq model in the column 'Did the technician search for and note a relevant Solution article, if one exists? (Section 8b)'
#     resolution_check = row['Did the technician search for and note a relevant Solution article, if one exists? (Section 8b)']

#     # Initialize the Notes field (if empty, set as empty string)
#     notes = row.get('Notes', "")

#     if resolution_check == "New article submitted":
#         if notes:
#             row['Notes'] = notes + ", Technician submitted a new article"
#         else:
#             row['Notes'] = "Technician submitted a new article"
#     elif resolution_check == "Yes":
#         if notes:
#             row['Notes'] = notes + ", Technician referred to KBA/Solution Article"
#         else:
#             row['Notes'] = "Technician referred to KBA/Solution Article"
#     elif resolution_check == "Manual Audit required":
#         if notes:
#             row['Notes'] = notes + ", No KBA or Solution Article mentioned"
#         else:
#             row['Notes'] = "No KBA or Solution Article mentioned"
    
#     return row

# # Applying the function to the DataFrame
# audit_df = audit_df.apply(update_notes_based_on_resolution, axis=1)

def check_solution_article_with_groq1(resolution,row):
    # Define the prompt with instructions that require user confirmation
    prompt = f"""
    Please read the following resolution and analyze it carefully. Based on the content of the resolution:

    1. If the issue was auto-resolved (for example, the user resolved it themselves, or the issue resolved on its own), AND the user confirmed it, return: Yes
    2. If the technician referred to a Knowledge Base Article (KBA) or followed Solution Article steps AND the user confirmed that the solution was followed, return: Yes
    3. If the technician submitted a new KBA or Solution Article request AND it is mentioned that user feedback led to it, return: New article submitted
    4. If none of the above apply or user confirmation is not mentioned, return: Manual audit required

    Make sure to base your response only on the presence of **user confirmation** in the resolution text. For example, phrases like:
    - "User confirmed issue resolved on its own"
    - "User confirmed they followed technician steps"
    - "User followed the KBA instructions and confirmed resolution"

    Only respond with one of the following exact outputs: **Yes**, **New article submitted**, or **Manual audit required**.

    Resolution:
    {resolution}

    Your response:
    """

    try:
        # Invoke the prompt with Groq
        response = chat.invoke([prompt])  # Pass the prompt inside a list
        response_text = response.content.strip().lower()  # Extract the response text and ensure it's lowercase
        # print(response_text)
        if response_text == "yes":
            return "Yes"
        else:
            # Update Notes only if it's not "yes"
            current_note = str(audit_df.at[row.name, 'Notes']) if pd.notna(audit_df.at[row.name, 'Notes']) else ""
            audit_df.at[row.name, 'Notes'] = current_note + f" |Q5- {response_text}"
            return "Manual audit required"

    except Exception as e:
        print(f"Error calling Groq model1: {e}")
        return "Manual Audit required"  # Default to "Manual Audit required" in case of error

    
def check_solution_article_with_groq2(resolution,row):
    if not resolution or pd.isna(resolution):
        return "No"  # Handle missing or empty resolutions gracefully

    # Define the prompt for Groq model
    prompt = f"""
    Please carefully analyze the following resolution and determine the appropriate response based on the context:
    1. If the resolution describes the issue as being resolved automatically (for example, the user resolved the issue themselves, or the issue resolved on its own or without manual intervention), return 'n/a - Existing Article'.
    2. If the resolution describes that the technician referred to or used a Knowledge Base Article (KBA) or followed the steps from a Solution article to resolve the issue, return 'n/a - Existing Article'.
    3. If the resolution describes a new KBA submitted (example- 'KBA -105' or 'KBA -106') or if a new solution article has been uploaded or submitted as part of resolving the issue, return 'Yes'.
    4. If none of the above conditions are met, return 'No'.

    Resolution: {resolution}
    """

    try:
        # Invoke the prompt with the Groq model
        response = chat.invoke([prompt])  # Pass the prompt inside a list
        response_text = response.content.strip().lower()  # Extract text from the first response object
        # print(response_text)
        # Normalize text
        cleaned_text = response_text.replace(".", "").replace("–", "-").strip()
        # print(cleaned_text)

        # Return based on model's understanding of the context
        if "n/a" in cleaned_text and "article" in cleaned_text:
            return "n/a - Existing Article"
        elif "yes" in cleaned_text:
            return "Yes"
        elif "no" in cleaned_text:
            return "No"
        else:
            return "No"

    except Exception as e:
        print(f"Error calling Groq model2: {e}")
        return "Normal Audit required"  # Default to "Normal Audit required" in case of error

#Q-3
# Apply the function to check each resolution and update the audit dataframe
# audit_df['Did the technician search for and note a relevant Solution article, if one exists? (Section 8b)'] = current_day_df['Resolution'].apply(
#      lambda x: check_solution_article_with_groq(x)    
# )
audit_df['Notes'] = audit_df['Notes'].astype(str)
audit_df['Did the technician search for and note a relevant Solution article, if one exists? (Section 8b)'] = current_day_df.apply(
    lambda row: check_solution_article_with_groq(row['Resolution'], row),
    axis=1
)


#Q-5
# audit_df['Are the resolutions notes clearly and fully documented, including the exact steps taken for resolution? (Section 10a)'] = current_day_df['Resolution'].apply(
#     lambda x: check_solution_article_with_groq1(x)
# )
#new
audit_df['Are the resolutions notes clearly and fully documented, including the exact steps taken for resolution? (Section 10a)'] = current_day_df.apply(
    lambda row: check_solution_article_with_groq1(row['Resolution'], row),
    axis=1)
#Q-6
# audit_df['If no solution article existed, did the technician submit a new solution article request? (Section 10e)'] = current_day_df['Resolution'].apply(
#     lambda x: check_solution_article_with_groq2(x)
# )
# for idx, row in current_day_df.iterrows():
#     response = check_solution_article_with_groq2(row['Resolution'], row)
#     audit_df.at[idx, 'If no solution article existed, did the technician submit a new solution article request? (Section 10e)'] = response

audit_df['If no solution article existed, did the technician submit a new solution article request? (Section 10e)'] = current_day_df.apply(
    lambda row: check_solution_article_with_groq2(row['Resolution'], row),
    axis=1
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
#Q-4
audit_df["Did the technician provide clear and detailed notes, documenting all steps taken during troubleshooting? (Section 9)"] = audit_df["Did the technician search for and note a relevant Solution article, if one exists? (Section 8b)"]

# Final check before saving
# print(audit_df[['Did the technician search for and note a relevant Solution article, if one exists? (Section 8b)',
#                 'Did the technician provide clear and detailed notes, documenting all steps taken during troubleshooting? (Section 9)']].head())

# Save the updated DataFrame back to an Excel file
updated_audit_path = r'D:\Automation\doc\doc1\Updated_Service_Desk_Incident_Management_Ticket_Audits.xlsx'
audit_df.to_excel(updated_audit_path, index=False)

print("Service Desk Incident Management Ticket Audits updated successfully!")
