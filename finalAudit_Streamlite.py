import pandas as pd
import os
from langchain_core.prompts import ChatPromptTemplate
from langchain_groq import ChatGroq
from dotenv import load_dotenv
from fuzzywuzzy import fuzz
from difflib import SequenceMatcher
from sentence_transformers import SentenceTransformer, util
import streamlit as st
from io import BytesIO
# Load environment variables from .env file

load_dotenv()

# Streamlit UI
st.title("🧾 Audit Automation")
# Ensure that your Groq API Key is set as an environment variable
api_key = os.getenv("GROQ_API_KEY")
if not api_key:
    raise ValueError("API key for Groq is missing. Please set the GROQ_API_KEY environment variable.")

# Initialize ChatGroq with the Groq model you want to use
chat = ChatGroq(temperature=0, model_name="llama-3.3-70b-versatile")
# Load the model globally (once)
semantic_model = SentenceTransformer('all-MiniLM-L6-v2')
# Define the path for the Excel files
# current_day_path = r'D:\Automation\doc\doc1\Current_day_by_Technician.xlsx'
# audit_path = r'D:\Automation\doc\doc1\Service_Desk_Incident_Management_Ticket_Audits.xlsx'

current_day_path = st.file_uploader("Upload 'Current_day_by_Technician.xlsx'", type=["xlsx"])
# audit_path = st.file_uploader("Upload 'Service_Desk_Incident_Management_Ticket_Audits.xlsx'", type=["xlsx"])
audit_path = r'D:\Automation\doc\doc1\Service_Desk_Incident_Management_Ticket_Audits.xlsx'

# Load the data from the Excel files
if current_day_path is not None and audit_path is not None:
    if st.button("Start Audit Processing"):
        try:
          with st.spinner('Excel files loaded successfully! Started processing...'):          
            current_day_df = pd.read_excel(current_day_path)
            audit_df = pd.read_excel(audit_path)
            # st.session_state.files_loaded = True
            # if st.session_state.get("files_loaded"):
            #    st.success("Excel files loaded successfully! Started processing...")
            
            # 👉 Your processing logic starts here...    


            if current_day_df is not None and not current_day_df.empty:
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
                        (str(row['Request Mode']).lower() not in ['phone', 'service portal']) or
                        pd.isna(row['On Behalf Of User']) or
                        str(row['On Behalf Of User']).lower() == 'not assigned' or
                        row['Requester'] == row['On Behalf Of User']
                    ) else "Manual Audit required", axis=1
                )

                def check_subject_with_model(issue_description, actual_subject, row, audit_df):
                    if pd.isna(issue_description) or pd.isna(actual_subject):
                        return "Manual Audit required"

                    prompt = f"""
                    Generate a subject header for the following issue description. The subject line should begin with the service name in UPPERCASE, followed by a hyphen (-) and then a brief and clear description of the issue.
                    - Correct Example: CLOCK - Adjust and Add EST clock Here, "CLOCK" is the service name in uppercase, followed by a hyphen and a short issue description.
                    - Incorrect Example: Password Reset - Password Expired In this example, the service name is not in uppercase, which doesn’t meet the formatting requirement.

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

                # Apply the function to check the subject format using the Groq model
                #Q-2
                audit_df['Has the ticket "Subject" been updated to leverage the naming convention "SERVICE - Brief Description of Issue or Request"? (Section 5a ix)'] = current_day_df.apply(
                    lambda row: check_subject_with_model(row['Issue Description'], row['Subject'], row, audit_df),
                    axis=1
                )

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

                def check_solution_article_with_groq1(resolution,row):
                    # Define the prompt with instructions that require user confirmation
                    prompt = f"""
                    Read the resolution notes provided and analyze it based on the following criteria. Response should be determined by the presence of user confirmation in the resolution text.


                    Read the resolution notes provided and analyze it based on the following criteria. Response should be determined by the presence of user confirmation in the resolution text.

                    Evaluation Guidelines:

                    1.Auto-Resolved with User Confirmation:
                    If the issue resolved on its own (e.g., the user fixed it themselves or it resolved without intervention) and the user confirmed it, return: Yes

                    2.Existing KBA Followed with User Confirmation:
                    If the technician referred to a Knowledge Base Article (KBA) or followed documented solution steps, and the user confirmed the solution was followed and effective, return: Yes

                    3.New KBA Submission Based on User Feedback:
                    If the technician submitted a new KBA or solution article request and noted that it was based on user feedback, return: New article submitted

                    4.No User Confirmation or Unclear Resolution:
                    If none of the above conditions are met, or if user confirmation is not clearly mentioned, return: Manual audit required

                    Key Phrases to Look For (Examples of User Confirmation):
                    “User confirmed issue resolved on its own (i.e., auto-resolved)”
                    “User confirmed/acknowledged the issue is resolved and working as expected”

                    Resolution:
                    {resolution}

                    Your response:
                    """

                    try:
                        # Invoke the prompt with Groq
                        response = chat.invoke([prompt])  # Pass the prompt inside a list
                        response_text = response.content.strip().lower()  # Extract the response text and ensure it's lowercase
                        # print(response_text)
                        if "yes" in response_text or "new article submitted" in response_text:
                            return "Yes"
                        else:
                            # Update Notes only if it's not "yes"
                            current_note = str(audit_df.at[row.name, 'Notes']) if pd.notna(audit_df.at[row.name, 'Notes']) else ""
                            audit_df.at[row.name, 'Notes'] = current_note + f" |Q5- No User Confirmation or Unclear Resolution: {response_text}"
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

                audit_df['Notes'] = audit_df['Notes'].astype(str)
                #Q-3
                audit_df['Did the technician search for and note a relevant Solution article, if one exists? (Section 8b)'] = current_day_df.apply(
                    lambda row: check_solution_article_with_groq(row['Resolution'], row),
                    axis=1
                )


                #Q-5
                audit_df['Are the resolutions notes clearly and fully documented, including the exact steps taken for resolution? (Section 10a)'] = current_day_df.apply(
                    lambda row: check_solution_article_with_groq1(row['Resolution'], row),
                    axis=1)
                #Q-6
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
                # #Q-4
                # audit_df["Did the technician provide clear and detailed notes, documenting all steps taken during troubleshooting? (Section 9)"] = audit_df["Did the technician search for and note a relevant Solution article, if one exists? (Section 8b)"]

                # Mapping long questions to short labels
                column_rename_map = {
                    'Has the "Requester Name" been updated to reflect who the ticket is for? (Section 2b)': 'Q-1',
                    'Has the ticket "Subject" been updated to leverage the naming convention "SERVICE - Brief Description of Issue or Request"? (Section 5a ix)': 'Q-2',
                    'Did the technician search for and note a relevant Solution article, if one exists? (Section 8b)': 'Q-3',
                    'Did the technician provide clear and detailed notes, documenting all steps taking during troubleshooting? (Section 9)': 'Q-4',
                    'Are the resolutions notes clearly and fully documented, including the exact steps taken for resolution? (Section 10a)': 'Q-5',
                    'If no solution article existed, did the technician submit a new solution article request? (Section 10e)': 'Q-6'
                }
                # Apply renaming
                audit_df.rename(columns=column_rename_map, inplace=True)
                audit_df["Q-4"] = audit_df["Q-3"]
                # Output
                st.success("✅ Audit Completed!")
                st.dataframe(audit_df)
                output = BytesIO()
                audit_df.to_excel(output, index=False)
                output.seek(0)
                st.download_button(
                        label="📥 Download Audit Report",
                        data=output,
                        file_name="Updated_Service_Desk_Incident_Management_Ticket_Audits.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                # Display reference for Q-codes
                with st.expander("ℹ️ Question Reference Guide"):
                    st.markdown("""
                    **Q-1:** Has the "Requester Name" been updated to reflect who the ticket is for? (Section 2b)  
                    **Q-2:** Has the ticket "Subject" been updated to leverage the naming convention "SERVICE - Brief Description of Issue or Request"? (Section 5a ix)  
                    **Q-3:** Did the technician search for and note a relevant Solution article, if one exists? (Section 8b)  
                    **Q-4:** Did the technician provide clear and detailed notes, documenting all steps taking during troubleshooting? (Section 9)  
                    **Q-5:** Are the resolutions notes clearly and fully documented, including the exact steps taken for resolution? (Section 10a)  
                    **Q-6:** If no solution article existed, did the technician submit a new solution article request? (Section 10e)  
                    """)
            else:
            # Fallback handling if current_day_df is None or empty
                st.warning("⚠️ No current day technician data available. Please check the input file.")
        except Exception as e:
            st.error(f"Error loading Excel files: {e}")            
else:
    st.info("Please upload files to enable the audit processing.")