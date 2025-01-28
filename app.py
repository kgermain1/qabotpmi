import streamlit as st
from openai import OpenAI
from docx import Document
import json
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import re

client = OpenAI(api_key=st.secrets["openai"]["api_key"])

# Get the credentials JSON from the secrets manager
credentials_info = json.loads(st.secrets["google"]["credentials_json"])

# Load the credentials from Streamlit secrets
credentials_json = st.secrets["google"]["credentials_json"]
credentials_dict = json.loads(credentials_json)

# Path to your logo image (can be a local path or a URL)
logo_path = "philip-morris-international-pmi-vector-logo.png"
st.image(logo_path, use_container_width=True)

# Google Sheets Credentials
GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1INz9LD7JUaZiIbY4uoGId0riIhkltlE6aroLKOAWtNo/edit#gid=0"

# Use the latest OpenAI GPT model
MODEL_NAME = "gpt-4"

# Initialize Google Sheets
def get_google_sheets():
    # Set the scope and authenticate
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_dict(credentials_dict, scope)
    client = gspread.authorize(creds)
    return client.open_by_url(GOOGLE_SHEET_URL)

def get_tab_names(sheet):
    """Fetch tab names from the Google Sheet."""
    return [worksheet.title for worksheet in sheet.worksheets()]

def get_rules(sheet, client, market):
    """Retrieve rules and rule names for the selected client and market."""
    # Load the "ALL CLIENTS" tab
    all_clients_tab = sheet.worksheet("ALL CLIENTS")
    all_clients_rules = pd.DataFrame(all_clients_tab.get_all_records())

    # Load the client-specific tab
    client_tab = sheet.worksheet(client)
    client_rules = pd.DataFrame(client_tab.get_all_records())

    # Filter "ALL CLIENTS" rules to include only those for the selected market or "All"
    filtered_all_clients_rules = all_clients_rules[
        (all_clients_rules["Market"] == market) | (all_clients_rules["Market"] == "All")
    ]

    # Filter client-specific rules for the selected market or "All"
    filtered_client_rules = client_rules[
        (client_rules["Market"] == market) | (client_rules["Market"] == "All")
    ]

    # Combine the filtered rules from both tabs
    combined_rules = pd.concat([filtered_all_clients_rules, filtered_client_rules], ignore_index=True)

    # Ensure the "Rule" and "Rule Name" columns are present
    if "Rule" not in combined_rules or "Rule Name" not in combined_rules:
        raise ValueError("The required 'Rule' or 'Rule Name' columns are missing in the Google Sheet.")

    return combined_rules
    """Retrieve rules and rule names for the selected client and market."""
    all_clients_tab = sheet.worksheet("ALL CLIENTS")
    client_tab = sheet.worksheet(client)

    # Get all rules from the relevant tabs
    all_clients_rules = pd.DataFrame(all_clients_tab.get_all_records())
    client_rules = pd.DataFrame(client_tab.get_all_records())

    # Filter rules by the selected market or "All"
    market_specific_rules = client_rules[
        (client_rules["Market"] == market) | (client_rules["Market"] == "All")
    ]

    # Combine rules from "ALL CLIENTS" and client-specific tabs
    combined_rules = pd.concat([all_clients_rules, market_specific_rules], ignore_index=True)

    # Ensure the "Rule" and "Rule Name" columns are present
    if "Rule" not in combined_rules or "Rule Name" not in combined_rules:
        raise ValueError("The required 'Rule' or 'Rule Name' columns are missing in the Google Sheet.")

    return combined_rules

def group_rules_by_ruleset(rules_df):
    """Group rules by Ruleset."""
    return {ruleset: group for ruleset, group in rules_df.groupby("Ruleset")}

def read_docx(file):
    """Reads the content of a Word document."""
    doc = Document(file)
    text = "\n".join([paragraph.text for paragraph in doc.paragraphs if paragraph.text])
    return text

def check_compliance(document_text, rules_df):
    """Check compliance for each Ruleset."""
    grouped_rules = group_rules_by_ruleset(rules_df)
    reports = {}

    for ruleset, rules in grouped_rules.items():
        # Create a mapping of Rules to Rule Names
        rule_name_mapping = dict(zip(rules["Rule"], rules["Rule Name"]))
        rules_list = "\n".join([f"- {rule} (Rule Name: {rule_name})" for rule, rule_name in rule_name_mapping.items()])

        messages = [
            {"role": "system", "content": "You are an expert in compliance and tone-of-voice review."},
            {
                "role": "user",
                "content": f"""
Document Content:
{document_text}

Rules for {ruleset}:
{rules_list}

Analyze the document for compliance with the rules. For any violations, reference the 'Rule Name' exactly as provided in the list of rules (do not invent or modify Rule Names).

Format the report as follows:
- State whether the document is "Compliant" or "Non-Compliant".
- Provide details for any violations, referencing only the Rule Names provided, using the following format:
    (number) Rule Name: State the Rule Name associated with the violation (as provided in the rules list).
    Explanation: Provide a short explanation for the violation.
    
Rules that have not been violated should not be featured in your report, and do not mention which rules have not been violated.
""",
            },
        ]

        try:
            response = client.chat.completions.create(
                model=MODEL_NAME,
                messages=messages,
                max_tokens=2000,
                temperature=0.5,
            )

            raw_report = response.choices[0].message.content

            # Replace all rule texts with their corresponding Rule Names using strict mapping
            for rule, rule_name in rule_name_mapping.items():
                raw_report = re.sub(rf"\b{re.escape(rule)}\b", rule_name, raw_report)

            reports[ruleset] = raw_report
        except Exception as e:
            reports[ruleset] = f"An error occurred: {str(e)}"

    return reports


# Streamlit App
st.title("Welcome to QAbot")

st.sidebar.title("Instructions")
st.sidebar.write("""
1. Select your market.
2. Upload a Word document (.docx).
3. Click the "Check Compliance" button to evaluate.
""")

# Google Sheets Initialization
try:
    sheet = get_google_sheets()
    tab_names = get_tab_names(sheet)
except Exception as e:
    st.error("Error connecting to Google Sheets. Please check your credentials.")
    st.stop()

# Dropdowns for client and market
selected_client = "PMI"

if selected_client:
    market_tab = pd.DataFrame(sheet.worksheet(selected_client).get_all_records())
    available_markets = market_tab["Market"].unique().tolist()
    selected_market = st.selectbox("Select a Market", available_markets)

# File Upload
uploaded_file = st.file_uploader("Upload a Word document", type=["docx"])

if st.button("Check Compliance"):
    if uploaded_file and selected_client and selected_market:
        document_text = read_docx(uploaded_file)
        rules_df = get_rules(sheet, selected_client, selected_market)

        with st.spinner("Checking compliance..."):
            compliance_reports = check_compliance(document_text, rules_df)

            for ruleset, report in compliance_reports.items():
                st.subheader(f"{ruleset} Report")
                if report.startswith("An error occurred"):
                    st.error(report)
                else:
                    st.text_area(f"{ruleset} Report", value=report, height=300, disabled=True)
    else:
        st.error("Please upload a document, select a client, and a market.")
