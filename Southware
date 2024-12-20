import streamlit as st
import os
import requests
import pandas as pd
from azure.core.credentials import AzureKeyCredential
from azure.ai.formrecognizer import DocumentAnalysisClient
from rapidfuzz import fuzz

# Azure Form Recognizer Credentials
endpoint = "https://southware.cognitiveservices.azure.com/"
key = "c39681fbf8654d7e918222af34b02080"
headers = {
    "Ocp-Apim-Subscription-Key": "a9f7a6f74b644a62a44c997695d9c1ce",
    "Accept": "application/json"
}

# Initialize the Document Analysis Client
document_analysis_client = DocumentAnalysisClient(
    endpoint=endpoint, credential=AzureKeyCredential(key)
)

# Function to clean unwanted characters from text
def clean_text(value):
    if value:
        return value.replace(",", "").replace(":", "").replace(".", "")
    return value

# Function to apply address abbreviations
def apply_abbreviations(address):
    abbreviations = {
        "STREET": "ST", "AVENUE": "AVE", "BOULEVARD": "BLVD", "ROAD": "RD",
        "DRIVE": "DR", "LANE": "LN", "COURT": "CT", "CIRCLE": "CIR", "PLACE": "PL",
        "SQUARE": "SQ", "TERRACE": "TER", "PARKWAY": "PKWY", "HIGHWAY": "HWY",
        "APARTMENT": "APT", "SUITE": "STE", "BUILDING": "BLDG", "NORTH": "N", "SOUTH": "S",
        "EAST": "E", "WEST": "W", "FLOOR": "FL"
    }
    if address and isinstance(address, str):
        words = address.upper().split()
        return " ".join([abbreviations.get(word, word) for word in words])
    return "N/A"

# Streamlit App Layout
st.title("Document Data Extraction App")

uploaded_file = st.file_uploader("Upload a PDF or SM file", type=["pdf", "sm"])

if uploaded_file:
    with st.spinner("Analyzing document..."):
        try:
            poller = document_analysis_client.begin_analyze_document("prebuilt-document", document=uploaded_file)
            result = poller.result()

            needed_fields = [
                "Legal Business Name", "Company/dba:", "Mailing Address:", "City:", "State:", "Zip:", "Zip",
                "Acct Contact Name:", "Contact Phone", "Contact email", "AP Contact Name:",
                "Phone:", "AP Email:", "Invoice Email:", "PO Required?", "Yes", "No"
            ]

            extracted_data = {field.upper(): "N/A" for field in needed_fields}

            for kv_pair in result.key_value_pairs:
                for field in needed_fields:
                    match_ratio = fuzz.ratio(kv_pair.key.content, field)
                    if match_ratio > 95:
                        key_content = field.upper()
                        value_content = kv_pair.value.content if kv_pair.value else "N/A"
                        cleaned_value = clean_text(value_content.upper())
                        if extracted_data[key_content] == "N/A":
                            extracted_data[key_content] = cleaned_value
                        break

            # Apply abbreviations
            mailing_address = extracted_data.get("MAILING ADDRESS:")
            if mailing_address != "N/A":
                extracted_data["MAILING ADDRESS:"] = apply_abbreviations(mailing_address)

            # Display extracted data
            st.subheader("Extracted Data")
            st.json(extracted_data)

            state = extracted_data.get("STATE:")
            zip_code = extracted_data.get("ZIP")

            if zip_code and '-' in zip_code:
                zip_code = zip_code.split('-')[0]

            if zip_code and zip_code != "N/A":
                url = f"https://global.metadapi.com/zipc/v1/zipcodes/{zip_code}"
                response = requests.get(url, headers=headers)
                if response.status_code == 200:
                    data = response.json()
                    county_name = data['data']['titleCaseCountyName']
                    state_code = data['data']['stateCode']
                    st.write(f"**County Name:** {county_name}, **State:** {state_code}")
                else:
                    st.error(f"Failed to retrieve data. Status code: {response.status_code}")

            # Excel Lookup Section (Bundled File)
            st.subheader("Excel Lookup")
            file_path = "Southware4.xlsx"
            if os.path.exists(file_path):
                sheet_name = "Sheet1"
                df = pd.read_excel(file_path, sheet_name=sheet_name)
                
                def find_city_or_county(df, state_code, extracted_city, extracted_county):
                    extracted_city = extracted_city.upper() if isinstance(extracted_city, str) else ""
                    extracted_county = extracted_county.upper() if isinstance(extracted_county, str) else ""

                    for index, row in df.iterrows():
                        city_value = row['CityColumn'] if isinstance(row['CityColumn'], str) else ""
                        if state_code == row['St'] and extracted_city in city_value.upper():
                            st.write(f"**Found matching city and state.** TaxCode: {row['TaxCode']}")
                            return

                    for index, row in df.iterrows():
                        county_value = row['CountyName'] if isinstance(row['CountyName'], str) else ""
                        if state_code == row['St'] and extracted_county in county_value.upper():
                            st.write(f"**Found matching county and state.** TaxCode: {row['TaxCode']}")
                            return

                find_city_or_county(df, state_code, extracted_data.get("CITY:"), county_name)
            else:
                st.error("Bundled Excel file not found.")

        except Exception as e:
            st.error(f"An error occurred: {e}")
