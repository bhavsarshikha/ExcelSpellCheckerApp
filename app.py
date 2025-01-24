import streamlit as st
import pandas as pd
from spellchecker import SpellChecker
import re

# Function to check if a value is text (not numeric or date-like)
def is_text(value):
    return isinstance(value, str) and not bool(re.search(r'^\d+|\d{2}/\d{2}/\d{4}', value))

# Function to detect misspelled words
def find_misspelled_words(data):
    spell = SpellChecker()
    spell.word_frequency.load_words(["nov.", "dec.", "aug.", "sept.", "revrt", "backout"])  # Add custom words
    misspelled = {}
    for col in data.columns:
        if data[col].dtype == 'object':
            for value in data[col].dropna().unique():
                if is_text(value):
                    words = value.split()
                    for word in words:
                        if word not in spell and word.isalpha():  # Check only alphabetic words
                            misspelled[word] = spell.candidates(word)
    return misspelled

# Function to apply corrections to the DataFrame
def apply_corrections(data, corrections):
    for wrong_word, correct_word in corrections.items():
        if correct_word == "Do Not Correct" or not isinstance(correct_word, str) or not correct_word.strip():
            continue  # Skip corrections
        data = data.applymap(lambda x: str(x).replace(wrong_word, correct_word) if isinstance(x, str) else x)
    return data

# Function to enable file download
def download_file(data):
    output_file = "corrected_file.xlsx"
    data.to_excel(output_file, index=False)
    with open(output_file, "rb") as f:
        st.download_button(
            label="Download Corrected Excel File",
            data=f,
            file_name="Corrected_Spelling_File.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

# Streamlit UI
st.title("Excel Spelling Error Correction Tool")
st.write("Upload your Excel file to detect and correct spelling errors.")

# File upload
uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])

if uploaded_file:
    # Load and display the original file
    data = pd.read_excel(uploaded_file)
    st.write("### Original File Data:")
    st.dataframe(data)

    # Detect spelling errors
    st.write("### Identified Spelling Errors:")
    misspelled_words = find_misspelled_words(data)
    if misspelled_words:
        correction_dict = {}
        for word, suggestions in misspelled_words.items():
            user_choice = st.selectbox(
                f"Select a correction for '{word}'",
                options=["Do Not Correct", "Auto-Correct"] + list(suggestions),
            )
            # Handle "Auto-Correct" selection
            if user_choice == "Auto-Correct":
                correction_dict[word] = SpellChecker().correction(word)  # Top suggestion
            else:
                correction_dict[word] = user_choice

        # Apply corrections
        if st.button("Apply Corrections"):
            updated_data = apply_corrections(data, correction_dict)
            st.write("### Corrected File Data:")
            st.dataframe(updated_data)

            # Enable file download
            st.write("### Download Corrected File:")
            download_file(updated_data)
    else:
        st.write("No spelling errors found in the uploaded file.")
