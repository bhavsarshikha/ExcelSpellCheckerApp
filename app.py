import streamlit as st
import pandas as pd
from spellchecker import SpellChecker
from io import BytesIO

# Function to process spelling corrections
def correct_spelling(data):
    spell = SpellChecker()
    corrections = {}

    # Identify misspelled words and suggest corrections
    for col in data.select_dtypes(include=[object]).columns:
        data[col] = data[col].astype(str)  # Ensure all text is treated as strings
        misspelled = spell.unknown(" ".join(data[col]).split())
        for word in misspelled:
            corrections[word] = spell.candidates(word)

    return corrections

# Function to replace corrected words in the DataFrame
def apply_corrections(data, corrections):
    for col in data.select_dtypes(include=[object]).columns:
        for wrong_word, correct_word in corrections.items():
            data[col] = data[col].str.replace(f"\\b{wrong_word}\\b", correct_word, regex=True)
    return data

# App Interface
st.title("Excel Spelling Error Correction Tool")

uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])
if uploaded_file:
    data = pd.read_excel(uploaded_file)
    st.write("### Original File Data:")
    st.dataframe(data)

    corrections = correct_spelling(data)
    st.write("### Identified Spelling Errors:")
    if corrections:
        correction_dict = {}
        for word, suggestions in corrections.items():
            st.write(f"Misspelled Word: `{word}`")
            suggestion = st.selectbox(f"Select a correction for `{word}`", suggestions, key=word)
            correction_dict[word] = suggestion

        if st.button("Apply Corrections"):
            updated_data = apply_corrections(data, correction_dict)
            st.write("### Corrected Data:")
            st.dataframe(updated_data)

            # Save to Excel
            output = BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                updated_data.to_excel(writer, index=False, sheet_name="Corrected Data")
            output.seek(0)

            st.download_button(
                label="Download Corrected Excel File",
                data=output,
                file_name="corrected_file.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.write("No spelling errors found!")
