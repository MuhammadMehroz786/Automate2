import streamlit as st
import pandas as pd
import openai
import os


st.set_page_config(page_title="ASK YOUR Sheet")

def main():
    api_key = st.text_input("Enter your OpenAI API Key", type="password")

    st.header("Ask your Sheet")

    file_path = st.file_uploader("Upload your Excel Sheet", type="xlsx")

    if file_path is not None:
        try:
            df = pd.read_excel(file_path)
        except Exception as e:
            st.write(f"Error reading the Excel file: {e}")
            return

        # Convert the data into a single string
        text = ' '.join(df.stack().astype(str))

        if api_key:
            openai.api_key = api_key
        else:
            st.write("Please enter your OpenAI API Key.")
            return

        # Define your messages
        messages = [
            {"role": "system", "content": "You are a helpful assistant and very very knowledgeable about CEFR levels. and only give a two word answer for each word, which is the word it self and the CEFR level for each of the words entered and do this very very carefully so that there are no Mistakes do the same and give the answer in the form Abandon B2 Ability B1 Able A2 Abortion B2 About A1 etc"},
            {"role": "user", "content": text}
        ]

        # Call the OpenAI API to get the completion
        try:
            response = openai.ChatCompletion.create(
                model="gpt-3.5-turbo",
                messages=messages,
                max_tokens=1000
            )
        except Exception as e:
            st.write(f"Error communicating with OpenAI API: {e}")
            return

        # Get the answer from the API response
        answer = response.choices[0]['message']['content']

        # Split the answer into words and their CEFR levels
        words_cefr_levels = answer.split()

        # Create a list to store tuples of (word, CEFR level)
        word_cefr_pairs = []
        for i in range(0, len(words_cefr_levels), 2):
            word = words_cefr_levels[i]
            cefr_level = words_cefr_levels[i + 1]
            word_cefr_pairs.append((word, cefr_level))

        # Create a DataFrame from the list of tuples
        cefr_df = pd.DataFrame(word_cefr_pairs, columns=['Word', 'CEFR Level'])

        # Save DataFrame to Excel
        output_file_path = '/path/to/output_file.xlsx'  # Replace with your desired output file path
        try:
            cefr_df.to_excel(output_file_path, index=False)
            st.success(f"CEFR levels saved to {output_file_path}")
        except Exception as e:
            st.write(f"Error saving Excel file: {e}")

if __name__ == '__main__':
    main()
