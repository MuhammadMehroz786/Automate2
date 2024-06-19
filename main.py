import openai
import pandas as pd
import time
import streamlit as st
from io import BytesIO

st.set_page_config(page_title="CEFR Level Sheet")

def main():
    api_key = st.text_input("Enter your OpenAI API Key", type="password")
    st.header("Ask your Sheet")

    file_path = st.file_uploader("Upload your Excel Sheet", type="xlsx")

    if file_path is not None and api_key:
        df = pd.read_excel(file_path)
        words = df['Table 1'].dropna().tolist()[0:800]

        openai.api_key = api_key

        def get_cefr_level(word):
            retries = 0
            max_retries = 5
            backoff_factor = 1.5

            while retries < max_retries:
                try:
                    response = openai.ChatCompletion.create(
                        model="gpt-3.5-turbo",
                        messages=[
                            {"role": "system", "content": "You are a helpful assistant knowledgeable about CEFR levels. Only give a one-word answer which is the CEFR level of the word asked."},
                            {"role": "user", "content": f"What is the CEFR level of the word '{word}'?"}
                        ],
                        max_tokens=10
                    )
                    return response['choices'][0]['message']['content'].strip()
                except openai.error.RateLimitError:
                    wait = (backoff_factor ** retries) * 2
                    time.sleep(wait)
                    retries += 1
                except Exception as e:
                    print(f"An error occurred: {e}")
                    break

            return "Rate limit exceeded, try again later."

        word_cefr_levels = []

        progress_bar = st.progress(0)
        total_words = len(words)

        for i, word in enumerate(words):
            try:
                cefr_level = get_cefr_level(word)
                word_cefr_levels.append((word, cefr_level))
                time.sleep(1)
            except Exception as e:
                print(f"An error occurred for word '{word}': {e}")
                word_cefr_levels.append((word, 'Error'))
            
            progress_bar.progress((i + 1) / total_words)

        cefr_df = pd.DataFrame(word_cefr_levels, columns=['Word', 'CEFR Level'])

        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            cefr_df.to_excel(writer, index=False)
        excel_data = output.getvalue()

        st.success('CEFR levels processed successfully.')

        st.download_button(
            label="Download CEFR Data as Excel",
            data=excel_data,
            file_name="cefr_words.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

if __name__ == "__main__":
    main()
