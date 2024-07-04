import streamlit as st
import pandas as pd
import openpyxl
import matplotlib.pyplot as plt
import openai

# OpenAI API key
openai.api_key = 'YOUR_OPENAI_API_KEY'

# Load the Excel file
@st.cache
def load_data(file):
    df = pd.read_excel(file)
    return df

# Function to generate plot description using OpenAI
def generate_plot_description(df):
    prompt = f"Generate a plot description for the following data:\n{df.head()}"
    response = openai.Completion.create(
        engine="text-davinci-003",
        prompt=prompt,
        max_tokens=100
    )
    return response.choices[0].text.strip()

# Streamlit app
st.title('Excel Plotting Agent')

uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")

if uploaded_file is not None:
    df = load_data(uploaded_file)
    st.write("Data Preview:")
    st.dataframe(df)

    # Generate plot description
    plot_description = generate_plot_description(df)
    st.write("Plot Description:")
    st.write(plot_description)

    # Plotting
    st.write("Plot:")
    plt.figure(figsize=(10, 6))
    for column in df.columns[1:]:
        plt.plot(df[df.columns[0]], df[column], label=column)
    plt.xlabel(df.columns[0])
    plt.ylabel('Values')
    plt.title('Line Plot')
    plt.legend()
    st.pyplot(plt)

    # Allow user to download the plot
    plt.savefig('plot.png')
    st.download_button(
        label="Download Plot",
        data=open('plot.png', 'rb').read(),
        file_name='plot.png',
        mime='image/png'
    )
