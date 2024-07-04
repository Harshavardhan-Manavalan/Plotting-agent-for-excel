import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from io import BytesIO
import numpy as np

# Streamlit app title
st.title("Excel Plotting Agent")

# File uploader
uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx", "xls", "csv"])

if uploaded_file:
    # Read the Excel or CSV file
    if uploaded_file.name.endswith('.csv'):
        df = pd.read_csv(uploaded_file)
    else:
        df = pd.read_excel(uploaded_file)

    # Display the data
    st.write("Data from the uploaded file:")
    st.write(df)

    # Initial analysis report
    st.write("### Initial Analysis Report")
    st.write(f"The dataset contains {df.shape[0]} rows and {df.shape[1]} columns.")
    st.write(f"The columns in the dataset are: {', '.join(df.columns)}.")
    st.write(f"The first few rows of the dataset are shown below:")
    st.write(df.head())

    # Select columns for plotting
    columns = df.columns.tolist()
    x_axis = st.selectbox("Select the column for X-axis", columns)
    y_axis = st.selectbox("Select the column for Y-axis", columns)

    # Select plot type
    plot_types = [
        "Line Plot", "Bar Plot", "Area Plot", "Scatter Plot", 
        "Histogram", "Stacked Column Chart", "Box Plot"
    ]
    plot_type = st.selectbox("Select the type of plot", plot_types)

    if st.button("Generate Plot"):
        fig = plt.figure()
        ax = fig.add_subplot(111)

        if plot_type == "Line Plot":
            ax.plot(df[x_axis], df[y_axis])
            ax.set_xlabel(x_axis)
            ax.set_ylabel(y_axis)
            ax.set_title(f"Line Plot: {y_axis} vs {x_axis}")
        elif plot_type == "Bar Plot":
            ax.bar(df[x_axis], df[y_axis])
            ax.set_xlabel(x_axis)
            ax.set_ylabel(y_axis)
            ax.set_title(f"Bar Plot: {y_axis} vs {x_axis}")
        elif plot_type == "Area Plot":
            ax.fill_between(df[x_axis], df[y_axis])
            ax.set_xlabel(x_axis)
            ax.set_ylabel(y_axis)
            ax.set_title(f"Area Plot: {y_axis} vs {x_axis}")
        elif plot_type == "Scatter Plot":
            ax.scatter(df[x_axis], df[y_axis])
            ax.set_xlabel(x_axis)
            ax.set_ylabel(y_axis)
            ax.set_title(f"Scatter Plot: {y_axis} vs {x_axis}")
        elif plot_type == "Histogram":
            ax.hist(df[y_axis], bins='auto')
            ax.set_xlabel(y_axis)
            ax.set_ylabel("Frequency")
            ax.set_title(f"Histogram: {y_axis}")
        elif plot_type == "Stacked Column Chart":
            df.plot(kind='bar', stacked=True, x=x_axis, y=y_axis, ax=ax)
            ax.set_xlabel(x_axis)
            ax.set_ylabel(y_axis)
            ax.set_title(f"Stacked Column Chart: {y_axis} by {x_axis}")
        elif plot_type == "Box Plot":
            ax.boxplot(df[y_axis])
            ax.set_ylabel(y_axis)
            ax.set_title(f"Box Plot: {y_axis}")

        st.pyplot(fig)

        # Detailed analysis report
        st.write("### Detailed Analysis Report")
        st.write(f"Plot Type: {plot_type}")
        st.write(f"X-axis: {x_axis}, Y-axis: {y_axis}")
        st.write(f"Summary of {y_axis}:")
        st.write(df[y_axis].describe())

        # Save plot to Excel
        output = BytesIO()
        writer = pd.ExcelWriter(output, engine='xlsxwriter')
        df.to_excel(writer, sheet_name='Sheet1', index=False)
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']
        imgdata = BytesIO()
        fig.savefig(imgdata, format='png')
        worksheet.insert_image('E2', 'plot.png', {'image_data': imgdata})
        writer.close()

        # Download button for Excel file
        st.download_button(
            label="Download Excel file with Plot",
            data=output.getvalue(),
            file_name="modified_excel.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
