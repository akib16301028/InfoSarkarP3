import streamlit as st
import pandas as pd
from io import StringIO

st.title("NE & Port Matching Checker")

uploaded_file = st.file_uploader("Upload Excel or CSV file (with Sheet1 and Sheet2)", type=["xlsx", "xls", "csv"])

if uploaded_file:
    file_type = uploaded_file.name.split(".")[-1]
    
    if file_type in ["xlsx", "xls"]:
        sheet1 = pd.read_excel(uploaded_file, sheet_name=0)
        sheet2 = pd.read_excel(uploaded_file, sheet_name=1)
    elif file_type == "csv":
        df = pd.read_csv(uploaded_file)
        st.warning("CSV files do not support multiple sheets. Comparing the same CSV twice as Sheet1 and Sheet2.")
        sheet1 = df.copy()
        sheet2 = df.copy()
    else:
        st.error("Unsupported file format.")
        st.stop()

    required_columns = ['Source NE', 'Destination NE', 'Source Port', 'Destination Port']
    if not all(col in sheet1.columns and col in sheet2.columns for col in required_columns):
        st.error(f"Missing required columns in one or both sheets: {required_columns}")
        st.stop()

    comparison_df = sheet1[required_columns].copy()
    comparison_df['NE Match'] = (sheet1['Source NE'] == sheet2['Source NE']) & (sheet1['Destination NE'] == sheet2['Destination NE'])
    comparison_df['Port Match'] = (sheet1['Source Port'] == sheet2['Source Port']) & (sheet1['Destination Port'] == sheet2['Destination Port'])

    comparison_df['Status'] = comparison_df['NE Match'].apply(lambda x: 'Matched' if x else 'Mismatched')

    # Add details of mismatched ports
    comparison_df['Mismatched Ports from Sheet2'] = ''
    mask_port_mismatch = ~comparison_df['Port Match']
    comparison_df.loc[mask_port_mismatch, 'Mismatched Ports from Sheet2'] = \
        'Src: ' + sheet2.loc[mask_port_mismatch, 'Source Port'].astype(str) + \
        ', Dst: ' + sheet2.loc[mask_port_mismatch, 'Destination Port'].astype(str)

    # Style for red highlighting mismatches
    def highlight_mismatches(row):
        color = 'background-color: red'
        default = ''
        return [
            color if not row['NE Match'] and col in ['Source NE', 'Destination NE'] else default
            for col in row.index
        ]

    styled_df = comparison_df.style.apply(highlight_mismatches, axis=1)

    st.write("### Comparison Result:")
    st.dataframe(styled_df, use_container_width=True)

    csv = comparison_df.to_csv(index=False).encode('utf-8')
    st.download_button("Download Result CSV", csv, "comparison_result.csv", "text/csv")

