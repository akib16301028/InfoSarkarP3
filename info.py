import streamlit as st
import pandas as pd
from io import StringIO

st.title("NE & Port Matching Checker")

uploaded_file = st.file_uploader("Upload Excel file with Sheet1 and Sheet2", type=["xlsx", "xls"])

if uploaded_file:
    try:
        # Read both sheets
        sheet1 = pd.read_excel(uploaded_file, sheet_name=0)
        sheet2 = pd.read_excel(uploaded_file, sheet_name=1)
        
        # Clean data - replace N/A values
        sheet1 = sheet1.replace(['N/A', 'NA', 'n/a', 'N/a', ''], pd.NA)
        sheet2 = sheet2.replace(['N/A', 'NA', 'n/a', 'N/a', ''], pd.NA)
        
        # Check required columns in both sheets
        required_columns = ['Source NE', 'Destination NE', 'Source Port', 'Destination Port']
        if not all(col in sheet1.columns and col in sheet2.columns for col in required_columns):
            st.error(f"Both sheets must have these columns: {required_columns}")
            st.stop()
            
        # Create comparison dataframe starting with Sheet1 data
        comparison_df = sheet1[required_columns].copy()
        
        # Add Status column - check if NE pairs match
        comparison_df['Status'] = sheet1.apply(
            lambda row: 'Matched' if ((sheet2['Source NE'] == row['Source NE']) & 
                                     (sheet2['Destination NE'] == row['Destination NE'])).any()
                        else 'Mismatched', 
            axis=1
        )
        
        # For matched rows, compare ports and show Sheet2 ports if different
        comparison_df['Sheet2 Ports'] = ''
        
        for idx, row in sheet1.iterrows():
            # Find matching row in Sheet2
            match = sheet2[
                (sheet2['Source NE'] == row['Source NE']) & 
                (sheet2['Destination NE'] == row['Destination NE'])
            ]
            
            if not match.empty:
                sheet2_ports = f"Src: {match.iloc[0]['Source Port']}, Dst: {match.iloc[0]['Destination Port']}"
                
                # Check if ports match
                port_match = (row['Source Port'] == match.iloc[0]['Source Port']) and \
                             (row['Destination Port'] == match.iloc[0]['Destination Port'])
                
                if not port_match:
                    comparison_df.at[idx, 'Sheet2 Ports'] = sheet2_ports
        
        # Style function to highlight mismatches
        def highlight_mismatches(row):
            styles = [''] * len(row)
            if row['Status'] == 'Mismatched':
                styles[0] = 'background-color: red'  # Source NE
                styles[1] = 'background-color: red'  # Destination NE
            return styles
        
        # Apply styling
        styled_df = comparison_df.style.apply(highlight_mismatches, axis=1)
        
        st.write("### Comparison Results:")
        st.dataframe(styled_df, use_container_width=True)
        
        # Download button
        csv = comparison_df.to_csv(index=False).encode('utf-8')
        st.download_button(
            "Download Comparison Results",
            csv,
            "ne_port_comparison.csv",
            "text/csv"
        )
        
    except Exception as e:
        st.error(f"Error processing file: {str(e)}")
