import streamlit as st
import pandas as pd

st.title("NE Pair & Interface Matching Checker")

uploaded_file = st.file_uploader("Upload Excel file with Sheet1 and Sheet2", type=["xlsx", "xls"])

if uploaded_file:
    try:
        # Read both sheets
        sheet1 = pd.read_excel(uploaded_file, sheet_name=0)
        sheet2 = pd.read_excel(uploaded_file, sheet_name=1)
        
        # Clean data
        sheet1 = sheet1.replace(['N/A', 'NA', 'n/a', 'N/a', ''], pd.NA)
        sheet2 = sheet2.replace(['N/A', 'NA', 'n/a', 'N/a', ''], pd.NA)
        
        # Check required columns
        required_cols = ['Source NE', 'Destination NE', 'Source Port', 'Destination Port']
        if not all(col in sheet1.columns and col in sheet2.columns for col in required_cols):
            st.error(f"Both sheets need these columns: {required_cols}")
            st.stop()
        
        # Create comparison DataFrame from Sheet1
        comparison = sheet1[required_cols].copy()
        
        # Add Status column (Matched if NE pair exists in Sheet2)
        comparison['Status'] = comparison.apply(
            lambda row: 'Matched' if ((sheet2['Source NE'] == row['Source NE']) & 
                                     (sheet2['Destination NE'] == row['Destination NE'])).any()
                        else 'Mismatched',
            axis=1
        )
        
        # Add Sheet2 Interface Info (ONLY when interfaces don't match)
        comparison['Sheet2 Interfaces'] = ''
        
        for idx, row in sheet1.iterrows():
            # Find matching row in Sheet2
            match = sheet2[
                (sheet2['Source NE'] == row['Source NE']) & 
                (sheet2['Destination NE'] == row['Destination NE'])
            ]
            
            if not match.empty:
                sheet2_row = match.iloc[0]
                # Only show Sheet2 interfaces if they don't match
                if (row['Source Port'] != sheet2_row['Source Port']) or \
                   (row['Destination Port'] != sheet2_row['Destination Port']):
                    comparison.at[idx, 'Sheet2 Interfaces'] = \
                        f"Src: {sheet2_row['Source Port']}, Dst: {sheet2_row['Destination Port']}"
        
        # Style function to highlight mismatches
        def highlight_mismatches(row):
            styles = [''] * len(comparison.columns)
            if row['Status'] == 'Mismatched':
                # Highlight Source NE and Destination NE in red
                styles[0] = 'background-color: #FFCCCB'  # Source NE
                styles[1] = 'background-color: #FFCCCB'  # Destination NE
            elif row['Sheet2 Interfaces'] != '':
                # Highlight interface columns in yellow when ports don't match
                styles[2] = 'background-color: #FFFF99'  # Source Port
                styles[3] = 'background-color: #FFFF99'  # Destination Port
            return styles
        
        # Apply styling and display
        st.dataframe(
            comparison.style.apply(highlight_mismatches, axis=1),
            use_container_width=True,
            height=400
        )
        
        # Download button
        csv = comparison.to_csv(index=False).encode('utf-8')
        st.download_button("Download Results", csv, "ne_interface_comparison.csv", "text/csv")
        
    except Exception as e:
        st.error(f"Error processing file: {str(e)}")
