import streamlit as st
import pandas as pd

st.title("NE Pair Missing Links Analyzer")

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
        
        # Function to create normalized NE pair (sorted to treat A→B and B→A as same)
        def normalize_ne_pair(source, dest):
            return tuple(sorted([str(source), str(dest)]))
        
        # Add normalized NE pairs to both sheets
        sheet1['NE_Pair'] = sheet1.apply(lambda x: normalize_ne_pair(x['Source NE'], x['Destination NE']), axis=1)
        sheet2['NE_Pair'] = sheet2.apply(lambda x: normalize_ne_pair(x['Source NE'], x['Destination NE']), axis=1)
        
        # Find entries in Sheet2 that are missing from Sheet1
        missing_in_sheet1 = sheet2[~sheet2['NE_Pair'].isin(sheet1['NE_Pair'])]
        
        if not missing_in_sheet1.empty:
            st.warning(f"Found {len(missing_in_sheet1)} connections in Sheet2 that are missing in Sheet1")
            
            # Prepare the report of missing connections
            missing_report = missing_in_sheet1[required_cols].copy()
            missing_report['Status'] = 'Missing in Sheet1'
            
            # Highlight all rows in red
            def highlight_missing(row):
                return ['background-color: #FFCCCB'] * len(row)
            
            st.dataframe(
                missing_report.style.apply(highlight_missing, axis=1),
                use_container_width=True,
                height=400
            )
            
            # Generate recommended additions for Sheet1
            st.subheader("Recommended additions for Sheet1")
            additions = missing_in_sheet1[required_cols].copy()
            st.dataframe(additions)
            
            # Download buttons
            csv_missing = missing_report.to_csv(index=False).encode('utf-8')
            csv_additions = additions.to_csv(index=False).encode('utf-8')
            
            col1, col2 = st.columns(2)
            with col1:
                st.download_button(
                    "Download Missing Report",
                    csv_missing,
                    "missing_in_sheet1.csv",
                    "text/csv"
                )
            with col2:
                st.download_button(
                    "Download Recommended Additions",
                    csv_additions,
                    "recommended_additions.csv",
                    "text/csv"
                )
        else:
            st.success("All Sheet2 connections are already present in Sheet1!")
            
    except Exception as e:
        st.error(f"Error processing file: {str(e)}")
