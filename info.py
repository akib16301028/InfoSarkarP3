import streamlit as st
import pandas as pd
from io import StringIO

st.title("NE & Port Matching Checker")

uploaded_file = st.file_uploader("Upload Excel or CSV file", type=["xlsx", "xls", "csv"])

if uploaded_file:
    try:
        # Read the uploaded file
        if uploaded_file.name.endswith(('.xlsx', '.xls')):
            df = pd.read_excel(uploaded_file, sheet_name=0)
        else:
            df = pd.read_csv(uploaded_file)
            
        # Clean the data - replace N/A and empty strings with NaN
        df = df.replace(['N/A', 'NA', 'n/a', 'N/a', ''], pd.NA)
        
        # Check required columns
        required_columns = ['Source NE', 'Destination NE', 'Source Port', 'Destination Port']
        if not all(col in df.columns for col in required_columns):
            st.error(f"Missing required columns. Needed: {required_columns}")
            st.stop()
            
        st.write("### First 10 rows of the uploaded data:")
        st.dataframe(df.head(10))
        
        # Option 1: Compare with another file
        st.subheader("Option 1: Compare with another file")
        compare_file = st.file_uploader("Upload file to compare against", type=["xlsx", "xls", "csv"])
        
        if compare_file:
            try:
                if compare_file.name.endswith(('.xlsx', '.xls')):
                    df_compare = pd.read_excel(compare_file, sheet_name=0)
                else:
                    df_compare = pd.read_csv(compare_file)
                    
                df_compare = df_compare.replace(['N/A', 'NA', 'n/a', 'N/a', ''], pd.NA)
                
                # Check if both dataframes have the same columns
                if not all(col in df_compare.columns for col in required_columns):
                    st.error("Comparison file missing required columns")
                    st.stop()
                    
                # Merge the two dataframes for comparison
                merged = df.merge(
                    df_compare,
                    on=['Source NE', 'Destination NE'],
                    how='outer',
                    suffixes=('_Original', '_Compare'),
                    indicator=True
                )
                
                # Create comparison results
                results = pd.DataFrame({
                    'Source NE': merged['Source NE'],
                    'Destination NE': merged['Destination NE'],
                    'Status': merged['_merge'].map({
                        'left_only': 'Only in Original',
                        'right_only': 'Only in Comparison',
                        'both': 'In Both'
                    }),
                    'Original Source Port': merged['Source Port_Original'],
                    'Original Dest Port': merged['Destination Port_Original'],
                    'Compare Source Port': merged['Source Port_Compare'],
                    'Compare Dest Port': merged['Destination Port_Compare'],
                    'Port Match': (merged['Source Port_Original'] == merged['Source Port_Compare']) & 
                                 (merged['Destination Port_Original'] == merged['Destination Port_Compare'])
                })
                
                st.write("### Comparison Results:")
                
                # Color coding for the table
                def highlight_rows(row):
                    if row['Status'] == 'Only in Original':
                        return ['background-color: #FFCCCB'] * len(row)
                    elif row['Status'] == 'Only in Comparison':
                        return ['background-color: #ADD8E6'] * len(row)
                    elif not row['Port Match']:
                        return ['background-color: #FFFF99'] * len(row)
                    else:
                        return [''] * len(row)
                        
                styled_results = results.style.apply(highlight_rows, axis=1)
                st.dataframe(styled_results, use_container_width=True)
                
                # Download button
                csv = results.to_csv(index=False).encode('utf-8')
                st.download_button(
                    "Download Comparison Results",
                    csv,
                    "network_comparison_results.csv",
                    "text/csv"
                )
                
            except Exception as e:
                st.error(f"Error processing comparison file: {str(e)}")
        
        # Option 2: Find duplicates within the same file
        st.subheader("Option 2: Find duplicate connections in same file")
        if st.button("Check for Duplicates"):
            duplicates = df[df.duplicated(subset=['Source NE', 'Destination NE'], keep=False)]
            
            if not duplicates.empty:
                st.write("### Duplicate Connections Found:")
                st.dataframe(duplicates.sort_values(['Source NE', 'Destination NE']))
                
                # Count duplicates
                dup_counts = duplicates.groupby(['Source NE', 'Destination NE']).size().reset_index(name='Count')
                st.write("### Duplicate Counts:")
                st.dataframe(dup_counts)
                
                # Download duplicates
                csv = duplicates.to_csv(index=False).encode('utf-8')
                st.download_button(
                    "Download Duplicates",
                    csv,
                    "duplicate_connections.csv",
                    "text/csv"
                )
            else:
                st.success("No duplicate connections found!")
                
    except Exception as e:
        st.error(f"Error processing file: {str(e)}")
