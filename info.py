import streamlit as st
import pandas as pd
from io import BytesIO

st.title("Network Connection Analyzer & Fixer")

def normalize_ne_pair(source, dest):
    """Create normalized NE pair (treats A→B and B→A as same)"""
    return tuple(sorted([str(source), str(dest)]))

def process_sheets(sheet1, sheet2):
    """Process both sheets and return analysis results"""
    # Add normalized NE pairs
    sheet1['NE_Pair'] = sheet1.apply(lambda x: normalize_ne_pair(x['Source NE'], x['Destination NE']), axis=1)
    sheet2['NE_Pair'] = sheet2.apply(lambda x: normalize_ne_pair(x['Source NE'], x['Destination NE']), axis=1)
    
    # Create fixed version of Sheet1
    fixed_sheet1 = sheet1.copy()
    
    # Find entries in Sheet2 missing from Sheet1
    missing_in_sheet1 = sheet2[~sheet2['NE_Pair'].isin(sheet1['NE_Pair'])]
    
    # Find interface mismatches (where NE pair exists but ports differ)
    interface_mismatches = []
    
    for idx, row in sheet1.iterrows():
        ne_pair = normalize_ne_pair(row['Source NE'], row['Destination NE'])
        sheet2_matches = sheet2[sheet2['NE_Pair'] == ne_pair]
        
        if not sheet2_matches.empty:
            sheet2_row = sheet2_matches.iloc[0]
            
            # Check if direction is reversed
            is_reversed = (sheet2_row['Source NE'] != row['Source NE'])
            
            # Get correct ports from Sheet2 (accounting for direction)
            correct_src = sheet2_row['Destination Port'] if is_reversed else sheet2_row['Source Port']
            correct_dst = sheet2_row['Source Port'] if is_reversed else sheet2_row['Destination Port']
            
            # Check for interface mismatches
            if (row['Source Port'] != correct_src) or (row['Destination Port'] != correct_dst):
                interface_mismatches.append({
                    'Sheet1 Source NE': row['Source NE'],
                    'Sheet1 Destination NE': row['Destination NE'],
                    'Sheet1 Source Port': row['Source Port'],
                    'Sheet1 Destination Port': row['Destination Port'],
                    'Correct Source Port': correct_src,
                    'Correct Destination Port': correct_dst
                })
                
                # Update fixed Sheet1 with correct ports
                fixed_sheet1.at[idx, 'Source Port'] = correct_src
                fixed_sheet1.at[idx, 'Destination Port'] = correct_dst
    
    # Create fixed Sheet1 by adding missing connections
    fixed_sheet1 = pd.concat([fixed_sheet1, missing_in_sheet1]).drop_duplicates('NE_Pair')
    
    return {
        'Original_Sheet1': sheet1.drop(columns=['NE_Pair']),
        'Original_Sheet2': sheet2.drop(columns=['NE_Pair']),
        'Missing_In_Sheet1': missing_in_sheet1.drop(columns=['NE_Pair']),
        'Interface_Mismatches': pd.DataFrame(interface_mismatches),
        'Fixed_Sheet1': fixed_sheet1.drop(columns=['NE_Pair'])
    }

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
        
        # Process sheets
        results = process_sheets(sheet1, sheet2)
        
        # Create Excel file in memory
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            for sheet_name, df in results.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)
        
        # Show summary
        st.success("Analysis complete!")
        st.write(f"Found {len(results['Missing_In_Sheet1'])} missing connections in Sheet1")
        st.write(f"Found {len(results['Interface_Mismatches'])} interface mismatches")
        
        # Show previews
        with st.expander("Preview Missing Connections"):
            st.dataframe(results['Missing_In_Sheet1'])
        
        with st.expander("Preview Interface Mismatches"):
            st.dataframe(results['Interface_Mismatches'])
        
        with st.expander("Preview Fixed Sheet1"):
            st.dataframe(results['Fixed_Sheet1'].head())
        
        # Download button
        st.download_button(
            label="Download Full Analysis Report",
            data=output.getvalue(),
            file_name="network_connection_analysis.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    except Exception as e:
        st.error(f"Error processing file: {str(e)}")
