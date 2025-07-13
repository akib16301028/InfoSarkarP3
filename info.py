import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import tkinter as tk
from tkinter import filedialog
import os

def read_data_file(file_path):
    """Read data from either Excel or CSV file"""
    if file_path.endswith('.csv'):
        return pd.read_csv(file_path)
    else:  # Assume Excel for .xlsx, .xls
        return pd.read_excel(file_path)

def compare_data_files():
    # Create file dialog
    root = tk.Tk()
    root.withdraw()
    
    # Ask for first file
    file1_path = filedialog.askopenfilename(
        title="Select First File (Sheet1/File1)",
        filetypes=[("Excel/CSV files", "*.xlsx *.xls *.csv"), ("All files", "*.*")]
    )
    if not file1_path:
        print("No file selected. Exiting...")
        return
    
    # Ask for second file
    file2_path = filedialog.askopenfilename(
        title="Select Second File (Sheet2/File2)",
        filetypes=[("Excel/CSV files", "*.xlsx *.xls *.csv"), ("All files", "*.*")]
    )
    if not file2_path:
        print("No second file selected. Exiting...")
        return
    
    try:
        # Read both files
        sheet1 = read_data_file(file1_path)
        sheet2 = read_data_file(file2_path)
        
        # Check if required columns exist
        required_columns = ['Source NE', 'Destination NE', 'Source Port', 'Destination Port']
        for col in required_columns:
            if col not in sheet1.columns:
                raise ValueError(f"Column '{col}' not found in first file")
            if col not in sheet2.columns:
                raise ValueError(f"Column '{col}' not found in second file")
        
        # Merge for comparison
        merged = pd.merge(
            sheet1, 
            sheet2, 
            on=['Source NE', 'Destination NE'], 
            how='outer', 
            indicator='NE_Match',
            suffixes=('_File1', '_File2')
        )
        
        # Convert match indicator
        merged['NE_Match'] = merged['NE_Match'].map({
            'left_only': 'Mismatched (Only in File1)',
            'right_only': 'Mismatched (Only in File2)',
            'both': 'Matched'
        })
        
        # Compare ports
        def compare_ports(row):
            if row['NE_Match'] != 'Matched':
                return 'N/A'
            
            source_match = row['Source Port_File1'] == row['Source Port_File2']
            dest_match = row['Destination Port_File1'] == row['Destination Port_File2']
            
            if source_match and dest_match:
                return 'Ports Matched'
            else:
                mismatches = []
                if not source_match:
                    mismatches.append(f"Source Port File2: {row['Source Port_File2']}")
                if not dest_match:
                    mismatches.append(f"Destination Port File2: {row['Destination Port_File2']}")
                return ', '.join(mismatches)
        
        merged['Port_Comparison'] = merged.apply(compare_ports, axis=1)
        
        # Save results
        output_dir = os.path.dirname(file1_path)
        output_name = f"comparison_result_{os.path.splitext(os.path.basename(file1_path))[0]}.xlsx"
        output_path = os.path.join(output_dir, output_name)
        
        # Save to Excel with formatting
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            merged.to_excel(writer, index=False)
            
            # Apply formatting
            workbook = writer.book
            worksheet = writer.sheets['Sheet1']
            
            red_fill = PatternFill(start_color='FFFF0000', end_color='FFFF0000', fill_type='solid')
            
            for row in range(2, worksheet.max_row + 1):
                ne_match = worksheet.cell(row=row, column=merged.shape[1] - 1).value
                port_match = worksheet.cell(row=row, column=merged.shape[1]).value
                
                if 'Mismatched' in str(ne_match):
                    for col in range(1, worksheet.max_column + 1):
                        worksheet.cell(row=row, column=col).fill = red_fill
                elif port_match != 'Ports Matched' and port_match != 'N/A':
                    for col in range(1, worksheet.max_column + 1):
                        if '_File2' in str(worksheet.cell(row=1, column=col).value):
                            worksheet.cell(row=row, column=col).fill = red_fill
        
        print(f"Comparison completed. Results saved to: {output_path}")
    
    except Exception as e:
        print(f"Error: {str(e)}")

if __name__ == "__main__":
    compare_data_files()
