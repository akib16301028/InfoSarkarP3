import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import tkinter as tk
from tkinter import filedialog

def compare_excel_sheets():
    # Create a file dialog to select the Excel file
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    
    file_path = filedialog.askopenfilename(
        title="Select Excel File",
        filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
    )
    
    if not file_path:
        print("No file selected. Exiting...")
        return
    
    try:
        # Read both sheets
        sheet1 = pd.read_excel(file_path, sheet_name=0)
        sheet2 = pd.read_excel(file_path, sheet_name=1)
        
        # Check if required columns exist in both sheets
        required_columns = ['Source NE', 'Destination NE', 'Source Port', 'Destination Port']
        for col in required_columns:
            if col not in sheet1.columns or col not in sheet2.columns:
                raise ValueError(f"Column '{col}' not found in one or both sheets")
        
        # Merge the two sheets for comparison
        merged = pd.merge(
            sheet1, 
            sheet2, 
            on=['Source NE', 'Destination NE'], 
            how='outer', 
            indicator='NE_Match',
            suffixes=('_Sheet1', '_Sheet2')
        )
        
        # Convert the indicator to more readable values
        merged['NE_Match'] = merged['NE_Match'].map({
            'left_only': 'Mismatched (Only in Sheet1)',
            'right_only': 'Mismatched (Only in Sheet2)',
            'both': 'Matched'
        })
        
        # Add port comparison columns
        def compare_ports(row):
            if row['NE_Match'] != 'Matched':
                return 'N/A'
            
            source_port_match = row['Source Port_Sheet1'] == row['Source Port_Sheet2']
            dest_port_match = row['Destination Port_Sheet1'] == row['Destination Port_Sheet2']
            
            if source_port_match and dest_port_match:
                return 'Ports Matched'
            else:
                mismatches = []
                if not source_port_match:
                    mismatches.append(f"Source Port Sheet2: {row['Source Port_Sheet2']}")
                if not dest_port_match:
                    mismatches.append(f"Destination Port Sheet2: {row['Destination Port_Sheet2']}")
                return ', '.join(mismatches)
        
        merged['Port_Comparison'] = merged.apply(compare_ports, axis=1)
        
        # Save the comparison results to a new Excel file
        output_path = file_path.replace('.xlsx', '_compared.xlsx')
        merged.to_excel(output_path, index=False)
        
        # Apply formatting to highlight mismatches
        wb = load_workbook(output_path)
        ws = wb.active
        
        # Define red fill for mismatches
        red_fill = PatternFill(start_color='FFFF0000', end_color='FFFF0000', fill_type='solid')
        
        # Apply formatting
        for row in range(2, ws.max_row + 1):
            ne_match = ws.cell(row=row, column=ws.max_column - 1).value
            port_match = ws.cell(row=row, column=ws.max_column).value
            
            # Highlight entire row if NE is mismatched
            if 'Mismatched' in str(ne_match):
                for col in range(1, ws.max_column + 1):
                    ws.cell(row=row, column=col).fill = red_fill
            
            # Highlight port cells if ports are mismatched
            elif port_match != 'Ports Matched' and port_match != 'N/A':
                for col in range(1, ws.max_column + 1):
                    if '_Sheet2' in str(ws.cell(row=1, column=col).value):
                        ws.cell(row=row, column=col).fill = red_fill
        
        # Save the formatted file
        wb.save(output_path)
        
        print(f"Comparison completed successfully. Results saved to: {output_path}")
    
    except Exception as e:
        print(f"An error occurred: {str(e)}")

if __name__ == "__main__":
    compare_excel_sheets()
