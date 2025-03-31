import os
import sys
import re
import pandas as pd
from datetime import datetime
import shutil

def parse_rpt_file(file_path):
    """Parse an RPT file and extract the table data."""
    print(f"Parsing file: {file_path}")
    
    try:
        with open(file_path, 'r', encoding='utf-8-sig') as file:
            content = file.read()
    except Exception as e:
        print(f"Error reading file: {e}")
        return None
    
    # Find the header row (starting with column names)
    header_match = re.search(r'^\s*(\S+\s+)+$', content, re.MULTILINE)
    if not header_match:
        print("Could not find header row in the file.")
        return None
    
    # Extract column headers
    header_line = header_match.group(0).strip()
    
    # Find column positions by looking at the dash line below headers
    dash_line_match = re.search(r'^\s*([-]+\s+)+', content, re.MULTILINE)
    if not dash_line_match:
        print("Could not find dash line below headers.")
        return None
    
    dash_line = dash_line_match.group(0)
    
    # Clean up the dash line to ensure consistent spacing
    dash_line = dash_line.replace('\t', ' ')
    
    # Find column positions
    positions = []
    for match in re.finditer(r'([-]+)', dash_line):
        positions.append((match.start(), match.end()))
    
    if not positions:
        print(f"Error: Could not find column separators in the dash line.")
        print(f"Dash line: {dash_line}")
        return None
    
    print(f"Found {len(positions)} columns based on dash line separators.")
    
    # Clean up the header line
    header_line = header_line.replace('\t', ' ')
    
    # Alternative approach to extract column names
    column_names = []
    
    # Add the first column
    column_names.append(header_line[:min(positions[0][1], len(header_line))].strip())
    
    # For each dash section, extract the corresponding header
    for i in range(1, len(positions)):
        start = positions[i-1][1]
        end = positions[i][0]
        
        if start < len(header_line):
            col_name = header_line[start:min(end, len(header_line))].strip()
            if not col_name:
                col_name = f"Column{i+1}"
            column_names.append(col_name)
        else:
            column_names.append(f"Column{i+1}")
    
    # Add the last column if there's content after the last dash segment
    if len(header_line) > positions[-1][1]:
        last_col = header_line[positions[-1][1]:].strip()
        if last_col:
            column_names.append(last_col)
    
    print(f"Extracted column names: {len(column_names)} columns")
    
    # Find data rows
    data_section = content[dash_line_match.end():]
    
    # Find the line that indicates end of data (like "(6 rows affected)")
    end_match = re.search(r'\(\d+\s+rows? affected\)', data_section)
    if end_match:
        data_section = data_section[:end_match.start()]
    
    # Parse data rows with improved algorithm
    data_rows = []
    for line in data_section.strip().split('\n'):
        if not line.strip() or "Completion time:" in line:
            continue
        
        # Clean the line
        line = line.replace('\t', ' ')
        
        row_data = []
        
        # Extract first column
        first_col_end = min(positions[0][1], len(line))
        value = line[:first_col_end].strip()
        if value == "NULL":
            value = None
        row_data.append(value)
        
        # Extract middle columns
        for j in range(1, len(positions)):
            try:
                start = positions[j-1][1]
                if start < len(line):
                    end = positions[j][0]
                    width = min(end - start, len(line) - start)
                    
                    value = line[start:start+width].strip()
                    if value == "NULL":
                        value = None
                    row_data.append(value)
                else:
                    row_data.append(None)
            except Exception as e:
                print(f"Warning: Error extracting data for column {j}: {e}")
                row_data.append(None)
        
        # Extract last column if it exists
        if len(line) > positions[-1][1]:
            try:
                value = line[positions[-1][1]:].strip()
                if value == "NULL":
                    value = None
                row_data.append(value)
            except Exception as e:
                print(f"Warning: Error extracting last column: {e}")
                row_data.append(None)
        
        # Make sure we have the correct number of columns
        while len(row_data) < len(column_names):
            row_data.append(None)
        
        # Only add rows that have data
        if any(val is not None and val != "" for val in row_data):
            data_rows.append(row_data)
    
    print(f"Parsed {len(data_rows)} data rows")
    
    # Create DataFrame
    df = pd.DataFrame(data_rows, columns=column_names)
    print(f"Successfully extracted {len(df)} rows with {len(column_names)} columns.")
    return df

def create_folders():
    """Create the 'excel' and 'rpt' folders if they don't exist."""
    current_dir = os.getcwd()
    
    # Create 'excel' folder for output
    excel_folder = os.path.join(current_dir, 'excel')
    if not os.path.exists(excel_folder):
        print(f"Creating 'excel' folder for output files...")
        os.makedirs(excel_folder)
    
    # Create 'rpt' folder for input
    rpt_folder = os.path.join(current_dir, 'rpt')
    if not os.path.exists(rpt_folder):
        print(f"Creating 'rpt' folder for input files...")
        os.makedirs(rpt_folder)
    
    return excel_folder, rpt_folder

def convert_to_excel(df, input_path):
    """Convert DataFrame to Excel file in the 'excel' folder."""
    try:
        # Create folders
        excel_folder, _ = create_folders()
        
        # Create output path in 'excel' folder
        output_filename = os.path.splitext(os.path.basename(input_path))[0] + '.xlsx'
        output_path = os.path.join(excel_folder, output_filename)
        
        # Create Excel writer
        writer = pd.ExcelWriter(output_path, engine='openpyxl')
        
        # Write DataFrame to Excel
        df.to_excel(writer, index=False, sheet_name='RPT Data')
        
        # Auto-adjust columns' width
        for column in df:
            column_width = max(df[column].astype(str).map(len).max(), len(column))
            col_idx = df.columns.get_loc(column)
            writer.sheets['RPT Data'].column_dimensions[chr(65 + col_idx)].width = column_width + 2
        
        # Save the Excel file
        writer.close()
        print(f"Excel file created successfully: {output_path}")
        return True, output_path
    except Exception as e:
        print(f"Error creating Excel file: {e}")
        return False, None

def main():
    # Create necessary folders
    _, rpt_folder = create_folders()
    
    # Check if file path is provided
    if len(sys.argv) < 2:
        print("Usage: python rpt_to_excel.py <path_to_rpt_file>")
        print("Or drag and drop an RPT file onto this script.")
        return
    
    # Get file path from command line argument
    file_path = sys.argv[1]
    
    # Check if file exists
    if not os.path.isfile(file_path):
        print(f"File not found: {file_path}")
        return
    
    # Check if file is an RPT file
    if not file_path.lower().endswith('.rpt'):
        print(f"Not an RPT file: {file_path}")
        return
    
    # Copy the file to the rpt folder if it's not already there
    if not os.path.dirname(os.path.abspath(file_path)) == os.path.abspath(rpt_folder):
        rpt_file_in_folder = os.path.join(rpt_folder, os.path.basename(file_path))
        try:
            shutil.copy2(file_path, rpt_file_in_folder)
            print(f"Copied {file_path} to {rpt_folder}")
            # Use the copied file for processing
            file_path = rpt_file_in_folder
        except Exception as e:
            print(f"Warning: Could not copy file to rpt folder: {e}")
    
    # Parse RPT file
    df = parse_rpt_file(file_path)
    if df is None:
        print("Failed to parse RPT file.")
        return
    
    # Convert to Excel
    success, output_path = convert_to_excel(df, file_path)
    if success:
        print(f"Conversion completed: {file_path} -> {output_path}")
    else:
        print("Conversion failed.")

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
    
    # Keep console window open if script was double-clicked
    if len(sys.argv) <= 1:
        input("Press Enter to exit...")
