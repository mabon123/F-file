import pandas as pd

def consolidate_excel_sheets(input_file, output_file):
    excel_data = pd.ExcelFile(input_file)
    consolidated_data = []

    row_ranges = [
        (56, 177),  # First blog
        (178, 328), # Second blog
        (329, 479), # Third blog
        (480, 520)  # Fourth blog
    ]

    for sheet_name in excel_data.sheet_names:
        if sheet_name.startswith('S'):
            # Read the sheet with header=None to avoid pandas interpreting first row as header
            df = pd.read_excel(input_file, sheet_name=sheet_name, header=None)
            
            try:

                if df.iloc[15, 8] == 0:
                    print(f"Sheet {sheet_name} has no data")
                    continue
                elif df.iloc[15, 8] >= 1:
                    try:
                        # Get values from specific cells in row 2
                        province_value = df.iloc[1, 10]  # K2
                        district = df.iloc[1, 16]       # Q2
                        commune = df.iloc[1, 22]       # W2
                        village = df.iloc[1, 29]      # AC2
                        school = df.iloc[1, 39]      # AM2
                        
                    except IndexError:
                        print(f"Warning: Could not find some cell values in sheet {sheet_name}")
                        province_value = district = commune = village = school = None
                    
                    for start_row, end_row in row_ranges:
                        blog_data = df.iloc[start_row:end_row + 1].copy()
                        blog_data = blog_data.dropna(how='all')
                        
                        if not blog_data.empty:
                            blog_data['SheetName'] = sheet_name
                            blog_data["province"] = province_value
                            blog_data["district"] = district
                            blog_data["commune"] = commune
                            blog_data["village"] = village
                            blog_data["school"] = school
                            consolidated_data.append(blog_data)
            except IndexError:
                print(f"Sheet {sheet_name} is not in the expected format")
                continue


    if consolidated_data:
        final_df = pd.concat(consolidated_data, ignore_index=True)
    else:
        final_df = pd.DataFrame()

    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        final_df.to_excel(writer, index=False, sheet_name='ConsolidatedData')

    print(f"Data has been consolidated into {output_file}")


def consolidate_excel_sheets(self, input_file: str, output_file: str,
    cell_configs: List[Dict[str, str]], row_ranges: List[Dict[str, int]]):
    excel_data = pd.ExcelFile(input_file)
    consolidated_data = []
    
    for sheet_name in excel_data.sheet_names:
        if sheet_name.startswith('S'):
            df = pd.read_excel(input_file, sheet_name=sheet_name, header=None)
            
            # Check condition df.iloc[15, 8] >= 1
            try:
                if not (pd.to_numeric(df.iloc[15, 8], errors='coerce') >= 1):
                    print(f"Skipping sheet {sheet_name}: Condition not met at position (15, 8)")
                    continue
            except Exception as e:
                print(f"Error checking condition in sheet {sheet_name}: {e}")
                continue
            
            # Get values from configured cells
            cell_values = {}
            try:
                for config in cell_configs:
                    col_idx = self.column_to_index(config['column'])
                    row_idx = int(config['row']) - 1  # Convert to 0-based index
                    value = df.iloc[row_idx, col_idx]
                    cell_values[f"{config['column']}{config['row']}_value"] = value
                
                print(f"Found values in {sheet_name}:", cell_values)
                
            except (IndexError, ValueError) as e:
                print(f"Warning: Could not find some cell values in sheet {sheet_name}: {e}")
                continue
            
            # Process each row range
            for range_config in row_ranges:
                start_row = range_config['start']
                end_row = range_config['end']
                
                blog_data = df.iloc[start_row:end_row + 1].copy()
                blog_data = blog_data.dropna(how='all')
                
                if not blog_data.empty:
                    blog_data['SheetName'] = sheet_name
                    # Add all cell values as columns
                    for key, value in cell_values.items():
                        blog_data[key] = value
                    consolidated_data.append(blog_data)
    
    if consolidated_data:
        final_df = pd.concat(consolidated_data, ignore_index=True)
    else:
        final_df = pd.DataFrame()
    
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        final_df.to_excel(writer, index=False, sheet_name='ConsolidatedData')
# Example usage
input_file = "F2_ស្ថិតិមធ្យមសិក្សា_សម្រាប់ការិ_អយក_2025 (2).xlsx"
output_file = "now copy.xlsx"
consolidate_excel_sheets(input_file, output_file)