import pandas as pd
import json
import os
from collections import Counter

def is_boolean_field(series):
    """Check if values match boolean patterns like Yes/No, Y/N"""
    unique_values = set(str(val).strip().lower() for val in series.dropna().unique())
    boolean_sets = [
        {'yes', 'no'},
        {'y', 'n'},
        {'yes'},
        {'no'},
        {'y'},
        {'n'},
        {'true'},
        {'false'}
    ]
    return any(unique_values == s for s in boolean_sets)

def is_likely_choice_field(series):
    """Determine if a column is a choice based on length, unique values, total rows."""
    unique_values = series.dropna().unique()
    total_rows = len(series.dropna())
    
    if len(unique_values) == 0:
        return False
        
    unique_str_values = [str(val).strip() for val in unique_values]
    max_length = max(len(val) for val in unique_str_values)
    unique_ratio = len(unique_values) / total_rows
    
    conditions = [
        len(unique_values) <= 20,
        max_length <= 50,
        unique_ratio < 0.1,
        total_rows >= 10
    ]
    
    return all(conditions)

def analyse_column_types(excel_file):
    all_sheets = {}
    excel = pd.ExcelFile(excel_file)
    
    for sheet_name in excel.sheet_names:
        df = pd.read_excel(excel_file, sheet_name=sheet_name)
        column_types = {}
        column_counts = Counter()
        
        for column in df.columns:
            base_name = column
            column_counts[base_name] += 1
            unique_column_name = f"{base_name}_{column_counts[base_name]}" if column_counts[base_name] > 1 else base_name
            
            dtype = df[column].dtype
            
            if str(dtype) == 'object':
                if is_boolean_field(df[column]):
                    column_types[unique_column_name] = "boolean"
                elif is_likely_choice_field(df[column]):
                    unique_values = sorted(df[column].dropna().unique())
                    column_types[unique_column_name] = f"choice ({', '.join(str(val) for val in unique_values)})"
                else:
                    max_length = df[column].astype(str).str.len().max()
                    has_newlines = df[column].astype(str).str.contains('\n').any()
                    
                    if has_newlines:
                        column_types[unique_column_name] = "multiline text"
                    elif max_length > 255:
                        column_types[unique_column_name] = "long text"
                    else:
                        column_types[unique_column_name] = "short text"
            else:
                column_types[unique_column_name] = str(dtype)
        
        all_sheets[sheet_name] = column_types
    
    base_name = os.path.splitext(excel_file)[0]
    json_file = f"{base_name}.json"

    with open(json_file, 'w') as f:
        json.dump(all_sheets, f, indent=2)

    return json.dumps(all_sheets, indent=2)

if __name__ == "__main__":
    excel_file = "sample2.xlsx"
    output_file = analyse_column_types(excel_file)
    print(output_file)