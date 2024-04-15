import pandas as pd

def compare_and_mark_unmatched(input_excel, output_excel, comparison_excel, key_column):
    # Read the input Excel files
    df_input = pd.read_excel(input_excel)
    df_comparison = pd.read_excel(comparison_excel)
    
    # Check if the specified key column exists in both DataFrames
    if key_column not in df_input.columns or key_column not in df_comparison.columns:
        print(f"Column '{key_column}' not found in one of the input Excel files.")
        return
    
    # Perform comparison and calculate verification result
    verification_result = []
    for index, row in df_input.iterrows():
        key_value = row[key_column]
        corresponding_rows = df_comparison[df_comparison[key_column] == key_value]
        if not corresponding_rows.empty:
            verification_result.append('Match')
        else:
            verification_result.append('Not Matching')

    # Add verification result to input DataFrame
    df_input['Verification Result'] = verification_result

    # Write updated DataFrame to the output Excel file
    df_input.to_excel(output_excel, index=False)

# Example usage:
input_excel = 'output.xlsx'  # Input Excel file name
output_excel = 'Comparison_Output_Excel.xlsx'  # Output Excel file name
comparison_excel = 'ProdExcel.xlsx'  # Comparison Excel file name
key_column = 'VALORNR'  # Key column name

compare_and_mark_unmatched(input_excel, output_excel, comparison_excel, key_column)
