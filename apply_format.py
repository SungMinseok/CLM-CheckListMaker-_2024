import pandas as pd
from openpyxl import load_workbook

def apply_formatting(template_file, output_df):
    # Load template.xlsx
    template_wb = load_workbook(template_file)

    # Access 'A' sheet in template.xlsx
    template_sheet = template_wb['A']

    # Create a new DataFrame for the formatted output
    output_processed_df = pd.DataFrame(columns=output_df.columns)

    # Iterate through columns and apply formatting from template to output_df
    for col in output_df.columns:
        # Get column index (1-based)
        col_index = output_df.columns.get_loc(col) + 1

        # Apply formatting from the template to the corresponding column in output_df
        for row_idx, cell in enumerate(template_sheet[col], start=1):
            output_df.at[row_idx - 1, col] = cell.value  # Set the value
            output_processed_df.at[row_idx - 1, col] = cell.value  # Set the value for processed DataFrame

            # Copy styles from the template to the output_processed_df
            output_processed_df.at[row_idx - 1, col] = output_df.at[row_idx - 1, col]
            output_processed_df.at[row_idx - 1, col].number_format = cell.number_format
            output_processed_df.at[row_idx - 1, col].font = cell.font
            output_processed_df.at[row_idx - 1, col].alignment = cell.alignment

    # Save the changes to the template.xlsx
    template_wb.save(template_file)

    return output_processed_df
