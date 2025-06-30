import pandas as pd
from datetime import datetime
import os

# File paths
excel_file = r'Recent Data Extract Example - Copy.xlsx'
output_file = r'Output_Analysis_Projected_Below.xlsx'
current_date = datetime.today().date()

# Load and clean all sheets
def load_and_select_columns(file):
    all_sheets = pd.read_excel(file, sheet_name=None)
    cleaned_rows = []

    for sheet_df in all_sheets.values():
        df = sheet_df[[
            'Order', 'Line', 'Item', 'Order Date', 'Name',
            'Item Description', 'Customer Item', 'Qty Ordered',
            'U/M', 'Unit Price', 'Extended Price', 'Dock Date'
        ]].copy()

        df['Order Date'] = pd.to_datetime(df['Order Date']).dt.date
        df['Dock Date'] = pd.to_datetime(df['Dock Date']).dt.date
        cleaned_rows.append(df)

    # Combine all sheets into one DataFrame
    return pd.concat(cleaned_rows, ignore_index=True)

# Categorize rows based on the start of the 'Item' value
def split_by_item_prefix(df):
    df['Category'] = df['Item'].astype(str).str.extract(r'^(999|NRE|ENG)', expand=False)

    # Drop rows that don't match any category
    df = df.dropna(subset=['Category'])

    # Split into category-wise DataFrames
    return {
        category: df_group.drop(columns='Category')
        for category, df_group in df.groupby('Category')
    }

# Create projection table per category
def create_monthly_projection(df):
    projected_data = []
    for month_num in range(1, 13):
        month_year = datetime(current_date.year, month_num, 1).strftime('%b - %Y')
        month_data = df[df['Dock Date'].apply(lambda x: x.month == month_num and x.year == current_date.year)]
        total = month_data['Extended Price'].sum()
        projected_data.append([total, month_year, month_num])
    return pd.DataFrame(projected_data, columns=['Projected Below', 'Month', 'Month #'])

# Save results to Excel with overdue highlighting
def save_to_excel(original_dfs, projections, output_file):
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        workbook = writer.book
        red_format = workbook.add_format({'bg_color': '#FFC7CE'})

        for category, df in original_dfs.items():
            # Write original data
            data_sheet = f"{category}_Data"
            df.to_excel(writer, sheet_name=data_sheet, index=False)
            ws_data = writer.sheets[data_sheet]

            for i, dock_date in enumerate(df['Dock Date']):
                if dock_date < current_date:
                    ws_data.set_row(i + 1, None, red_format)

            # Write projection table
            proj_df = projections[category]
            proj_df.to_excel(writer, sheet_name=f"{category}_Projection", index=False)

    print(f"Excel report generated at: {output_file}")

# === MAIN EXECUTION ===
full_df = load_and_select_columns(excel_file)
category_dfs = split_by_item_prefix(full_df)
projections = {cat: create_monthly_projection(df) for cat, df in category_dfs.items()}
save_to_excel(category_dfs, projections, output_file)

# Optional preview
print("\nProjection for NRE:")
print(projections.get("NRE", pd.DataFrame()).head())
