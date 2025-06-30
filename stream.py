import streamlit as st
import pandas as pd
from datetime import datetime, date
from io import BytesIO

# === CONFIG ===
st.set_page_config(page_title="Excel Report Generator", layout="centered")
st.title("📊 Excel Report Generator")

current_date = datetime.today().date()

# === FUNCTIONS ===

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

    return pd.concat(cleaned_rows, ignore_index=True)

def split_by_item_prefix(df):
    df['Category'] = df['Item'].astype(str).str.extract(r'^(999|NRE|ENG)', expand=False)
    df = df.dropna(subset=['Category'])
    return {
        category: df_group.drop(columns='Category')
        for category, df_group in df.groupby('Category')
    }

def create_monthly_projection(df):
    projected_data = []
    for month_num in range(1, 13):
        month_year = datetime(current_date.year, month_num, 1).strftime('%b - %Y')
        month_data = df[df['Dock Date'].apply(lambda x: x.month == month_num and x.year == current_date.year)]
        total = month_data['Extended Price'].sum()
        projected_data.append([total, month_year, month_num])
    return pd.DataFrame(projected_data, columns=['Projected Below', 'Month', 'Month #'])

def generate_excel(original_dfs, projections):
    output = BytesIO()

    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        red_format = workbook.add_format({'bg_color': '#FFC7CE'})
        red_date_format = workbook.add_format({'bg_color': '#FFC7CE', 'num_format': 'yyyy-mm-dd'})

        for category, df in original_dfs.items():
            sheet_name = f"{category}_Data"
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            ws = writer.sheets[sheet_name]

            for i, dock_date in enumerate(df['Dock Date']):
                if dock_date < current_date:
                    for j, val in enumerate(df.iloc[i]):
                        if pd.isna(val) or val in [float('inf'), float('-inf')]:
                            val = ''
                            fmt = red_format
                            ws.write(i + 1, j, val, fmt)

                        elif isinstance(val, (datetime, pd.Timestamp)):
                            ws.write_datetime(i + 1, j, val.to_pydatetime(), red_date_format)

                        elif isinstance(val, date):
                            ws.write_datetime(i + 1, j, datetime.combine(val, datetime.min.time()), red_date_format)

                        else:
                            ws.write(i + 1, j, val, red_format)

            # Write projection table
            projections[category].to_excel(writer, sheet_name=f"{category}_Projection", index=False)

    output.seek(0)
    return output

# === STREAMLIT UI ===

uploaded_file = st.file_uploader("📁 Upload your Excel file", type=["xlsx"])

if uploaded_file:
    try:
        full_df = load_and_select_columns(uploaded_file)
        category_dfs = split_by_item_prefix(full_df)
        projections = {cat: create_monthly_projection(df) for cat, df in category_dfs.items()}
        excel_output = generate_excel(category_dfs, projections)

        # === PREVIEW CLEANED DATA ===
        st.subheader("📋 Preview: Cleaned Combined Data")
        st.dataframe(full_df, use_container_width=True)

        # === PREVIEW PROJECTIONS ===
        st.subheader("📈 Preview: Monthly Projection by Category")
        selected_category = st.selectbox("Select a category to view:", list(projections.keys()))
        st.dataframe(projections[selected_category], use_container_width=True)


        st.success("✅ Analysis complete!")
        st.download_button(
            label="📥 Download Processed Excel",
            data=excel_output,
            file_name="Processed_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"❌ Error processing file: {e}")
