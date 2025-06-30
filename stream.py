import streamlit as st
import pandas as pd
from datetime import datetime, date
from io import BytesIO
import plotly.express as px

# === CONFIG ===
st.set_page_config(page_title="Excel Report Generator", layout="wide")
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

        # Write original data and projections
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

            projections[category].to_excel(writer, sheet_name=f"{category}_Projection", index=False)

        # --- ADD CHARTS FOR ENG CATEGORY INTO EXCEL ---
        if 'ENG' in original_dfs:
            eng_df = original_dfs['ENG']

            # Create a worksheet for charts
            chart_ws = workbook.add_worksheet("ENG_Graphs")

            # Prepare data for pie chart (Company totals)
            company_totals = (
                eng_df.groupby("Name")["Extended Price"]
                .sum()
                .sort_values(ascending=False)
                .reset_index()
            )

            # Write pie chart data starting at A1
            chart_ws.write_row('A1', ['Name', 'Total Extended Price'])
            for row_num, (name, total) in enumerate(company_totals.itertuples(index=False), start=1):
                chart_ws.write(row_num, 0, name)
                chart_ws.write(row_num, 1, total)

            # Create Pie Chart
            pie_chart = workbook.add_chart({'type': 'pie'})
            pie_chart.add_series({
                'name':       'ENG Company Extended Price',
                'categories': ['ENG_Graphs', 1, 0, len(company_totals), 0],
                'values':     ['ENG_Graphs', 1, 1, len(company_totals), 1],
                'data_labels': {'percentage': True, 'category': True},
                'points': [{'fill': {'color': '#5ABA10'}}, {'fill': {'color': '#FE110E'}}],  # example colors
            })

            
            pie_chart.set_title({'name': 'ENG Extended Price by Company'})
            pie_chart.set_style(10)

            # Insert pie chart at E2
            chart_ws.insert_chart('E2', pie_chart, {'x_scale': 1.5, 'y_scale': 1.5})

            # Prepare data for time series line chart (monthly extended price per item)
            eng_df['Order Date'] = pd.to_datetime(eng_df['Order Date'])
            eng_2025 = eng_df[eng_df['Order Date'].dt.year == current_date.year].copy()
            eng_2025['MonthStart'] = eng_2025['Order Date'].values.astype('datetime64[M]')

            item_to_name = (
                eng_2025[['Item', 'Name']]
                .drop_duplicates(subset='Item')
                .set_index('Item')['Name']
                .to_dict()
            )

            grouped = (
                eng_2025.groupby(['MonthStart', 'Item'])['Extended Price']
                .sum()
                .reset_index()
            )

            pivot_df = grouped.pivot(index='MonthStart', columns='Item', values='Extended Price').fillna(0)
            renamed_columns = {
                item: f"{item_to_name.get(item, 'Unknown')} ({item})"
                for item in pivot_df.columns
            }
            pivot_df.rename(columns=renamed_columns, inplace=True)

            # Write time series data starting at A20 (or below pie chart data)
            start_row = len(company_totals) + 3
            chart_ws.write(start_row, 0, 'MonthStart')
            for col_num, col_name in enumerate(pivot_df.columns, start=1):
                chart_ws.write(start_row, col_num, col_name)

            for r_idx, (date_val, row) in enumerate(pivot_df.iterrows(), start=start_row + 1):
                chart_ws.write_datetime(r_idx, 0, pd.to_datetime(date_val).to_pydatetime(), workbook.add_format({'num_format': 'mmm yyyy'}))
                for c_idx, val in enumerate(row, start=1):
                    chart_ws.write(r_idx, c_idx, val)

            # Create Line Chart
            line_chart = workbook.add_chart({'type': 'line'})

            num_rows = len(pivot_df)
            num_cols = len(pivot_df.columns)

            for i in range(num_cols):
                # Excel columns are 0-based; data starts at col=1
                line_chart.add_series({
                    'name':       [ 'ENG_Graphs', start_row, i + 1],
                    'categories': ['ENG_Graphs', start_row + 1, 0, start_row + num_rows, 0],
                    'values':     ['ENG_Graphs', start_row + 1, i + 1, start_row + num_rows, i + 1],
                    'marker':     {'type': 'circle', 'size': 5},
                    'line':       {'width': 2},
                })

            line_chart.set_title({'name': 'ENG Companies: Monthly Extended Price (2025)'})
            line_chart.set_x_axis({'name': 'Month', 'date_axis': True, 'num_format': 'mmm yyyy'})
            line_chart.set_y_axis({'name': 'Extended Price ($)', 'major_gridlines': {'visible': False}})
            line_chart.set_legend({'position': 'bottom'})
            line_chart.set_style(10)

            # Insert line chart below pie chart at E20 (approx)
            chart_ws.insert_chart('E20', line_chart, {'x_scale': 2, 'y_scale': 1.5})

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


        # === PIE CHART FOR ENG COMPANIES ===
        if selected_category == "ENG":
            st.subheader("🍰 ENG Extended Price by Company")

            eng_df = category_dfs["ENG"]
            company_totals = (
                eng_df.groupby("Name")["Extended Price"]
                .sum()
                .sort_values(ascending=False)
                .reset_index()
            )

            fig = px.pie(
                company_totals,
                names="Name",
                values="Extended Price",
                title="Total Extended Price by Company (ENG Items)",
                hole=0.3,  # Donut-style
            )

            fig.update_traces(
                textinfo="percent+label",
                pull=[0.05]*len(company_totals),  # Slight "pop-out" animation
                hovertemplate="<b>%{label}</b><br>Value: %{value:,.2f}<extra></extra>",
                marker=dict(line=dict(color="#000000", width=1))
            )

            fig.update_layout(
                height=600,
                showlegend=True,
                legend_title_text="Company",
                font=dict(size=14),
                paper_bgcolor="rgba(0,0,0,0)",  # transparent background
                plot_bgcolor="rgba(0,0,0,0)"
            )

            st.plotly_chart(fig, use_container_width=True)


            st.subheader("📈 ENG Company Revenue Over Time")

            eng_df["Order Date"] = pd.to_datetime(eng_df["Order Date"])
            eng_2025 = eng_df[eng_df["Order Date"].dt.year == current_date.year].copy()

            # Use actual datetime objects for grouping
            eng_2025["MonthStart"] = eng_2025["Order Date"].values.astype("datetime64[M]")

            # Build Item → Company map
            item_to_name = (
                eng_2025[["Item", "Name"]]
                .drop_duplicates(subset="Item")
                .set_index("Item")["Name"]
                .to_dict()
            )

            # Group by month start and item
            grouped = (
                eng_2025.groupby(["MonthStart", "Item"])["Extended Price"]
                .sum()
                .reset_index()
            )

            # Pivot so rows = date, columns = item
            pivot_df = grouped.pivot(index="MonthStart", columns="Item", values="Extended Price").fillna(0)

            # Rename columns: Item → Company (Item)
            renamed_columns = {
                item: f"{item_to_name.get(item, 'Unknown')} ({item})"
                for item in pivot_df.columns
            }
            pivot_df.rename(columns=renamed_columns, inplace=True)

            # Plot
            fig2 = px.line(
                pivot_df,
                x=pivot_df.index,
                y=pivot_df.columns,
                title="ENG Companies: Monthly Extended Price (2025)",
                markers=True
            )

            fig2.update_layout(
                xaxis_title="Month",
                yaxis_title="Extended Price ($)",
                legend_title_text="Company (Item Code)",
                height=600,
                font=dict(size=13),
                hovermode="x",
                plot_bgcolor="rgba(0,0,0,0)",
                paper_bgcolor="rgba(0,0,0,0)",
                xaxis=dict(
                    tickformat="%b",  # Format as Jan, Feb, etc.
                    dtick="M1",       # Monthly ticks
                )
            )

            fig2.update_traces(
                line=dict(width=2),
                marker=dict(size=6),
                hovertemplate="<b>%{y:,.2f}</b><br>Month: %{x|%b %Y}<extra></extra>"
            )

            st.plotly_chart(fig2, use_container_width=True)








        st.success("✅ Analysis complete!")
        st.download_button(
            label="📥 Download Processed Excel",
            data=excel_output,
            file_name="Processed_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"❌ Error processing file: {e}")
