import streamlit as st
import pandas as pd
import os
import re
from datetime import datetime
import io  # For in-memory Excel output

st.title("Customer Order Classification Report Generator")
st.write("Upload your 'Customer Order Report_*.xlsx' file below. The app will process it and generate a summary report (Rate Freeze: Yes/Manual only, subtotals excluded).")

uploaded_file = st.file_uploader("Choose your Excel file", type="xlsx")

if uploaded_file is not None:
    try:
        # Read data (same as your script)
        df = pd.read_excel(uploaded_file, skiprows=4, engine='openpyxl')
        df.columns = df.columns.str.strip().str.replace(r'\s+', ' ', regex=True)

        # Column detection (same)
        def find_column(contains):
            for col in df.columns:
                if contains.lower() in col.lower():
                    return col
            return None

        group_col = find_column("Classification") or "Classificationg Roup"
        rate_col = find_column("Rate Freeze")
        date_col = find_column("Date") or "Date"
        weight_cols = [find_column(k) for k in ['Gross Wt', 'Net Wt', 'Fine Wt', 'Metal Amount'] if find_column(k)]

        if not all([group_col, rate_col, date_col]) or not weight_cols:
            st.error("Critical columns not found in the uploaded file. Check your Excel format.")
        else:
            # Clean numeric columns
            for col in weight_cols:
                df[col] = pd.to_numeric(df[col].astype(str).str.replace(',', ''), errors='coerce')

            # Filter (same)
            df_filtered = df[
                df[rate_col].notna() &
                (~df[rate_col].astype(str).str.strip().str.lower().isin(['', 'no', 'NO']))
            ].copy()

            # Remove summary rows
            subtotal_patterns = ['sub total', 'total', 'printed by', 'subtotal']
            df_filtered = df_filtered[
                ~df_filtered[date_col].astype(str).str.lower().str.contains('|'.join(subtotal_patterns), na=False)
            ]

            # Group mapping (same)
            group_mapping = {
                'gold jewellery 22k': 'Gold Jewellery',
                'gold jewellery 18k': 'Gold Jewellery',
                'diamond jewellery 18k': 'Diamond Jewellery 18karat',
                'silver': 'Silver',
                'standard bar': 'Standard Bar',
                'coin gold': 'Standard Bar',
                'gold bar': 'Standard Bar',
            }

            def map_group(val):
                if pd.isna(val):
                    return 'Unknown'
                text = str(val).strip().lower()
                for k, v in group_mapping.items():
                    if k in text:
                        return v
                return text.title()

            df_filtered['Group'] = df_filtered[group_col].apply(map_group)

            # Summary (same)
            summary = df_filtered.groupby('Group', as_index=False)[weight_cols].sum().round(3)

            # Nice column names
            nice_names = ['Group', 'Gross Wt', 'Net Wt', 'Fine Wt', 'Metal Amount']
            if len(weight_cols) <= 4:
                summary.columns = nice_names[:len(weight_cols) + 1]

            # Desired order
            order_list = ['Gold Jewellery', 'Silver', 'Diamond Jewellery 18karat', 'Standard Bar']
            summary = summary[summary['Group'].isin(order_list)].sort_values(
                'Group', key=lambda x: x.map({k: i for i, k in enumerate(order_list)})
            )

            # Display results
            st.success("Processing complete!")
            st.write("### Final Classification Group Wise Report")
            st.dataframe(summary)

            # Download button
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                summary.to_excel(writer, index=False)
            output.seek(0)
            st.download_button(
                label="Download Report as XLSX",
                data=output,
                file_name=f"Final_Classification_Report_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    except Exception as e:
        st.error(f"Error processing file: {str(e)}")
