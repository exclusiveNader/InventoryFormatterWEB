
import streamlit as st
import pandas as pd
import re
from io import BytesIO
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Inventory Formatter", layout="centered")

from PIL import Image

logo = Image.open("logo.png")
st.image(logo, width=200)

st.title("📦 Inventory Formatter (Web)")
st.markdown("""
### How to Use:
1. Download your INVENTORY OVERVIEW REPORT on Leaflink
2. 📁 Upload your raw inventory file (.csv or .xlsx)
3. 📊 The app will automatically format it:
   - Grouped by Brand → Product Line
   - Adds subtotals and clean spacing
   - Missing Product Lines become **"Uncategorized"**
4. 📥 Click the green button to download your formatted Excel file

**Works on Mac, PC, and mobile browsers.**
""")

uploaded_file = st.file_uploader("Upload raw inventory file (.csv or .xlsx)", type=["csv", "xlsx"])

# Helper function to clean price
def clean_wholesale_price(value):
    if pd.isna(value): return 0.0
    if isinstance(value, (int, float)): return float(value)
    cleaned = re.sub(r"[^0-9.]", "", str(value))
    try: return float(cleaned)
    except: return 0.0

def process_inventory(df_raw):
    column_map = {
        "Wholesale Price ($)": "Wholesale Price",
        "Wholesale Price": "Wholesale Price",
        "Available Inventory (Units)": "Available Inventory (Units)"
    }
    required = ["Name", "Wholesale Price", "Brand", "Product Line", "Classification", "Listing State", "Available Inventory (Units)"]
    for alt, canonical in column_map.items():
        if alt in df_raw.columns and canonical not in df_raw.columns:
            df_raw.rename(columns={alt: canonical}, inplace=True)
    for col in required:
        if col not in df_raw.columns:
            raise ValueError(f"Missing column: {col}")

    df = df_raw[required].copy()
    df["Wholesale Price"] = df["Wholesale Price"].apply(clean_wholesale_price).round(2)
    df["Product Line"] = df["Product Line"].fillna("Uncategorized")
    df.sort_values(["Brand", "Product Line", "Name"], inplace=True, ignore_index=True)

    grouped = []
    for (brand, line), group in df.groupby(["Brand", "Product Line"]):
        grouped.append(group)
        total_units = group["Available Inventory (Units)"].sum()
        total_row = {
            "Name": f"TOTAL - {line}",
            "Wholesale Price": "",
            "Brand": brand,
            "Product Line": line,
            "Classification": "",
            "Listing State": "",
            "Available Inventory (Units)": total_units
        }
        grouped.append(pd.DataFrame([total_row]))
        grouped.append(pd.DataFrame([{}]))  # blank line

    return pd.concat(grouped, ignore_index=True)

def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Formatted Inventory", index=False)
        sheet = writer.sheets["Formatted Inventory"]

        # Formatting
        red_bold = Font(color="FF0000", bold=True)
        bold = Font(bold=True)
        center = Alignment(horizontal="center", vertical="center")

        for col_num, col_name in enumerate(df.columns, 1):
            cell = sheet.cell(row=1, column=col_num)
            cell.font = red_bold
            cell.alignment = center
            col_letter = get_column_letter(col_num)
            sheet.column_dimensions[col_letter].width = 30 if col_name != "Name" else 60

        for row in range(2, sheet.max_row + 1):
            name_val = sheet.cell(row=row, column=1).value
            if isinstance(name_val, str) and name_val.startswith("TOTAL"):
                for col in range(1, len(df.columns) + 1):
                    sheet.cell(row=row, column=col).font = bold

        sheet.auto_filter.ref = sheet.dimensions
        sheet.freeze_panes = "A2"

    output.seek(0)
    return output

if uploaded_file:
    try:
        if uploaded_file.name.endswith(".csv"):
            df_raw = pd.read_csv(uploaded_file)
        else:
            df_raw = pd.read_excel(uploaded_file)

        df_formatted = process_inventory(df_raw)
        xlsx_data = to_excel(df_formatted)

        st.success("✅ File formatted successfully!")
        st.download_button("📥 Download Formatted File", xlsx_data, file_name="Formatted_Inventory.xlsx")

    except Exception as e:
        st.error(f"⚠️ Error: {e}")
