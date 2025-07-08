
# ───── Order Report Formatter ─────
with tabs[2]:
    st.subheader("📝 Order Report Formatter")
    uploaded_file = st.file_uploader("Upload raw order report file", type=["csv", "xlsx"], key="order")

    if uploaded_file:
        try:
            if uploaded_file.name.endswith(".csv"):
                df = pd.read_csv(uploaded_file)
            else:
                df = pd.read_excel(uploaded_file)

            df.columns = [c.strip() for c in df.columns]

            # Correct columns: Quantity and Line Item Total
            df["Product Count (Units)"] = pd.to_numeric(df["Product Count (Units)"], errors="coerce").fillna(0).astype(int)
            df["Total"] = df["Total"].replace('[\$,]', '', regex=True).replace(',', '', regex=True).astype(float)

            df_out = df[["Buyer Name", "Brand", "Product Count (Units)", "Total"]].copy()
            df_out.columns = ["Customer", "Brand", "Qty (Units)", "Line Item Total"]
            df_out.sort_values(["Customer", "Brand"], inplace=True)

            # Group and subtotal
            grouped_output = []
            for customer, group in df_out.groupby("Customer"):
                grouped_output.append(group)

                total_qty = group["Qty (Units)"].sum()
                total_price = group["Line Item Total"].sum()

                total_row = {
                    "Customer": f"TOTAL - {customer}",
                    "Qty (Units)": total_qty,
                    "Line Item Total": f"${total_price:,.2f}"
                }
                grouped_output.append(pd.DataFrame([total_row]))
                grouped_output.append(pd.DataFrame([{}]))

            df_final = pd.concat(grouped_output, ignore_index=True)

            def to_excel_order(df):
                output = BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    df.to_excel(writer, sheet_name="Formatted", index=False)
                    sheet = writer.sheets["Formatted"]

                    red_bold = Font(color="FF0000", bold=True)
                    bold = Font(bold=True)
                    center = Alignment(horizontal="center", vertical="center")

                    for col_num, col_name in enumerate(df.columns, 1):
                        cell = sheet.cell(row=1, column=col_num)
                        cell.font = red_bold
                        cell.alignment = center
                        col_letter = get_column_letter(col_num)
                        sheet.column_dimensions[col_letter].width = 28 if col_name != "Customer" else 60

                    for row in range(2, sheet.max_row + 1):
                        val = sheet.cell(row=row, column=1).value
                        if isinstance(val, str) and val.startswith("TOTAL -"):
                            for col in range(1, len(df.columns) + 1):
                                sheet.cell(row=row, column=col).font = bold

                    sheet.auto_filter.ref = sheet.dimensions
                    sheet.freeze_panes = "A2"

                output.seek(0)
                return output

            xlsx_order = to_excel_order(df_final)
            st.success("✅ Order Report formatted!")
            st.download_button("📥 Download Order Report Excel", data=xlsx_order, file_name="Formatted_Order_Report.xlsx")

        except Exception as e:
            st.error(f"⚠️ Error: {e}")
