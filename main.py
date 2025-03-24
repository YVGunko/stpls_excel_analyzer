import pandas as pd
import os
from tkinter import Tk, filedialog
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment

def trim_product_name(product_name):
    if not isinstance(product_name, str) or not product_name.strip():
        return ""
    words = product_name.split()
    if product_name.lower().startswith("подошва") and len(words) > 3:
        return " ".join(words[:3])
    elif product_name.lower().startswith("стелька") and len(words) > 2:
        return " ".join(words[:2])
    return product_name

def analyze_excel():
    # Open file dialog to select input Excel file
    Tk().withdraw()  # Hide the root window
    input_path = filedialog.askopenfilename(title="Select Excel File", filetypes=[("Excel Files", "*.xlsx")])
    if not input_path:
        print("No file selected. Exiting.")
        return
    
    # Generate output file name
    output_path = os.path.join(os.path.dirname(input_path), "s_" + os.path.basename(input_path))
    
    # Load the Excel file
    xls = pd.ExcelFile(input_path)
    df = pd.read_excel(xls, sheet_name=xls.sheet_names[0], dtype=str)
    
    # Copy the first 7 rows as they are
    df_pre_analysis = df.iloc[:7].copy()
    
    # Identify the correct header row
    header_row_index = 7
    df_headers = pd.read_excel(xls, sheet_name=xls.sheet_names[0], skiprows=header_row_index, nrows=1, dtype=str)
    clean_headers = df_headers.iloc[0].fillna("").astype(str).tolist()
    df = pd.read_excel(xls, sheet_name=xls.sheet_names[0], skiprows=header_row_index + 1, dtype=str)
    df.columns = clean_headers
    
    # Find the first row where analysis should start
    first_data_row = df[df.iloc[:, 1] == "1"].index.min()
    
    # Extract required columns
    order_col = df.columns[df.columns.str.contains("№", case=False, na=False)][0]
    product_col = df.columns[df.columns.str.contains("Товар", case=False, na=False)][0]
    quantity_col = df.columns[df.columns.str.contains("Количество", case=False, na=False)][0]
    places_col = df.columns[df.columns.str.contains("Мест", case=False, na=False)][0]
    total_col = df.columns[df.columns.str.contains("Сумма", case=False, na=False)][0]
    
    # Process relevant rows
    df_data = df.iloc[first_data_row:].reset_index(drop=True)
    
    # Group similar product names
    results = []
    current_order = 1
    current_product = ""
    current_places = ""
    total_quantity = 0
    total_sum = 0
    
    for _, row in df_data.iterrows():
        product_name = " ".join([word.capitalize() for word in trim_product_name(row[product_col]).split()])

        if len (product_name.strip()) == 0 : continue
        quantity = row[quantity_col].split()[0] if pd.notna(row[quantity_col]) else "0"
        total_value = row[total_col] if pd.notna(row[total_col]) else "0"
        
        if product_name != current_product and current_product:
            results.append([current_order, current_product, current_places, total_quantity, total_sum])
            current_order = current_order + 1
            total_quantity, total_sum = 0, 0
        
        current_product = product_name
        total_quantity += int(quantity)
        total_sum += float(total_value)
    
    if current_product:
        results.append([current_order, current_product, current_places, total_quantity, total_sum])
    
    # Create output DataFrame
    output_df = pd.DataFrame(results, columns=["№", "Товар", "Мест", "Количество", "Сумма"])

    # Save the DataFrame to an Excel file
    output_df.to_excel(output_path, index=False, engine="openpyxl")

    # **Save with basic formatting**
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        output_df.to_excel(writer, index=False, sheet_name="Analysis")
        workbook = writer.book
        worksheet = writer.sheets["Analysis"]

        # **Auto-adjust column width**
        for col_idx, col in enumerate(output_df.columns, start=1):
            max_length = max(output_df[col].astype(str).map(len).max(), len(col)) + 2
            worksheet.column_dimensions[worksheet.cell(row=1, column=col_idx).column_letter].width = max_length


        # Set header style
        for cell in worksheet[1]:
            cell.font = cell.font.copy(bold=True)

        # Format numbers
        for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row, min_col=3, max_col=5):
            for cell in row:
                cell.number_format = "#,##0.00"

    print(f"Analysis complete. Output saved to {output_path}")


# Run the function
analyze_excel()
