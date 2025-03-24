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
    
    header_texts = {"Накладная": "", "Поставщик:": "", "Покупатель:": ""}
    for row in df.iloc[:7].values:  
        for i, cell in enumerate(row):
            if isinstance(cell, str):
                for key in header_texts:
                    if key in cell:
                        if key == "Накладная":
                            header_texts[key] = cell.strip()
                        else:
                            value = ""
                            for j in range(i + 1, len(row)):  
                                if isinstance(row[j], str) and row[j].strip():
                                    value = row[j].strip()
                                    break  
                            header_texts[key] = f"{cell.strip()} {value}".strip()

    # **Find footer texts (from last 7 rows)**
    footer_texts = {"Итого:": "", "НДС:": ""}
    for row in df.iloc[-8:].values:  # Scan last 7 rows
        for i, cell in enumerate(row):
            if isinstance(cell, str):
                for key in footer_texts:
                    if key in cell:
                        value = ""
                        for j in range(i + 1, len(row)):  
                            if isinstance(row[j], str) and row[j].strip():
                                value = row[j].strip()
                                break  
                        footer_texts[key] = f"{cell.strip()} {value}".strip()    
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
    print (results)
    for _, row in df_data.iterrows():
        product_name = " ".join([word.capitalize() for word in trim_product_name(row[product_col]).split()])

        if len (product_name.strip()) == 0 : continue
        quantity = row[quantity_col].split()[0] if pd.notna(row[quantity_col]) else "0"
        total_value = row[total_col] if pd.notna(row[total_col]) else "0"
        
        if product_name != current_product and current_product:
            results.append([current_order, current_product, current_places, total_quantity, total_sum])
            print (results)
            current_order = current_order + 1
            total_quantity, total_sum = 0, 0
        
        current_product = product_name
        total_quantity += int(quantity)
        total_sum += float(total_value)
    
    if current_product:
        results.append([current_order, current_product, current_places, total_quantity, total_sum])
    
    # Create output DataFrame (WITHOUT inserting extra rows yet)
    output_df = pd.DataFrame(results, columns=["№", "Товар", "Мест", "Количество", "Сумма"])
    print (output_df)
    output_df.to_excel(output_path, index=False, engine="openpyxl")

    # Now modify the file with openpyxl
    wb = load_workbook(output_path)
    ws = wb.active

    # **Insert header texts (without shifting the analysis)**
    ws.insert_rows(1, amount=3)
    ws["B1"], ws["B2"], ws["B3"] = header_texts["Накладная"], header_texts["Поставщик:"], header_texts["Покупатель:"]

    # **Insert footer texts after last row**
    footer_row = ws.max_row + 1
    ws.append([""])  # Empty row before totals
    print([footer_texts["Итого:"]])
    ws.append([footer_texts["Итого:"]])  # Corrected from header_texts
    ws.append([footer_texts["НДС:"]])  # Corrected from header_texts

    # **Apply bold styling to headers (corrected row number)**
    for cell in ws[4]:  # Headers now on row 4
        cell.font = Font(bold=True)

    # **Auto-adjust column width**
    for col in ws.columns:
        max_length = 0
        col_letter = col[0].column_letter
        for cell in col:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass
        ws.column_dimensions[col_letter].width = max_length + 2

    # Save final result
    wb.save(output_path)

    print(f"Analysis complete. Output saved to {output_path}")


# Run the function
analyze_excel()
