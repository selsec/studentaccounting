import pandas as pd
import tkinter as tk
from tkinter import filedialog
import sys
import xlsxwriter

def get_excel_file_path():
    root = tk.Tk()
    root.withdraw()
    excel_file_path = filedialog.askopenfilename(title="Select an Excel file", filetypes=[("Excel Files", "*.xlsx")])
    return excel_file_path

def create_excel_sheets(output_excel_file_path, data, original_df):
    writer = pd.ExcelWriter(output_excel_file_path, engine='xlsxwriter')
    
    # Create the "master" sheet
    master_data = []
    for sheet_name, content in data.items():
        master_row = [sheet_name, "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""]
        master_data.append(master_row)
    
    master_df = pd.DataFrame(master_data, columns=["student name", "marching/concert", "grade", "instrument", 
                                                   "FairShare ($350/$450)", "Percussion Fee ($100)", "Bibbers($60)",
                                                   "Shoes ($30)", "Suit ($50)", "Dress ($70)", "All County ($10)", 
                                                   "S&E", "State", "Indoor Winds", "Leadership Cord", "Gloves", 
                                                   "Chaperone Shirt", "Extra Show Shirt", "Fundraiser 1", "Senior Banners"])
    
    master_df.to_excel(writer, sheet_name="master", index=False)
    
    # Create other sheets
    for sheet_name, content in data.items():
        df = pd.DataFrame(content)
        
        # Add columns for detailed information
        df.insert(0, "Date", "")
        df.insert(1, "Amount", "")
        df.insert(2, "MySchoolBucks", "")
        df.insert(3, "Check #", "")
        df.insert(4, "Receipt #", "")
        df.insert(5, "Comment", "")
        
        # Add "Grade" and "Instrument" columns
        df.insert(6, "Grade", original_df.loc[df.index, 'GRADE'])
        df.insert(7, "Instrument", original_df.loc[df.index, 'INSTRUMENT'])
        
        # Add rows for expenses and fees
        expenses_rows = [
            "Fairshare", "Percussion Fee", "Bibbers", "Shoes", "Suit", "Dress",
            "All County", "SE", "State", "Indoor Winds", "Indoor Guard",
            "Leadership Cord", "Gloves", "Chaperone Shirt", "Extra Show Shirt",
            "Fundraiser 1", "Senior Banners"
        ]
        
        for row_idx, row_name in enumerate(expenses_rows, start=20):
            df.loc[row_idx, "Date"] = row_name
        
        df.to_excel(writer, sheet_name=sheet_name, index=False)
    
    writer.close()
    print("Sheets created and saved to", output_excel_file_path)

def main():
    try:
        if len(sys.argv) > 1:
            excel_file_path = sys.argv[1]
        else:
            excel_file_path = get_excel_file_path()

        if not excel_file_path:
            print("No Excel file selected.")
            return
        
        sheet_name = 'Sheet1'  # Change this to your sheet name
        column_a_name = 'FIRST NAME'  # Change this to your first column name
        column_b_name = 'LAST NAME'  # Change this to your second column name
        grade_column_name = 'GRADE'  # Change this to the name of the Grade column
        instrument_column_name = 'INSTRUMENT'  # Change this to the name of the Instrument column
        
        df = pd.read_excel(excel_file_path, sheet_name=sheet_name)
        
        combined_data = []

        for index, row in df.iterrows():
            combined_value = f"{row[column_b_name]}, {row[column_a_name]}"
            combined_data.append(combined_value)
        
        if not combined_data:
            print("No data to combine.")
            return
        
        df['Combined'] = combined_data

        output_excel_file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
        df.to_excel(output_excel_file_path, index=False)

        print("Columns combined and saved to", output_excel_file_path)
        
        data = {}

        for combined_value in combined_data:
            if combined_value not in data:
                data[combined_value] = df[df['Combined'] == combined_value].drop(columns=['Combined', 'FIRST NAME', 'LAST NAME', 'GRADE', 'INSTRUMENT'])
        
        create_excel_sheets(output_excel_file_path, data, df[[grade_column_name, instrument_column_name]])
        
    except Exception as e:
        print("An error occurred:", e)

if __name__ == "__main__":
    main()
