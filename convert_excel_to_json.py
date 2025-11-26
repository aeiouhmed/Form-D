import pandas as pd
import json
import re
from pathlib import Path

# Paths (Adjust if necessary)
EXCEL_PATH = Path("server/templates/HSCODE.xlsx")
JSON_PATH = Path("server/templates/HSCODE.json")

def normalize_header(value):
    return re.sub(r"[\s_\-]+", "", str(value or "").strip().lower())

def get_mapping(df):
    # Find columns
    norm_cols = {normalize_header(c): c for c in df.columns}
    
    # Logic to find 'HS Code' and 'Unit' columns
    code_col, unit_col = None, None
    for variant in ["Hs Code", "HS Code", "Hscode", "HSCode", "HsCode"]:
        if normalize_header(variant) in norm_cols:
            code_col = norm_cols[normalize_header(variant)]
            break
    
    for variant in ["Unit", "Units", "UOM", "UNTnit"]:
        if normalize_header(variant) in norm_cols:
            unit_col = norm_cols[normalize_header(variant)]
            break
            
    if not code_col or not unit_col:
        return {}

    # Extract data
    mapping = {}
    for code, unit in zip(df[code_col], df[unit_col]):
        # Sanitize Code: Keep only digits
        clean_code = re.sub(r"\D", "", str(code))
        if not clean_code: continue
        
        # Sanitize Unit: Uppercase, strip
        clean_unit = str(unit or "").strip().upper()
        if not clean_unit or clean_unit in ["NAN", "NONE", ""]: continue
        
        mapping[clean_code] = clean_unit
    return mapping

def main():
    print(f"Reading {EXCEL_PATH}...")
    sheets = pd.read_excel(EXCEL_PATH, sheet_name=None, dtype=str)
    
    final_mapping = {}
    
    # Process 'Sheet2' first (priority), then others
    sheet_names = [n for n in sheets.keys() if n == "Sheet2"] + \
                  [n for n in sheets.keys() if n != "Sheet2"]
                  
    for name in sheet_names:
        print(f"Processing sheet: {name}")
        data = get_mapping(sheets[name])
        # Update mapping (existing keys are NOT overwritten to preserve priority)
        for k, v in data.items():
            if k not in final_mapping:
                final_mapping[k] = v
                
    print(f"Extracted {len(final_mapping)} HS Codes.")
    
    with open(JSON_PATH, "w") as f:
        json.dump(final_mapping, f, indent=2)
    
    print(f"Success! Saved to {JSON_PATH}")

if __name__ == "__main__":
    main()