import re
import pandas as pd

def is_date_format(s: str) -> bool:
    """
    Check if a given string matches the format MM/DD/YYYY.
    This is a simple pattern match for dates such as '10/15/2021'.
    
    :param s: The string to check
    :return: True if it matches MM/DD/YYYY, False otherwise
    """
    return bool(re.match(r'^\d{1,2}/\d{1,2}/\d{4}$', s))

def convert_number(s: str) -> float:
    """
    Convert a string representing a currency amount to a float.
    This function removes commas, handles parentheses for negative amounts,
    and returns a float value.
    
    Examples:
      - '5,200.00'  -> 5200.00
      - '(5,837.50)' -> -5837.50
      - '' (empty) -> 0.0  (You can modify this behavior if needed)
    
    :param s: The string containing a numeric amount
    :return: A float representation of the numeric value
    """
    s = s.strip()
    # Remove commas
    s = s.replace(',', '')
    # Convert parentheses notation to negative
    if s.startswith('(') and s.endswith(')'):
        s = '-' + s[1:-1]
    # Handle empty string as zero
    if s == '':
        return 0.0
    return float(s)

# ------------------------------------------------------------------------
# 1) Set input/output file paths
# ------------------------------------------------------------------------
input_path = r'C:\Users\slee\OneDrive - SBP\Desktop\sungkeun\report tranformation\INPUT BNA FA Disposal Reeb.xls'
output_path = r'C:\Users\slee\OneDrive - SBP\Desktop\sungkeun\report tranformation\Output BNA_FA_Disposal_Reeb.xlsx'

# ------------------------------------------------------------------------
# 2) Read the Excel file into a pandas DataFrame
#    header=None means we treat every row as data, ignoring any default header row
# ------------------------------------------------------------------------
df_raw = pd.read_excel(input_path, header=None)

# We will store parsed records in a list of dictionaries
records = []

# Variables to hold the current GL Account info and asset description
current_gl_full = None   # e.g., "80-000-10002030"
current_gl_short = None  # e.g., "10002030" extracted from the above
asset_desc = None        # For storing the asset description line

# ------------------------------------------------------------------------
# 3) Iterate over each row in the DataFrame to parse the data
# ------------------------------------------------------------------------
for idx, row in df_raw.iterrows():
    # (A) Combine all columns in this row into one string, skipping empty cells (NaN)
    row_str_list = []
    for cell in row:
        if pd.notnull(cell):
            row_str_list.append(str(cell))
    line = " ".join(row_str_list).strip()

    # (B) If the line starts with "Asset GL Acct #:", we extract the GL account number
    if line.startswith("Asset GL Acct #:"):
        # Example: "Asset GL Acct #: 80-000-10002030"
        m = re.search(r'Asset GL Acct #:\s*([\d\-]+)', line)
        if m:
            current_gl_full = m.group(1)  # "80-000-10002030"
            # Extract the portion after the first 7 characters, e.g. "10002030"
            if len(current_gl_full) >= 7:
                current_gl_short = current_gl_full[7:]
            else:
                current_gl_short = current_gl_full
        continue

    # (C) Skip empty lines or lines starting with "Subtotal:", "Page:", or "Printed:"
    if (not line) or line.startswith("Subtotal:") or line.startswith("Page:") or line.startswith("Printed:"):
        continue

    # (D) Check if this line is an asset description (e.g., "2009 VOLVO TRACTOR N281672")
    #     We define an asset description as a line that contains text/numbers but not a date
    if re.match(r'^[A-Za-z0-9\s]+$', line) and not re.search(r'\d{1,2}/\d{1,2}/\d{4}', line):
        asset_desc = line
        continue

    # (E) If the line might be an asset record, we split by whitespace and check
    #     if fields[1] and fields[2] look like dates (MM/DD/YYYY)
    fields = re.split(r'\s+', line)
    if len(fields) >= 7:
        # Check if the 2nd and 3rd items are dates, e.g. 10/15/2021, 03/05/2024
        if is_date_format(fields[1]) and is_date_format(fields[2]):
            # Parse the fields
            asset_id      = fields[0]
            placed_date   = fields[1]
            disposal_date = fields[2]
            cost_plus     = convert_number(fields[3])
            ltd_depr      = convert_number(fields[4])
            net_proceeds  = convert_number(fields[5])
            realized_gain = convert_number(fields[6])

            record = {
                'Asset ID': asset_id,
                'Asset Description': asset_desc if asset_desc else "",
                'Placed In Service': placed_date,
                'Disposal Date': disposal_date,
                'Cost Plus Exp. of Sale': cost_plus,     # float
                'LTD Depr & S179/A & AFYD': ltd_depr,    # float
                'Net Proceeds': net_proceeds,            # float
                'Realized Gain (Loss)': realized_gain,   # float
                'Asset GL Acct #': current_gl_short
            }
            records.append(record)

            # Clear the asset description so it doesn't carry over to the next item
            asset_desc = None

# ------------------------------------------------------------------------
# 4) Convert the list of records into a DataFrame and write to Excel
# ------------------------------------------------------------------------
df_result = pd.DataFrame(records)
df_result.to_excel(output_path, index=False)

print("Output file has been saved to:", output_path)
