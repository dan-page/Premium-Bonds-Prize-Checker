import requests
import os
from datetime import datetime
import pandas as pd
from dateutil.relativedelta import relativedelta
import zipfile

# Need setting to the correct file and path
filename = 'Winnings.xlsx'
path = '/Users/.../Winnings.xlsx'


def generate_months_to_check(num_months_back):
    current_date = datetime.now()
    months_to_check = [(current_date - relativedelta(months=i)).replace(day=1) for i in range(num_months_back, -1, -1)]
    return months_to_check

def find_missing_months(months_to_check, winnings_df):
    data_missing = False
    missing_months = []

    # Convert winnings DataFrame 'Year-Month' column to a list of Periods for easier checking
    recorded_periods = winnings_df['Year-Month'].unique().tolist()

    # Get the period for the current month
    current_month_period = pd.Timestamp(datetime.now()).to_period('M')

    # Track the status of the previous month to identify isolated missing months
    prev_month_missing = False

    for i, month in enumerate(months_to_check):
        month_period = pd.Timestamp(month).to_period('M')

        # Check if the current month is missing
        if month_period not in recorded_periods:
            # If it's the most recent month, mark it as missing
            if month_period == current_month_period:
                data_missing = True
                missing_months.append(month)
                prev_month_missing = True
            else:
                # Check if the next month is also missing or if the previous month was missing
                if (i < len(months_to_check) - 1 and pd.Timestamp(months_to_check[i + 1]).to_period('M') not in recorded_periods) or prev_month_missing:
                    data_missing = True
                    missing_months.append(month)
                    prev_month_missing = True
                else:
                    # If the next month is not missing and the previous month was not missing, it's an isolated missing month
                    prev_month_missing = False
        else:
            prev_month_missing = False

    return data_missing, missing_months

def generate_bond_numbers(df):
    # Initialize an empty list to hold all generated bond numbers
    held_bond_numbers = []

    # Iterate through each row in the DataFrame to get start and end bond numbers
    for index, row in df.iterrows():
        start_bond, end_bond = row['Starting Bond Number'], row['Ending Bond Number']

        # Find the shared prefix for the start and end bond numbers
        prefix = find_shared_prefix(start_bond, end_bond)

        # Extract the numeric part by removing the prefix
        start_seq = int(start_bond[len(prefix):])
        end_seq = int(end_bond[len(prefix):])

        # Generate all bond numbers in the range and add them to the list
        held_bond_numbers.extend([f"{prefix}{str(seq).zfill(len(start_bond) - len(prefix))}" for seq in range(start_seq, end_seq + 1)])

    return held_bond_numbers

def find_shared_prefix(a, b):
    min_length = min(len(a), len(b))
    for i in range(min_length):
        if a[i] != b[i]:
            return a[:i]
    return a[:min_length]  # In case one string is a complete prefix of the other

def extract_valid_prefixes(df):
    prefixes = set()
    for index, row in df.iterrows():
        start_bond, end_bond = row['Starting Bond Number'], row['Ending Bond Number']
        prefix = find_shared_prefix(start_bond, end_bond)
        prefixes.add(prefix)
    return list(prefixes)

def parse_content(content, valid_prefixes):
    bond_prizes = {}
    lines = content.splitlines()
    current_prize = None

    for line in lines:
        if '£' in line:
            # This assumes that the line with £ contains the prize amount
            current_prize = line.split('£')[1].split()[0].replace(',', '').strip()
        else:
            # Extract bond numbers and associate them with the current prize
            bond_numbers = line.split()
            for number in bond_numbers:
                # Check if the number starts with any of the valid prefixes
                if any(number.startswith(prefix) for prefix in valid_prefixes):
                    bond_prizes[number] = current_prize
    return bond_prizes

def generate_next_id(df):
    if df.empty or 'Unique Identifier' not in df.columns or not df['Unique Identifier'].str.startswith('P').any():
        return 'P1'  # Start from 'P1' if DataFrame is empty or no ID starts with 'P'
    else:
        max_id = df['Unique Identifier'].str.extract(r'P(\d+)').astype(int).max().iloc[0]
        return f'P{max_id + 1}'

def format_and_save_winnings_df(winnings_df, filename):
    # Drop the 'Year-Month' column if it exists
    if 'Year-Month' in winnings_df.columns:
        winnings_df.drop(columns=['Year-Month'], inplace=True)

    # Ensure 'Draw Date' is in datetime format
    winnings_df['Draw Date'] = pd.to_datetime(winnings_df['Draw Date'], dayfirst=True)

    # Convert 'Draw Date' to 'dd/mm/YYYY' format
    winnings_df['Draw Date'] = winnings_df['Draw Date'].dt.strftime('%d/%m/%Y')

    # Save the DataFrame to an Excel file
    with pd.ExcelWriter(filename, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        winnings_df.to_excel(writer, sheet_name="NS&I Winnings", index=False)

    print("Winnings data formatted and saved successfully.")


# Read the 'Easy Access' sheet to find the highest interest rate
#easy_access_df = pd.read_excel(path, sheet_name='Easy Access')
#highest_interest_rate = easy_access_df['Interest Rate'].max()

# Read the 'NS&I Winnings' sheet from the Excel file, without parsing dates initially
winnings_df = pd.read_excel(path, sheet_name='NS&I Winnings', engine='openpyxl')
winnings_df['Draw Date'] = pd.to_datetime(winnings_df['Draw Date'], dayfirst=True)
winnings_df['Year-Month'] = winnings_df['Draw Date'].dt.to_period('M')

holdings_df = pd.read_excel(path, sheet_name='NS&I Holdings', engine='openpyxl')

months_to_check = generate_months_to_check(6)
data_missing, missing_months = find_missing_months(months_to_check, winnings_df)

if data_missing:
    for missing_month_datetime in missing_months:
        # Format the missing month and year from the datetime object
        missing_month_str = missing_month_datetime.strftime("%m")
        missing_year_str = missing_month_datetime.strftime("%Y")
        
        # Convert the formatted month and year to integers if needed
        missing_month_int = int(missing_month_str)
        missing_year_int = int(missing_year_str)

        print(f"NS&I data missing for {missing_month_datetime.strftime('%B %Y')}.")

        # Calculate Held Bond Numbers
        held_bond_numbers = generate_bond_numbers(holdings_df)
    
        # Construct the URL for the missing month
        url = f'https://www.nsandi.com/files/asset/zip/premium-bonds-winning-bond-numbers-{missing_month_str}-{missing_year_str}.zip'

        response = requests.get(url)
        if response.status_code == 200:
            with open('premium_bonds.zip', 'wb') as file:
                file.write(response.content)

            extracted_files = []
            with zipfile.ZipFile('premium_bonds.zip', 'r') as zip_ref:
                zip_ref.extractall()
                extracted_files.extend(zip_ref.namelist())

                for file_name in zip_ref.namelist():
                    with zip_ref.open(file_name) as file:
                        content = file.read().decode('ISO-8859-1')

                        # Assuming 'df' is your DataFrame with bond ranges
                        valid_prefixes = extract_valid_prefixes(holdings_df)

                        # Assuming 'content' is the text content you're parsing
                        bond_prizes = parse_content(content, valid_prefixes)

                        # Create a datetime object for the first day of the missing month
                        draw_date = datetime(missing_year_int, missing_month_int, 1)

                        if isinstance(winnings_df, pd.DataFrame):
                            # Initialize a list to hold new rows DataFrames
                            new_rows = []
                            for bond in held_bond_numbers:
                                if bond in bond_prizes:
                                    unique_id = generate_next_id(winnings_df)
                                    winnings_amount = int(bond_prizes[bond])
                                    # Create a new DataFrame for the current winning detail
                                    new_row_df = pd.DataFrame({
                                        'Bond Number': [bond], 
                                        'Draw Date': [draw_date],
                                        'Winnings': [winnings_amount],
                                        'Unique Identifier': [unique_id],
                                        #'Max Interest': [highest_interest_rate],
                                    })

                                    # Add the new DataFrame to the list
                                    new_rows.append(new_row_df)

                                    # Update winnings_df with the new row to ensure the next ID is unique
                                    winnings_df = pd.concat([winnings_df, new_row_df], ignore_index=True)         
                        else:
                            print("Error: winnings_df is not a DataFrame.") 

            # Cleanup: Delete extracted files
            for file_name in extracted_files:
                os.remove(file_name)
            # Cleanup: Delete the ZIP file
            os.remove('premium_bonds.zip')
        else:
            print("Failed to download the file.")

format_and_save_winnings_df(winnings_df, filename)

