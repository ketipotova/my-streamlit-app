import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO


def read_excel_file(file):
    try:
        # Try reading with openpyxl engine first
        data = pd.read_excel(file, engine='openpyxl')
        return data
    except Exception as e:
        try:
            # If openpyxl fails, try with xlrd engine
            file.seek(0)  # Reset file pointer
            data = pd.read_excel(file, engine='xlrd')
            return data
        except Exception as e:
            st.error(f"Failed to read file. Error: {e}")
            return None


def process_data(main, pf_leaves, pf_id, shifts):
    # Merge relevant columns from 'pf_id' into 'pf_leaves'
    pf_leaves = pd.merge(pf_leaves, pf_id[['Email', 'ID number']], on='Email', how='left')

    # Format 'ID number' to remove decimals
    pf_leaves['ID number'] = pf_leaves['ID number'].apply(lambda x: '{:.0f}'.format(x))

    # Replace leave type values for consistency
    pf_leaves['Leave Type'] = pf_leaves['Leave Type'].replace({'Work from home': np.nan, 'BirthDay off': 'Paid leave'})

    # Debug: Print unique values in Leave Type column
    print("Unique values in Leave Type column:")
    print(pf_leaves['Leave Type'].unique())

    # Convert date columns to datetime format
    pf_leaves['Starts on'] = pd.to_datetime(pf_leaves['Starts on'])
    pf_leaves['Ends on'] = pd.to_datetime(pf_leaves['Ends on'])

    # Identify the overall date range for the dataset
    start_date = pf_leaves['Starts on'].min()
    end_date = pf_leaves['Ends on'].max()

    # Generate all dates within this range, including the hour part
    all_dates = pd.date_range(start=start_date, end=end_date, freq='D')
    date_columns = [date.strftime('%Y-%m-%d 00:00:00') for date in all_dates]

    # Initialize a DataFrame for flattened leave data
    flattened_leave_data = pd.DataFrame(columns=['Email'] + date_columns)
    flattened_leave_data['Email'] = pf_leaves['Email'].unique()

    # Populate the DataFrame with leave types, considering the hour part in the format
    for _, row in pf_leaves.iterrows():
        date_range = pd.date_range(start=row['Starts on'], end=row['Ends on'], freq='D')
        for date in date_range:
            flattened_leave_data.loc[
                flattened_leave_data['Email'] == row['Email'], date.strftime('%Y-%m-%d 00:00:00')] = row['Leave Type']

    # Prepare the original DataFrame by dropping date range and leave type columns to avoid redundancy
    pf_leaves_reduced = pf_leaves.drop(['Starts on', 'Ends on', 'Leave Type'], axis=1).drop_duplicates(subset=['Email'])

    # Merge the original DataFrame with the new flattened leave data
    merged_df = pd.merge(pf_leaves_reduced, flattened_leave_data, on='Email', how='right')

    pf_leaves = merged_df.copy()

    # Debug: Print sample of pf_leaves data
    print("\nSample of pf_leaves data:")
    print(pf_leaves.head())

    # Define a function to check if a value is numeric or 'OFF'
    def is_numeric_or_off(val):
        try:
            float(val)
            return True
        except ValueError:
            return str(val).strip().upper() == "OFF"

    # Apply the function to clean up numeric values and 'OFF' in shift columns
    for col in shifts.columns[5:]:
        shifts[col] = shifts[col].apply(lambda x: x if is_numeric_or_off(x) else np.nan)

    # Define a function to add leading zeros to ID values if necessary
    def add_leading_zero(id_value):
        if len(str(id_value)) < 11:
            return '0' * (11 - len(str(id_value))) + str(id_value)
        else:
            return str(id_value)

    # Apply the function to the ID columns in main, pf_leaves, and shifts DataFrames
    main['ID'] = main['ID'].astype(str).apply(add_leading_zero)
    pf_leaves['ID number'] = pf_leaves['ID number'].astype(str).apply(add_leading_zero)
    shifts['ID'] = shifts['ID'].astype(str).apply(add_leading_zero)

    # Convert DataFrame column names to strings for consistency
    main.columns = main.columns.map(str)
    pf_leaves.columns = pf_leaves.columns.map(str)

    # Identify common columns, excluding the ID columns
    common_columns = list(set(main.columns) & set(pf_leaves.columns) - {'ID', 'ID number'})

    # Loop through each common column and fill NaN values in 'main' from 'pf_leaves'
    for col in common_columns:
        mapping_dict = pf_leaves.set_index('ID number')[col].dropna().to_dict()
        main[col] = main['ID'].map(mapping_dict).fillna(main[col])

    # Convert DataFrame column names to strings to handle date-like names smoothly
    main.columns = main.columns.map(str)
    shifts.columns = shifts.columns.map(str)

    # Identify common columns, excluding the ID columns
    common_columns = list(set(main.columns) & set(shifts.columns) - {'ID'})

    # Loop through each common column and fill NaN values in 'main' from 'shifts'
    for col in common_columns:
        mapping_dict = shifts.set_index('ID')[col].dropna().to_dict()
        main[col] = main[col].where(main[col].notnull(), main['ID'].map(mapping_dict))

    # Define a function to check if a column name is date-like
    def is_date_like(col_name):
        try:
            pd.to_datetime(col_name)
            return True
        except ValueError:
            return False

    # Filter date-like columns in 'main'
    date_like_columns_main = [col for col in main.columns if is_date_like(col)]

    # Define a function to fill hours based on weekdays or weekends
    def fill_hours_based_on_day(df):
        date_columns = [col for col in df.columns if is_date_like(col)]
        for col_name in date_columns:
            col_date = pd.to_datetime(col_name)
            if col_date.weekday() < 5:
                fill_value = '8'
            else:
                fill_value = 'OFF'
            df[col_name] = df[col_name].apply(lambda x: fill_value if pd.isna(x) else x)

    # Apply the function to 'main'
    fill_hours_based_on_day(main)

    # Debug: Print sample of main data before processing
    print("\nSample of main data before processing:")
    print(main.head())

    # Define a function to calculate row summaries
    def calculate_row_summaries(row, date_columns):
        total_hours_first_half = 0
        total_hours_second_half = 0
        total_hours_month = 0
        total_days_worked = 0
        off_count = 0
        paid_leave_count = 0
        unpaid_leave_count = 0
        maternity_leave_count = 0
        sick_leave_count = 0
        mental_dayoff_count = 0

        print(f"Processing row: {row['ID']}")  # Debug: Print the ID of the row being processed

        for col_name in date_columns:
            day = pd.to_datetime(col_name).day
            value = str(row[col_name]).strip().lower()  # Convert to string, strip whitespace, and lowercase
            print(f"Column: {col_name}, Value: {value}")  # Debug: Print each column and its value

            numeric_value = pd.to_numeric(value, errors='coerce')

            if not pd.isna(numeric_value):
                total_hours_month += numeric_value
                if day <= 15:
                    total_hours_first_half += numeric_value
                else:
                    total_hours_second_half += numeric_value
                total_days_worked += 1
            else:
                if value == "off":
                    off_count += 1
                elif value in ["შვ", "paid leave"]:
                    paid_leave_count += 1
                    print(f"Paid leave found: {col_name}")  # Debug: Print when paid leave is found
                elif value in ["არ.შვ", "unpaid leave"]:
                    unpaid_leave_count += 1
                    print(f"Unpaid leave found: {col_name}")  # Debug: Print when unpaid leave is found
                elif value in ["დეკ", "maternity leave"]:
                    maternity_leave_count += 1
                elif value in ["ბიულ", "sick leave"]:
                    sick_leave_count += 1
                elif value in ["მენ.დღ", "mental dayoff"]:
                    mental_dayoff_count += 1

        print(f"Paid leave count: {paid_leave_count}")  # Debug: Print total paid leave count
        print(f"Unpaid leave count: {unpaid_leave_count}")  # Debug: Print total unpaid leave count

        month = pd.to_datetime(date_columns[0]).strftime('%B')
        row[f'ნამუშევარი საათი 1-15 {month}'] = total_hours_first_half
        row[f'ნამუშევარი საათი 16-{pd.to_datetime(date_columns[-1]).day} {month}'] = total_hours_second_half
        row[f'ნამუშევარი საათი {month}'] = total_hours_month
        row[f'ნამუშევარი დღე {month}'] = total_days_worked
        row['OFF'] = off_count
        row['ანაზღაურებადი შვებულება'] = paid_leave_count
        row['არა ანაზღაურებადი შვებულება'] = unpaid_leave_count
        row['დეკრეტული'] = maternity_leave_count
        row['ბიულეტენი'] = sick_leave_count
        row['Mental Dayoff'] = mental_dayoff_count
        row[
            'სულ არასამუშაო დღე'] = off_count + paid_leave_count + unpaid_leave_count + maternity_leave_count + sick_leave_count + mental_dayoff_count

        return row

    # Apply the function to each row of the DataFrame
    main = main.apply(lambda row: calculate_row_summaries(row, date_like_columns_main), axis=1)

    # Debug: Print sample of main data after processing
    print("\nSample of main data after processing:")
    print(main.head())

    # Drop unnecessary column
    main.drop(columns=['Unnamed: 16'], inplace=True, errors='ignore')

    # Replace leave type values with shorter versions
    replacement_dict = {
        'Paid leave': 'შვ',
        'Unpaid leave': 'არ.შვ',
        'Maternity leave': 'დეკ',
        'Sick leave': 'ბიულ',
        'Mental Dayoff': 'მენ.დღ'
    }
    main = main.replace(replacement_dict)

    # Translate month names to Georgian
    month_mapping = {
        'January': 'იანვარი',
        'February': 'თებერვალი',
        'March': 'მარტი',
        'April': 'აპრილი',
        'May': 'მაისი',
        'June': 'ივნისი',
        'July': 'ივლისი',
        'August': 'აგვისტო',
        'September': 'სექტემბერი',
        'October': 'ოქტომბერი',
        'November': 'ნოემბერი',
        'December': 'დეკემბერი'
    }

    # Replace English month names with Georgian in column headers
    for col in main.columns:
        for eng_month, geo_month in month_mapping.items():
            if eng_month in col:
                main.rename(columns={col: col.replace(eng_month, geo_month)}, inplace=True)

    # Anonymize 'ID' column by replacing last 4 characters with '****'
    main['ID'] = main['ID'].str[:-4] + '****'

    return main


def main():
    st.title("Excel File Processor")

    st.write("Please upload the required Excel files:")

    main_file = st.file_uploader("Upload 'main' Excel file", type=['xlsx', 'xls'])
    pf_id_file = st.file_uploader("Upload 'pf_id' Excel file", type=['xlsx', 'xls'])
    pf_leaves_file = st.file_uploader("Upload 'pf_leaves' Excel file", type=['xlsx', 'xls'])
    shifts_file = st.file_uploader("Upload 'shifts' Excel file", type=['xlsx', 'xls'])

    if main_file and pf_id_file and pf_leaves_file and shifts_file:
        main_df = read_excel_file(main_file)
        pf_id_df = read_excel_file(pf_id_file)
        pf_leaves_df = read_excel_file(pf_leaves_file)
        shifts_df = read_excel_file(shifts_file)

        if main_df is not None and pf_id_df is not None and pf_leaves_df is not None and shifts_df is not None:
            st.write("Processing files...")
            result_df = process_data(main_df, pf_leaves_df, pf_id_df, shifts_df)

            st.write("Processing complete. Preparing download...")

            # Display debug information
            st.subheader("Debug Information")
            st.text("Check the console for more detailed debug output")
            st.write("Sample of processed data:")
            st.write(result_df.head())

            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                result_df.to_excel(writer, index=False, sheet_name='Sheet1')

            output.seek(0)

            st.download_button(
                label="Download processed Excel file",
                data=output,
                file_name="processed_file.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )


if __name__ == "__main__":
    main()