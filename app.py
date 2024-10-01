import streamlit as st
import pandas as pd
import numpy as np
import io
import base64

# Helper functions
def add_leading_zero(id_value):
    if len(str(id_value)) < 11:
        return '0' * (11 - len(str(id_value))) + str(id_value)
    else:
        return str(id_value)

def is_numeric_or_off(val):
    if pd.isna(val):
        return True
    try:
        float(val)
        return True
    except ValueError:
        return str(val).upper() == "OFF"

def is_date_like(col_name):
    try:
        pd.to_datetime(col_name, format='mixed', dayfirst=True)
        return True
    except ValueError:
        return False

def fill_hours_based_on_day(df):
    date_columns = df.columns[17:]  # Assuming date columns start from index 17
    for col_name in date_columns:
        col_date = pd.to_datetime(col_name, format='%Y-%m-%d %H:%M:%S')
        fill_value = '8' if col_date.weekday() < 5 else 'OFF'
        df[col_name] = df[col_name].apply(lambda x: fill_value if pd.isna(x) else x)

def calculate_row_summaries(row, date_columns):
    totals = {'first_half': 0, 'second_half': 0, 'month': 0, 'days_worked': 0}
    counts = {'OFF': 0, 'Paid leave': 0, 'Unpaid leave': 0, 'Maternity leave': 0, 'Sick leave': 0, 'Mental Day Off': 0}

    for col_name in date_columns:
        day = pd.to_datetime(col_name, format='%Y-%m-%d %H:%M:%S').day
        value = row[col_name]
        numeric_value = pd.to_numeric(value, errors='coerce')

        if not pd.isna(numeric_value):
            totals['month'] += numeric_value
            totals['first_half' if day <= 15 else 'second_half'] += numeric_value
            totals['days_worked'] += 1
        elif value in counts:
            counts[value] += 1

    row['ნამუშევარი საათი 1-15 მარტი'] = totals['first_half']
    row['ნამუშევარი საათი 16-31 მარტი'] = totals['second_half']
    row['ნამუშევარი საათი მარტი'] = totals['month']
    row['ნამუშევარი დღე მარტი'] = totals['days_worked']
    row['OFF'] = counts['OFF']
    row['ანაზღაურებადი შვებულება'] = counts['Paid leave']
    row['არა ანაზღაურებადი შვებულება'] = counts['Unpaid leave']
    row['დეკრეტული'] = counts['Maternity leave']
    row['ბიულეტენი'] = counts['Sick leave']
    row['Mental Day Off'] = counts['Mental Day Off']
    row['სულ არასამუშაო დღე'] = sum(counts.values())

    return row

def process_data(main, pf_leaves, pf_id, shifts):
    # Merge and process data
    pf_leaves = pd.merge(pf_leaves, pf_id[['Email', 'ID number']], on='Email', how='left')
    pf_leaves['ID number'] = pf_leaves['ID number'].apply(lambda x: '{:.0f}'.format(x))

    # Replace leave type values for consistency
    pf_leaves['Leave Type'] = pf_leaves['Leave Type'].replace({
        'Work from home': np.nan,
        'BirthDay off': 'Paid leave',
        'Mental Day Off': 'Mental Day Off'
    })

    # Debug print
    print("Sample dates from pf_leaves before conversion:")
    print(pf_leaves['Starts on'].head())
    print(pf_leaves['Ends on'].head())

    pf_leaves['Starts on'] = pd.to_datetime(pf_leaves['Starts on'], format='mixed', dayfirst=True)
    pf_leaves['Ends on'] = pd.to_datetime(pf_leaves['Ends on'], format='mixed', dayfirst=True)

    # Debug print
    print("Sample dates from pf_leaves after conversion:")
    print(pf_leaves['Starts on'].head())
    print(pf_leaves['Ends on'].head())

    # Generate date range and flatten leave data
    start_date = pf_leaves['Starts on'].min()
    end_date = pf_leaves['Ends on'].max()
    all_dates = pd.date_range(start=start_date, end=end_date, freq='D')
    date_columns = [date.strftime('%Y-%m-%d 00:00:00') for date in all_dates]

    flattened_leave_data = pd.DataFrame(columns=['Email'] + date_columns)
    flattened_leave_data['Email'] = pf_leaves['Email'].unique()

    for _, row in pf_leaves.iterrows():
        date_range = pd.date_range(start=row['Starts on'], end=row['Ends on'], freq='D')
        for date in date_range:
            flattened_leave_data.loc[
                flattened_leave_data['Email'] == row['Email'], date.strftime('%Y-%m-%d 00:00:00')] = row['Leave Type']

    pf_leaves_reduced = pf_leaves.drop(['Starts on', 'Ends on', 'Leave Type'], axis=1).drop_duplicates(subset=['Email'])
    merged_df = pd.merge(pf_leaves_reduced, flattened_leave_data, on='Email', how='right')
    pf_leaves = merged_df.copy()

    # Clean up shifts data
    for col in shifts.columns[5:]:
        shifts[col] = shifts[col].apply(lambda x: x if is_numeric_or_off(x) else np.nan)

    # Add leading zeros to ID values
    main['ID'] = main['ID'].astype(str).apply(add_leading_zero)
    pf_leaves['ID number'] = pf_leaves['ID number'].astype(str).apply(add_leading_zero)
    shifts['ID'] = shifts['ID'].astype(str).apply(add_leading_zero)

    # Convert column names to strings
    main.columns = main.columns.map(str)
    pf_leaves.columns = pf_leaves.columns.map(str)
    shifts.columns = shifts.columns.map(str)

    # Fill NaN values in 'main' from 'pf_leaves' and 'shifts'
    common_columns = list(set(main.columns) & set(pf_leaves.columns) - {'ID', 'ID number'})
    for col in common_columns:
        mapping_dict = pf_leaves.set_index('ID number')[col].dropna().to_dict()
        main[col] = main['ID'].map(mapping_dict).fillna(main[col])

    common_columns = list(set(main.columns) & set(shifts.columns) - {'ID'})
    for col in common_columns:
        mapping_dict = shifts.set_index('ID')[col].dropna().to_dict()
        main[col] = main[col].where(main[col].notnull(), main['ID'].map(mapping_dict))

    # Fill hours based on weekdays or weekends
    fill_hours_based_on_day(main)

    # Calculate row summaries
    date_columns = [col for col in main.columns if is_date_like(col)]
    main = main.apply(lambda row: calculate_row_summaries(row, date_columns), axis=1)

    # Drop unnecessary column and replace leave type values
    main.drop(columns=['Unnamed: 16'], inplace=True, errors='ignore')
    replacement_dict = {
        'Paid leave': 'შვ',
        'Unpaid leave': 'არ.შვ',
        'Maternity leave': 'დეკ',
        'Sick leave': 'ბიულ',
        'Mental Day Off': 'Mental Day Off'
    }
    main = main.replace(replacement_dict)

    # Translate month names to Georgian
    month_mapping = {
        'January': 'იანვარი', 'February': 'თებერვალი', 'March': 'მარტი',
        'April': 'აპრილი', 'May': 'მაისი', 'June': 'ივნისი',
        'July': 'ივლისი', 'August': 'აგვისტო', 'September': 'სექტემბერი',
        'October': 'ოქტომბერი', 'November': 'ნოემბერი', 'December': 'დეკემბერი'
    }

    # Determine the current month from the date columns
    if date_columns:
        current_month = pd.to_datetime(date_columns[0]).strftime('%B')
        current_month_georgian = month_mapping.get(current_month, current_month)
    else:
        current_month_georgian = 'Unknown'

    # Replace 'მარტი' with the actual month name in specific columns
    position_index = main.columns.get_loc('პოზიცია')
    columns_to_update = main.columns[position_index + 1:position_index + 5]
    for col in columns_to_update:
        new_col_name = col.replace('მარტი', current_month_georgian)
        main.rename(columns={col: new_col_name}, inplace=True)

    # Anonymize 'ID' column
    main['ID'] = main['ID'].str[:-4] + '****'

    return main

def get_table_download_link(df):
    """Generates a link allowing the data in a given panda dataframe to be downloaded"""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    b64 = base64.b64encode(output.getvalue()).decode()
    return f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="processed_data.xlsx">Download Excel file</a>'

def read_excel_file(file):
    return pd.read_excel(file, engine='openpyxl')

# Streamlit app
st.title('Data Processing App')

st.write("""
This app processes four Excel files and generates a final Excel file.
Please upload the required files below.
""")

# File uploaders
main_file = st.file_uploader("Upload main file", type=['xlsx'])
pf_id_file = st.file_uploader("Upload pf_id file", type=['xlsx'])
pf_leaves_file = st.file_uploader("Upload pf_leaves file", type=['xlsx'])
shifts_file = st.file_uploader("Upload shifts file", type=['xlsx'])

if main_file and pf_id_file and pf_leaves_file and shifts_file:
    # Read the uploaded files
    try:
        main = read_excel_file(main_file)
        pf_id = read_excel_file(pf_id_file)
        pf_leaves = read_excel_file(pf_leaves_file)
        shifts = read_excel_file(shifts_file)

        # Process the data
        processed_data = process_data(main, pf_leaves, pf_id, shifts)

        # Display a sample of the processed data
        st.write("Sample of processed data:")
        st.dataframe(processed_data.head())

        # Provide download link
        st.markdown(get_table_download_link(processed_data), unsafe_allow_html=True)
    except Exception as e:
        st.error(f"An error occurred while processing the files: {str(e)}")
else:
    st.write("Please upload all required files to process the data.")