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

def validate_and_fix_date_columns(main, pf_leaves, shifts):
    """
    Validates that date columns across all input files use consistent years.
    Auto-fixes any mismatches by updating to the correct year from pf_leaves.
    Returns the fixed dataframes and shows warnings if fixes were applied.
    """
    # Get years from each source
    main_date_cols = [col for col in main.columns if isinstance(col, pd.Timestamp) or hasattr(col, 'year')]
    shifts_date_cols = [col for col in shifts.columns if isinstance(col, pd.Timestamp) or hasattr(col, 'year')]

    main_year = None
    shifts_year = None
    pf_leaves_year = None

    if main_date_cols:
        main_year = getattr(main_date_cols[0], 'year', None)

    if shifts_date_cols:
        shifts_year = getattr(shifts_date_cols[0], 'year', None)

    if 'Starts on' in pf_leaves.columns and not pf_leaves.empty:
        pf_leaves_year = pf_leaves['Starts on'].min().year

    # Use pf_leaves year as the source of truth
    target_year = pf_leaves_year

    if target_year is None:
        # No dates to validate against
        return main, shifts

    # Check for mismatches and fix
    warnings = []

    def fix_columns(df, df_name, current_year):
        if current_year and current_year != target_year:
            new_columns = []
            fixed_count = 0

            for col in df.columns:
                if isinstance(col, pd.Timestamp) or hasattr(col, 'year'):
                    if hasattr(col, 'year') and col.year != target_year:
                        new_col = col.replace(year=target_year)
                        new_columns.append(new_col)
                        fixed_count += 1
                    else:
                        new_columns.append(col)
                else:
                    new_columns.append(col)

            if fixed_count > 0:
                df = df.copy()
                df.columns = new_columns
                warnings.append(f"{df_name}: Fixed {fixed_count} date columns from {current_year} to {target_year}")

            return df
        return df

    main = fix_columns(main, "Main file", main_year)
    shifts = fix_columns(shifts, "Shifts file", shifts_year)

    if warnings:
        warning_msg = "⚠️ Date column years were automatically corrected:\n" + "\n".join(f"  • {w}" for w in warnings)
        st.warning(warning_msg)
        print(f"\n{warning_msg}")

    return main, shifts

def fill_hours_based_on_day(df):
    date_columns = df.columns[17:]  # Assuming date columns start from index 17
    for col_name in date_columns:
        col_date = pd.to_datetime(col_name, format='%Y-%m-%d %H:%M:%S')
        fill_value = '8' if col_date.weekday() < 5 else 'OFF'
        df[col_name] = df[col_name].apply(lambda x: fill_value if pd.isna(x) else x)

def calculate_row_summaries(row, date_columns, month_name='MONTH'):
    totals = {'first_half': 0, 'second_half': 0, 'month': 0, 'days_worked': 0}
    counts = {'OFF': 0, 'Paid leave': 0, 'Unpaid leave': 0, 'Maternity leave': 0, 'Sick leave': 0}

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
        elif value in ['Mental Day Off', 'BirthDay off']:
            # Count Mental Day Off and BirthDay off as Paid leave in summary
            counts['Paid leave'] += 1

    row[f'ნამუშევარი საათი 1-15 {month_name}'] = totals['first_half']
    row[f'ნამუშევარი საათი 16-31 {month_name}'] = totals['second_half']
    row[f'ნამუშევარი საათი {month_name}'] = totals['month']
    row[f'ნამუშევარი დღე {month_name}'] = totals['days_worked']
    row['OFF'] = counts['OFF']
    row['ანაზღაურებადი შვებულება'] = counts['Paid leave']
    row['არა ანაზღაურებადი შვებულება'] = counts['Unpaid leave']
    row['დეკრეტული'] = counts['Maternity leave']
    row['ბიულეტენი'] = counts['Sick leave']
    row['სულ არასამუშაო დღე'] = sum(counts.values())

    return row

def process_data(main, pf_leaves, pf_id, shifts):
    # Merge and process data
    pf_leaves = pd.merge(pf_leaves, pf_id[['Email', 'ID number']], on='Email', how='left')
    pf_leaves['ID number'] = pf_leaves['ID number'].apply(lambda x: '{:.0f}'.format(x))

    # Replace leave type values for consistency
    # Note: Mental Day Off and BirthDay off are kept as separate values in date columns
    # but counted together with Paid leave in summary columns
    pf_leaves['Leave Type'] = pf_leaves['Leave Type'].replace({
        'Work from home': np.nan
    })

    # Debug print
    print("Sample dates from pf_leaves before conversion:")
    print(pf_leaves['Starts on'].head())
    print(pf_leaves['Ends on'].head())

    pf_leaves['Starts on'] = pd.to_datetime(pf_leaves['Starts on'], format='mixed', dayfirst=True)
    pf_leaves['Ends on'] = pd.to_datetime(pf_leaves['Ends on'], format='mixed', dayfirst=True)

    # Validate and fix date column years to prevent mismatches
    main, shifts = validate_and_fix_date_columns(main, pf_leaves, shifts)

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

    # Determine the current month from the date columns first
    date_columns = [col for col in main.columns if is_date_like(col)]

    month_mapping = {
        'January': 'იანვარი', 'February': 'თებერვალი', 'March': 'მარტი',
        'April': 'აპრილი', 'May': 'მაისი', 'June': 'ივნისი',
        'July': 'ივლისი', 'August': 'აგვისტო', 'September': 'სექტემბერი',
        'October': 'ოქტომბერი', 'November': 'ნოემბერი', 'December': 'დეკემბერი'
    }

    if date_columns:
        current_month = pd.to_datetime(date_columns[0]).strftime('%B')
        current_month_georgian = month_mapping.get(current_month, current_month)
    else:
        current_month_georgian = 'Unknown'

    # Calculate row summaries with correct month name
    main = main.apply(lambda row: calculate_row_summaries(row, date_columns, current_month_georgian), axis=1)

    # Drop unnecessary column and replace leave type values
    main.drop(columns=['Unnamed: 16'], inplace=True, errors='ignore')

    # Normalize 'off' to 'OFF' for consistency
    main = main.replace({'off': 'OFF'})

    replacement_dict = {
        'Paid leave': 'შვ',
        'Unpaid leave': 'არ.შვ',
        'Maternity leave': 'დეკ',
        'Sick leave': 'ბიულ',
        'BirthDay off': 'Birthday',
        'Mental Day Off': 'Mental'
    }
    main = main.replace(replacement_dict)

    # Reorder columns: move summary columns right after პოზიცია
    summary_cols = [
        f'ნამუშევარი საათი 1-15 {current_month_georgian}',
        f'ნამუშევარი საათი 16-31 {current_month_georgian}',
        f'ნამუშევარი საათი {current_month_georgian}',
        f'ნამუშევარი დღე {current_month_georgian}',
        'OFF',
        'ანაზღაურებადი შვებულება',
        'არა ანაზღაურებადი შვებულება',
        'დეკრეტული',
        'ბიულეტენი',
        'სულ არასამუშაო დღე'
    ]

    # Get position index
    position_index = main.columns.get_loc('პოზიცია')

    # Get columns before პოზიცია (inclusive)
    cols_before = main.columns[:position_index + 1].tolist()

    # Get date columns
    date_cols = [col for col in main.columns if is_date_like(col)]

    # Get any remaining columns that are not in the above lists
    remaining_cols = [col for col in main.columns
                     if col not in cols_before
                     and col not in summary_cols
                     and col not in date_cols]

    # Reorder: before + summary + date columns + remaining
    new_order = cols_before + summary_cols + date_cols + remaining_cols
    main = main[new_order]

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

def identify_file_type(df, filename):
    """
    Identifies the type of uploaded file based on column names and filename.
    Returns: 'main', 'pf_id', 'pf_leaves', 'shifts', or None
    """
    columns = set(df.columns)
    filename_lower = filename.lower()

    # Check by filename first
    if 'main' in filename_lower:
        return 'main'
    elif 'pf_id' in filename_lower or 'peopleforce-id' in filename_lower or 'id' in filename_lower:
        return 'pf_id'
    elif 'leave' in filename_lower:
        return 'pf_leaves'
    elif 'shift' in filename_lower:
        return 'shifts'

    # Check by columns
    # pf_leaves: has Leave Type, Starts on, Ends on
    if 'Leave Type' in columns and 'Starts on' in columns and 'Ends on' in columns:
        return 'pf_leaves'

    # pf_id: has Email, ID number, Position (but no Leave Type)
    if 'Email' in columns and 'ID number' in columns and 'Position' in columns:
        return 'pf_id'

    # shifts: has ID and date columns, but not სახელი, გვარი
    has_id = 'ID' in columns
    has_georgian_names = 'სახელი' in columns and 'გვარი' in columns
    has_date_cols = any(isinstance(col, pd.Timestamp) for col in df.columns)

    if has_id and has_date_cols and not has_georgian_names:
        return 'shifts'

    # main: has ID, სახელი, გვარი, პოზიცია, and date columns
    if has_id and has_georgian_names and 'პოზიცია' in columns and has_date_cols:
        return 'main'

    return None

# Streamlit app
st.title('Data Processing App')

st.write("""
This app processes four Excel files and generates a final Excel file.
Upload all four required files at once, and the system will automatically detect which is which.
""")

# Single file uploader for multiple files
uploaded_files = st.file_uploader("Upload all Excel files (main, pf_id, pf_leaves, shifts)",
                                   type=['xlsx'],
                                   accept_multiple_files=True)

if uploaded_files and len(uploaded_files) >= 4:
    try:
        # Identify and categorize uploaded files
        st.write("### Identifying files...")
        file_map = {}
        identified = []

        for uploaded_file in uploaded_files:
            df = read_excel_file(uploaded_file)
            file_type = identify_file_type(df, uploaded_file.name)

            if file_type:
                file_map[file_type] = df
                identified.append(f"✓ {uploaded_file.name} → {file_type}")
            else:
                identified.append(f"✗ {uploaded_file.name} → Could not identify")

        # Display identification results
        for msg in identified:
            st.write(msg)

        # Check if we have all required files
        required_files = ['main', 'pf_id', 'pf_leaves', 'shifts']
        missing_files = [f for f in required_files if f not in file_map]

        if missing_files:
            st.error(f"Missing files: {', '.join(missing_files)}")
            st.write("Please upload all 4 required file types.")
        else:
            st.success("All files identified successfully!")

            # Process the data
            st.write("### Processing data...")
            processed_data = process_data(
                file_map['main'],
                file_map['pf_leaves'],
                file_map['pf_id'],
                file_map['shifts']
            )

            # Display a sample of the processed data
            st.write("### Sample of processed data:")
            st.dataframe(processed_data.head())

            # Provide download link
            st.write("### Download processed file:")
            st.markdown(get_table_download_link(processed_data), unsafe_allow_html=True)

    except Exception as e:
        st.error(f"An error occurred while processing the files: {str(e)}")
        import traceback
        st.code(traceback.format_exc())
elif uploaded_files and len(uploaded_files) < 4:
    st.warning(f"Please upload at least 4 files. You have uploaded {len(uploaded_files)} file(s).")
else:
    st.info("👆 Please upload all 4 required Excel files to process the data.")