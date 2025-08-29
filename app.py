import streamlit as st
import pandas as pd
import io
from datetime import datetime
import re

def parse_for_sorting(date_str):
    """
    Parses any known date format into a datetime object ONLY for sorting.
    Returns NaT on failure. This handles the dd/mm vs mm/dd conflict.
    """
    if pd.isna(date_str) or str(date_str).strip().upper() in ['TBA', 'NAN']:
        return pd.NaT
    
    date_str = str(date_str).strip()

    # 1. Handle ranges first, as they are the most specific format (mm/dd.../yyyy)
    if '-' in date_str and '/' in date_str:
        try:
            start_date_part = date_str.split('-')[0]
            year_part = date_str.split('/')[-1]
            full_start_date = f"{start_date_part}/{year_part}"
            return pd.to_datetime(full_start_date, format='%m/%d/%Y')
        except (ValueError, TypeError):
            return pd.NaT

    # 2. Handle YYYY-MM-DD format
    if '-' in date_str and date_str.count('-') == 2:
        try:
            date_part = date_str.split(' ')[0]
            return pd.to_datetime(date_part, format='%Y-%m-%d')
        except (ValueError, TypeError):
            return pd.NaT

    # 3. Handle all other slash formats (dd/mm/yy or dd/mm/yyyy)
    try:
        # dayfirst=True correctly interprets 05/09/25 as September 5th
        return pd.to_datetime(date_str, dayfirst=True)
    except (ValueError, TypeError):
        return pd.NaT

def parse_time_string(time_str):
    """
    Parse time strings into a consistent format, handling TBA.
    """
    if pd.isna(time_str) or str(time_str).strip().upper() in ['TBA', 'NAN', '']:
        return 'TBA'
    
    time_str = str(time_str).strip()
    
    # Handle excel's time object conversion to string "17:00:00"
    if re.match(r'^\d{2}:\d{2}:\d{2}$', time_str):
        try:
            return pd.to_datetime(time_str, format='%H:%M:%S').strftime('%I:%M %p')
        except (ValueError, TypeError):
            pass # Fallback to other formats
            
    time_formats = ['%I:%M %p', '%H:%M', '%I %p', '%I:%M%p']
    for fmt in time_formats:
        try:
            return pd.to_datetime(time_str, format=fmt).strftime('%I:%M %p')
        except (ValueError, TypeError):
            continue
    
    return time_str # Return original if no format matches

def get_sort_time(time_str):
    """
    Convert time string to a full datetime for sorting. TBA times are sent to the end.
    """
    if time_str == 'TBA':
        return pd.Timestamp.max # Send TBA times to the very end of any sort
    
    try:
        # Use pandas to parse time, which is more robust
        return pd.to_datetime(time_str, format='%I:%M %p')
    except (ValueError, TypeError):
        return pd.Timestamp.min # Send unparsed times to the beginning

def find_required_columns(df):
    """
    Find required columns, handling variations in names.
    """
    column_mapping = {
        'date': ['date', 'dates'],
        'time': ['time', 'times'],
        'opponent': ['opponent', 'opponents', 'vs', 'against'],
        'meet': ['meet', 'meets', 'event', 'competition'],
        'location': ['location', 'locations', 'venue', 'where', 'place'],
        'distance': ['distance from macu', 'distance', 'miles', 'distance (miles)']
    }
    found_columns = {}
    for standard_name, variations in column_mapping.items():
        for variation in variations:
            matching_cols = [col for col in df.columns if variation in col.lower().strip()]
            if matching_cols:
                found_columns[standard_name] = matching_cols[0]
                break
    return found_columns

def create_match_count_summary(df):
    """
    Create a summary showing match counts by date.
    """
    if df is None or df.empty:
        return None
    
    # Group by date and count matches
    date_counts = df.groupby('Date').size().reset_index(name='Match Count')
    
    # Sort by the original date order from the main dataframe
    unique_dates = df['Date'].unique()
    date_counts['sort_order'] = date_counts['Date'].map({date: i for i, date in enumerate(unique_dates)})
    date_counts = date_counts.sort_values('sort_order').drop('sort_order', axis=1)
    
    return date_counts

def combine_and_sort_schedules(uploaded_files):
    """
    Combines, correctly sorts, and displays data while preserving original date formats.
    """
    list_of_dataframes = []
    
    for file in uploaded_files:
        try:
            df = pd.read_excel(file, header=0, dtype=str).dropna(how='all')
            column_mapping = find_required_columns(df)
            
            if 'date' not in column_mapping:
                st.warning(f"Skipping file `{file.name}` - missing a 'Date' column.")
                continue

            df['Display Date'] = df[column_mapping.get('date')]
            df['MACU Team'] = file.name.rsplit('.', 1)[0]
            
            rename_dict = {v: k for k, v in column_mapping.items()}
            df.rename(columns=rename_dict, inplace=True)
            
            df['sort_date'] = df['date'].apply(parse_for_sorting)
            df['parsed_time'] = df['time'].apply(parse_time_string)
            df['sort_time'] = df['parsed_time'].apply(get_sort_time)
            
            df['sort_date'].fillna(pd.Timestamp.max, inplace=True)

            list_of_dataframes.append(df)
            
        except Exception as e:
            st.error(f"Error processing `{file.name}`: {e}")
            continue
    
    if not list_of_dataframes:
        return None
    
    combined_df = pd.concat(list_of_dataframes, ignore_index=True)
    combined_df.sort_values(by=['sort_date', 'sort_time'], inplace=True)

    def format_final_date(row):
        original_date = str(row['Display Date'])
        parsed_date = row['sort_date']
        
        if '-' in original_date and '/' in original_date:
            return original_date
        elif pd.notna(parsed_date) and parsed_date != pd.Timestamp.max:
            return parsed_date.strftime('%d/%m/%Y')
        else:
            return original_date

    combined_df['Final Date'] = combined_df.apply(format_final_date, axis=1)

    # --- FIX for ArrowTypeError and Column Name Change ---
    final_df = pd.DataFrame({
        'Date': combined_df['Final Date'],
        'Time': combined_df['parsed_time'].fillna('TBA'),
        'MACU Team': combined_df['MACU Team'].fillna(''),
        'Opponent': combined_df.get('opponent', '').fillna(''),
        'Meet': combined_df.get('meet', '').fillna(''),
        'Location': combined_df.get('location', '').fillna('TBA'),
        # New column name and conversion to string to fix the error
        'Distance from MACU (miles)': combined_df.get('distance', '0').fillna('0').astype(str)
    })
    
    return final_df

# --- Streamlit UI ---
st.set_page_config(page_title="MACU Schedule Combiner", layout="wide")
st.title("MACU Athletics Schedule Combiner 🗓️")
st.markdown("Upload multiple team schedule Excel files to combine them into a single, sorted master schedule.")

uploaded_files = st.file_uploader(
    "Choose Excel files (.xlsx, .xls)", 
    type=['xlsx', 'xls'], 
    accept_multiple_files=True
)

if uploaded_files:
    with st.spinner('Processing files and building schedule...'):
        combined_schedule = combine_and_sort_schedules(uploaded_files)
    
    if combined_schedule is not None and not combined_schedule.empty:
        # Create and display match count summary
        match_count_summary = create_match_count_summary(combined_schedule)
        
        col1, col2 = st.columns([1, 2])
        
        with col1:
            st.subheader("📊 Matches per Date")
            if match_count_summary is not None:
                st.dataframe(
                    match_count_summary, 
                    hide_index=True, 
                    use_container_width=True,
                    column_config={
                        "Date": st.column_config.TextColumn("Date"),
                        "Match Count": st.column_config.NumberColumn("Matches", format="%d")
                    }
                )
                
                # Show total matches
                total_matches = match_count_summary['Match Count'].sum()
                st.metric("Total Matches", total_matches)
        
        with col2:
            st.subheader("🗓️ Complete Schedule")
            # --- FIX for Deprecation Warning ---
            st.dataframe(combined_schedule, hide_index=True, use_container_width=True)
        
        # --- FIX for Excel Download ---
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            # Write the main schedule
            combined_schedule.to_excel(writer, index=False, sheet_name='Schedule')
            # Write the match count summary to a separate sheet
            if match_count_summary is not None:
                match_count_summary.to_excel(writer, index=False, sheet_name='Match Counts')

        st.download_button(
            label="📥 Download Schedule as Excel",
            data=output,
            file_name=f"combined_macu_schedule_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        # Display some useful statistics
        st.subheader("📈 Schedule Statistics")
        stats_col1, stats_col2, stats_col3 = st.columns(3)
        
        with stats_col1:
            unique_teams = combined_schedule['MACU Team'].nunique()
            st.metric("Teams", unique_teams)
        
        with stats_col2:
            unique_dates = combined_schedule['Date'].nunique()
            st.metric("Competition Days", unique_dates)
        
        with stats_col3:
            avg_matches_per_date = round(total_matches / unique_dates, 1) if unique_dates > 0 else 0
            st.metric("Avg Matches/Day", avg_matches_per_date)
            
    else:
        st.error("No valid schedule data could be extracted. Please check file formats and column names.")
else:
    st.info("Please upload one or more Excel files to get started.")