import xlwings as xw
import openpyxl
import pandas as pd
import datetime
import eikon as ek
from openpyxl.styles import PatternFill

# Set your Eikon API key
ek.set_app_key('YOUR_EIKON_API_KEY')

def validate_date(date_str):
    """Validate and parse the date string."""
    try:
        return datetime.datetime.strptime(date_str, '%Y-%m-%d')
    except ValueError:
        return None

def fetch_timeseries_data(instrument, start_date, end_date):
    """Fetch timeseries data from Eikon."""
    try:
        data = ek.get_timeseries(instrument, fields=["*"], start_date=start_date, end_date=end_date, interval="minute")
        return data
    except Exception as e:
        return str(e)

def write_heatmap(sheet, df_heatmap, start_row, start_col, tick_size, metric_name):
    """Generate and write heatmap data to Excel sheet."""
    sheet.range(start_row, start_col - 1).value = metric_name

    df_heatmap = df_heatmap.fillna(0)
    for row_idx, (date, row_data) in enumerate(df_heatmap.iterrows(), start=start_row + 1):
        for col_idx, value in enumerate(row_data, start=start_col):
            sheet.range(row_idx, col_idx).value = round(value / tick_size, 2) if tick_size else value

    for idx, date in enumerate(df_heatmap.index, start=start_row + 1):
        sheet.range(idx, start_col - 1).value = date
    for idx, time in enumerate(df_heatmap.columns, start=start_col):
        sheet.range(start_row, idx).value = time

    # Apply formatting (e.g., color mapping) here as needed

def main():
    wb = xw.Book(r'C:\Users\YASH\Downloads\Commodity_Market_Analytics_Dashboard-main\Dashbd_visualN.xlsx')
    sheet = wb.sheets['Dashboard']

    while True:
        try:
            # Check if the script should execute
            execute_flag = sheet.range('B1').value
            if execute_flag != 1:
                continue

            # Read inputs
            instrument = sheet.range('B2').value
            start_date = sheet.range('B3').value
            end_date = sheet.range('B4').value

            # Validate dates
            start_date = validate_date(start_date)
            end_date = validate_date(end_date)

            if not start_date or not end_date:
                sheet.range('B6').value = "Invalid date format. Use YYYY-MM-DD."
                continue

            # Fetch data
            data = fetch_timeseries_data(instrument, start_date, end_date)

            if isinstance(data, str):  # Error occurred
                sheet.range('B6').value = f"Error fetching data: {data}"
                continue

            # Process data for heatmap
            heatmap_metrics = ["Volume", "Range", "Change"]  # Example metrics
            start_row, start_col = 10, 2
            tick_size = 0.01  # Example tick size

            for metric in heatmap_metrics:
                metric_data = data[[metric]]  # Adjust this to match your data structure
                write_heatmap(sheet, metric_data, start_row, start_col, tick_size, metric)
                start_row += len(metric_data) + 5

            sheet.range('B6').value = "Heatmap generation complete."

        except Exception as e:
            sheet.range('B6').value = f"Unexpected error: {str(e)}"

if __name__ == "__main__":
    main()
