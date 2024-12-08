import gspread
from google.oauth2.service_account import Credentials
import pandas as pd

# Authenticate with Google Sheets
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
credentials = Credentials.from_service_account_file("credentials.json", scopes=scope)
client = gspread.authorize(credentials)

# Open the Google Sheet and access the required tab
sheet = client.open("Baseline_Copy_ASIN Level")
worksheet = sheet.worksheet("Baseline - Dec'24")

# Fetch data from the sheet
data = worksheet.get_all_records()

# Convert to pandas DataFrame
df = pd.DataFrame(data)

# Filter only the relevant columns
df = df[["Brand", "Hero asin/style", "Sales", "Units", "ASP"]]

# Ensure relevant columns are numeric
df["Sales"] = pd.to_numeric(df["Sales"], errors="coerce")
df["Units"] = pd.to_numeric(df["Units"], errors="coerce")
df["ASP"] = pd.to_numeric(df["ASP"], errors="coerce")

# Initialize variables
start_date = "2024-12"
end_date = "2025-03"
num_months = pd.date_range(start=start_date, end=end_date, freq="MS").size  # Calculate months dynamically
monthly_data = []

# Perform calculations and replicate for each month
for i in range(num_months):
    # Compute new month and year
    current_date = pd.to_datetime(start_date) + pd.DateOffset(months=i)
    formatted_date = current_date.strftime("%Y-%m")

    # Adjust Sales and ASP by 0.2%
    df["Sales"] = df["Sales"] * 1.002
    df["ASP"] = df["ASP"] * 1.002

    # Recalculate Units
    df["Units"] = df["Sales"] / df["ASP"]

    # Add the date column
    df["Date"] = formatted_date

    # Append the updated DataFrame to monthly_data
    monthly_data.append(df.copy())

# Concatenate all monthly data
result_df = pd.concat(monthly_data)

# Handle NaN values and infinities
result_df = result_df.fillna(0)  # Replace NaN with 0
result_df = result_df.replace([float("inf"), -float("inf")], 0)  # Replace infinite values with 0

# Check if the "Updated Data" tab exists
new_tab_name = "Updated Data"
try:
    # Try to add the tab
    new_worksheet = sheet.add_worksheet(title=new_tab_name, rows="1000", cols="20")
except gspread.exceptions.APIError:
    # If it exists, clear its contents
    new_worksheet = sheet.worksheet(new_tab_name)
    new_worksheet.clear()

# Update the worksheet with new data
new_worksheet.update([result_df.columns.values.tolist()] + result_df.values.tolist())
