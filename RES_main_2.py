import pandas as pd

# Load the Excel file into a pandas DataFrame.
df = pd.read_excel("RES_P11.xlsx")

# Decode column names from 'latin1' to 'utf-8' to handle potential encoding issues,
# especially useful for non-ASCII characters.
df.columns = [col.encode('latin1').decode('utf-8', errors='ignore') if isinstance(col, str) else col for col in df.columns]

# Convert the 'Time' column to datetime objects.
# 'errors='coerce'' will turn any unparseable dates into NaT (Not a Time).
df['Time'] = pd.to_datetime(df['Time'], errors='coerce')

# Set the 'Time' column as the DataFrame's index. This is useful for time-series operations.
df = df.set_index('Time')

# Define a dictionary to map original column names (from the Excel file)
# to new, more descriptive or translated column names for the output.
target_columns = {
    'Generic flat plate PV Power Output': 'خروجی سلول خورشیدی',
    'Generic 1kWh Lead Acid Input Power': 'ورودی باتری',
    'Total Electrical Load Served': 'بار',
}

# Define a dictionary to categorize months into seasons.
# Each season name is a key, and its value is a list of corresponding month numbers.
season_months = {
    "Winter": [1, 2, 3],
    "Spring": [4, 5, 6],
    "Summer": [7, 8, 9],
    "Autumn": [10, 11, 12]
}

# Create an ExcelWriter object to write multiple DataFrames to different sheets
# within a single Excel file named "Result_P11.xlsx".
with pd.ExcelWriter("Result_P11.xlsx") as writer:
    # Iterate through each season defined in the 'season_months' dictionary.
    for season_name, months_in_season in season_months.items():
        # Filter the DataFrame to include only the data for the current season's months.
        season_data = df[df.index.month.isin(months_in_season)]

        # Create a copy of the filtered data to avoid SettingWithCopyWarning,
        # ensuring modifications affect only this seasonal DataFrame.
        season_data = season_data.copy()
        
        # Extract the hour component from the 'Time' index and add it as a new 'Hour' column.
        season_data["Hour"] = season_data.index.hour

        # Initialize an empty DataFrame to store the hourly mean results for the current season.
        result_df = pd.DataFrame()

        # Iterate through the target columns defined earlier.
        for orig_col, new_col in target_columns.items():
            # Check if the original column exists in the current season's data.
            if orig_col in season_data.columns:
                # Group the seasonal data by 'Hour' and calculate the mean for the current target column.
                hourly_mean = season_data.groupby("Hour")[orig_col].mean()

                # Reindex the hourly_mean Series to ensure all 24 hours (0-23) are present.
                # If an hour has no data, its value will be None.
                hourly_mean = hourly_mean.reindex(range(24), fill_value=None)

                # Add the calculated hourly means to the result_df under the new column name.
                result_df[new_col] = hourly_mean.values
            else:
                # If a target column is not found in the data, print a message
                # and add a column of None values for that column in the result_df.
                print(f"The {orig_col} column is not found.")
                result_df[new_col] = [None]*24

        # Set the index of the result_df to a list of integers from 0 to 23, representing hours.
        result_df.index = list(range(24))
        
        # Name the index column 'Hour'.
        result_df.index.name = "Hour"

        # Write the result_df for the current season to a new sheet in the Excel file,
        # with the sheet name being the season's name.
        result_df.to_excel(writer, sheet_name=season_name)

# Print a success message after processing the first Excel file.
print("✅11")








# ---
# This sections of the code performs an identical set of operations as the first section,
# but it processes a different input file ("RES_P112.xlsx") and includes additional
# target columns for analysis.
df = pd.read_excel("RES_P112.xlsx")

df.columns = [col.encode('latin1').decode('utf-8', errors='ignore') if isinstance(col, str) else col for col in df.columns]
df['Time'] = pd.to_datetime(df['Time'], errors='coerce')
df = df.set_index('Time')

target_columns = {
    'Generic flat plate PV Power Output': 'خروجی سلول خورشیدی',
    'Generic 1kWh Lead Acid Input Power': 'ورودی باتری',
    'Total Electrical Load Served': 'بار',
    'Excess Electrical Production': 'فروش به شبکه',
    'Grid Purchases': 'خرید از شبکه',
}

season_months = {
    "Winter": [1, 2, 3],
    "Spring": [4, 5, 6],
    "Summer": [7, 8, 9],
    "Autumn": [10, 11, 12]
}

with pd.ExcelWriter("Result_P112.xlsx") as writer:
    for season_name, months_in_season in season_months.items():
        season_data = df[df.index.month.isin(months_in_season)]

        season_data = season_data.copy()
        season_data["Hour"] = season_data.index.hour

        result_df = pd.DataFrame()

        for orig_col, new_col in target_columns.items():
            if orig_col in season_data.columns:
                hourly_mean = season_data.groupby("Hour")[orig_col].mean()

                hourly_mean = hourly_mean.reindex(range(24), fill_value=None)

                result_df[new_col] = hourly_mean.values
            else:
                print(f"The {orig_col} column is not found.")
                result_df[new_col] = [None]*24

        result_df.index = list(range(24))
        result_df.index.name = "Hour"

        result_df.to_excel(writer, sheet_name=season_name)

print("✅112")











df = pd.read_excel("RES_P113.xlsx")

df.columns = [col.encode('latin1').decode('utf-8', errors='ignore') if isinstance(col, str) else col for col in df.columns]
df['Time'] = pd.to_datetime(df['Time'], errors='coerce')
df = df.set_index('Time')

target_columns = {
    'Generic 3 kW Power Output': 'خروجی توربین بادی',
    'Generic 1kWh Lead Acid Input Power': 'ورودی باتری',
    'Total Electrical Load Served': 'بار',
    # 'Excess Electrical Production': 'فروش به شبکه',
    # 'Grid Purchases': 'خرید از شبکه',
}

season_months = {
    "Winter": [1, 2, 3],
    "Spring": [4, 5, 6],
    "Summer": [7, 8, 9],
    "Autumn": [10, 11, 12]
}

with pd.ExcelWriter("Result_P113.xlsx") as writer:
    for season_name, months_in_season in season_months.items():
        season_data = df[df.index.month.isin(months_in_season)]

        season_data = season_data.copy()
        season_data["Hour"] = season_data.index.hour

        result_df = pd.DataFrame()

        for orig_col, new_col in target_columns.items():
            if orig_col in season_data.columns:
                hourly_mean = season_data.groupby("Hour")[orig_col].mean()

                hourly_mean = hourly_mean.reindex(range(24), fill_value=None)

                result_df[new_col] = hourly_mean.values
            else:
                print(f"The {orig_col} column is not found.")
                result_df[new_col] = [None]*24

        result_df.index = list(range(24))
        result_df.index.name = "Hour"

        result_df.to_excel(writer, sheet_name=season_name)

print("✅113")







df = pd.read_excel("RES_P114.xlsx")

df.columns = [col.encode('latin1').decode('utf-8', errors='ignore') if isinstance(col, str) else col for col in df.columns]
df['Time'] = pd.to_datetime(df['Time'], errors='coerce')
df = df.set_index('Time')

target_columns = {
    'Generic 3 kW Power Output': 'خروجی توربین بادی',
    'Generic 1kWh Lead Acid Input Power': 'ورودی باتری',
    'Total Electrical Load Served': 'بار',
    'Grid Sales': 'فروش به شبکه',
    'Grid Purchases': 'خرید از شبکه',
}

season_months = {
    "Winter": [1, 2, 3],
    "Spring": [4, 5, 6],
    "Summer": [7, 8, 9],
    "Autumn": [10, 11, 12]
}

with pd.ExcelWriter("Result_P114.xlsx") as writer:
    for season_name, months_in_season in season_months.items():
        season_data = df[df.index.month.isin(months_in_season)]

        season_data = season_data.copy()
        season_data["Hour"] = season_data.index.hour

        result_df = pd.DataFrame()

        for orig_col, new_col in target_columns.items():
            if orig_col in season_data.columns:
                hourly_mean = season_data.groupby("Hour")[orig_col].mean()

                hourly_mean = hourly_mean.reindex(range(24), fill_value=None)

                result_df[new_col] = hourly_mean.values
            else:
                print(f"The {orig_col} column is not found.")
                result_df[new_col] = [None]*24

        result_df.index = list(range(24))
        result_df.index.name = "Hour"

        result_df.to_excel(writer, sheet_name=season_name)

print("✅114")










df = pd.read_excel("RES_P115.xlsx")

df.columns = [col.encode('latin1').decode('utf-8', errors='ignore') if isinstance(col, str) else col for col in df.columns]
df['Time'] = pd.to_datetime(df['Time'], errors='coerce')
df = df.set_index('Time')

target_columns = {
    'Generic flat plate PV Power Output': 'خروجی سلول خورشیدی',
    'Generic 3 kW Power Output': 'خروجی توربین بادی',
    'Generic 1kWh Lead Acid Input Power': 'ورودی باتری',
    'Total Electrical Load Served': 'بار',
    # 'Excess Electrical Production': 'فروش به شبکه',
    # 'Grid Purchases': 'خرید از شبکه',
}

season_months = {
    "Winter": [1, 2, 3],
    "Spring": [4, 5, 6],
    "Summer": [7, 8, 9],
    "Autumn": [10, 11, 12]
}

with pd.ExcelWriter("Result_P115.xlsx") as writer:
    for season_name, months_in_season in season_months.items():
        season_data = df[df.index.month.isin(months_in_season)]

        season_data = season_data.copy()
        season_data["Hour"] = season_data.index.hour

        result_df = pd.DataFrame()

        for orig_col, new_col in target_columns.items():
            if orig_col in season_data.columns:
                hourly_mean = season_data.groupby("Hour")[orig_col].mean()

                hourly_mean = hourly_mean.reindex(range(24), fill_value=None)

                result_df[new_col] = hourly_mean.values
            else:
                print(f"The {orig_col} column is not found.")
                result_df[new_col] = [None]*24

        result_df.index = list(range(24))
        result_df.index.name = "Hour"

        result_df.to_excel(writer, sheet_name=season_name)

print("✅115")










df = pd.read_excel("RES_P116.xlsx")

df.columns = [col.encode('latin1').decode('utf-8', errors='ignore') if isinstance(col, str) else col for col in df.columns]
df['Time'] = pd.to_datetime(df['Time'], errors='coerce')
df = df.set_index('Time')

target_columns = {
    'Generic flat plate PV Power Output': 'خروجی سلول خورشیدی',
    'Generic 3 kW Power Output': 'خروجی توربین بادی',
    'Generic 1kWh Lead Acid Input Power': 'ورودی باتری',
    'Total Electrical Load Served': 'بار',
    'Grid Sales': 'فروش به شبکه',
    'Grid Purchases': 'خرید از شبکه',
}

season_months = {
    "Winter": [1, 2, 3],
    "Spring": [4, 5, 6],
    "Summer": [7, 8, 9],
    "Autumn": [10, 11, 12]
}

with pd.ExcelWriter("Result_P116.xlsx") as writer:
    for season_name, months_in_season in season_months.items():
        season_data = df[df.index.month.isin(months_in_season)]

        season_data = season_data.copy()
        season_data["Hour"] = season_data.index.hour

        result_df = pd.DataFrame()

        for orig_col, new_col in target_columns.items():
            if orig_col in season_data.columns:
                hourly_mean = season_data.groupby("Hour")[orig_col].mean()

                hourly_mean = hourly_mean.reindex(range(24), fill_value=None)

                result_df[new_col] = hourly_mean.values
            else:
                print(f"The {orig_col} column is not found.")
                result_df[new_col] = [None]*24

        result_df.index = list(range(24))
        result_df.index.name = "Hour"

        result_df.to_excel(writer, sheet_name=season_name)

print("✅116")