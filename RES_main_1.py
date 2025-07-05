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

# Define a dictionary of specific months (by their number) and their names
# for which data will be extracted.
target_months = {
    2: "February",
    5: "May",
    8: "August",
    11: "November"
}

# Create an ExcelWriter object to write multiple DataFrames to different sheets
# within a single Excel file named "Result_P11.xlsx".
with pd.ExcelWriter("Result_P11.xlsx") as writer:
    # Iterate through each target month defined in the 'target_months' dictionary.
    for month_num, month_name in target_months.items():
        # Initialize an empty DataFrame to store the results for the current month.
        result_df = pd.DataFrame()

        # Iterate through the target columns defined earlier.
        for orig_col, new_col in target_columns.items():
            # Check if the original column exists in the DataFrame.
            if orig_col in df.columns:
                # Filter the DataFrame to get data for the 15th day of the current month
                # for the specific original column.
                filtered = df[(df.index.month == month_num) & (df.index.day == 15)][orig_col]

                # Check if there are at least 24 values (for 24 hours) for the filtered data.
                if len(filtered) >= 24:
                    # If enough data is available, take the first 24 values and assign them
                    # to the result_df under the new column name.
                    result_df[new_col] = filtered.iloc[:24].values
                else:
                    # If fewer than 24 values are available, print a warning message
                    # and fill the result_df column with None for all 24 hours.
                    print(f"Just {len(filtered)} values are available for {new_col} in {month_name}")
                    result_df[new_col] = [None] * 24
            else:
                # If a target column is not found in the data, print a message
                # and add a column of None values for that column in the result_df.
                print(f"The {orig_col} column is not found.")
                result_df[new_col] = [None] * 24

        # Set the index of the result_df to a list of integers from 0 to 23, representing hours.
        result_df.index = list(range(24))
        
        # Name the index column 'Hour'.
        result_df.index.name = "Hour"

        # Write the result_df for the current month to a new sheet in the Excel file,
        # with the sheet name being the month's name.
        result_df.to_excel(writer, sheet_name=month_name)

# Print a success message after processing the first Excel file.
print("✅11")









# Second Part (Similar Functionality)

# This sections of the code performs an identical set of operations as the first section,
# but it processes a different input file ("RES_P12.xlsx") and includes additional
# target columns for analysis.
df = pd.read_excel("RES_P12.xlsx")

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
target_months = {
    2: "February",
    5: "May",
    8: "August",
    11: "November"
}

with pd.ExcelWriter("Result_P12.xlsx") as writer:
    for month_num, month_name in target_months.items():
        result_df = pd.DataFrame()

        for orig_col, new_col in target_columns.items():
            if orig_col in df.columns:
                filtered = df[(df.index.month == month_num) & (df.index.day == 15)][orig_col]

                if len(filtered) >= 24:
                    result_df[new_col] = filtered.iloc[:24].values
                else:
                    print(f"Just {len(filtered)} values are available for {new_col} in {month_name}")
                    result_df[new_col] = [None] * 24
            else:
                print(f"The {orig_col} column is not found.")
                result_df[new_col] = [None] * 24

        result_df.index = list(range(24))
        result_df.index.name = "Hour"

        result_df.to_excel(writer, sheet_name=month_name)

print("✅12")









df = pd.read_excel("RES_P13.xlsx")

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
target_months = {
    2: "February",
    5: "May",
    8: "August",
    11: "November"
}

# ایجاد فایل اکسل با چند شیت
with pd.ExcelWriter("Result_P13.xlsx") as writer:
    for month_num, month_name in target_months.items():
        result_df = pd.DataFrame()

        for orig_col, new_col in target_columns.items():
            if orig_col in df.columns:
                filtered = df[(df.index.month == month_num) & (df.index.day == 15)][orig_col]

                if len(filtered) >= 24:
                    result_df[new_col] = filtered.iloc[:24].values
                else:
                    print(f"Just {len(filtered)} values are available for {new_col} in {month_name}")
                    result_df[new_col] = [None] * 24
            else:
                print(f"The {orig_col} column is not found.")
                result_df[new_col] = [None] * 24

        result_df.index = list(range(24))
        result_df.index.name = "Hour"

        result_df.to_excel(writer, sheet_name=month_name)

print("✅13")







df = pd.read_excel("RES_P14.xlsx")

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
target_months = {
    2: "February",
    5: "May",
    8: "August",
    11: "November"
}

# ایجاد فایل اکسل با چند شیت
with pd.ExcelWriter("Result_P14.xlsx") as writer:
    for month_num, month_name in target_months.items():
        result_df = pd.DataFrame()

        for orig_col, new_col in target_columns.items():
            if orig_col in df.columns:
                filtered = df[(df.index.month == month_num) & (df.index.day == 15)][orig_col]

                if len(filtered) >= 24:
                    result_df[new_col] = filtered.iloc[:24].values
                else:
                    print(f"Just {len(filtered)} values are available for {new_col} in {month_name}")
                    result_df[new_col] = [None] * 24
            else:
                print(f"The {orig_col} column is not found.")
                result_df[new_col] = [None] * 24

        result_df.index = list(range(24))
        result_df.index.name = "Hour"

        result_df.to_excel(writer, sheet_name=month_name)

print("✅14")










df = pd.read_excel("RES_P15.xlsx")

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
target_months = {
    2: "February",
    5: "May",
    8: "August",
    11: "November"
}

# ایجاد فایل اکسل با چند شیت
with pd.ExcelWriter("Result_P15.xlsx") as writer:
    for month_num, month_name in target_months.items():
        result_df = pd.DataFrame()

        for orig_col, new_col in target_columns.items():
            if orig_col in df.columns:
                filtered = df[(df.index.month == month_num) & (df.index.day == 15)][orig_col]

                if len(filtered) >= 24:
                    result_df[new_col] = filtered.iloc[:24].values
                else:
                    print(f"Just {len(filtered)} values are available for {new_col} in {month_name}")
                    result_df[new_col] = [None] * 24
            else:
                print(f"The {orig_col} column is not found.")
                result_df[new_col] = [None] * 24

        result_df.index = list(range(24))
        result_df.index.name = "Hour"

        result_df.to_excel(writer, sheet_name=month_name)

print("✅15")










df = pd.read_excel("RES_P16.xlsx")

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
target_months = {
    2: "February",
    5: "May",
    8: "August",
    11: "November"
}

# ایجاد فایل اکسل با چند شیت
with pd.ExcelWriter("Result_P16.xlsx") as writer:
    for month_num, month_name in target_months.items():
        result_df = pd.DataFrame()

        for orig_col, new_col in target_columns.items():
            if orig_col in df.columns:
                filtered = df[(df.index.month == month_num) & (df.index.day == 15)][orig_col]

                if len(filtered) >= 24:
                    result_df[new_col] = filtered.iloc[:24].values
                else:
                    print(f"Just {len(filtered)} values are available for {new_col} in {month_name}")
                    result_df[new_col] = [None] * 24
            else:
                print(f"The {orig_col} column is not found.")
                result_df[new_col] = [None] * 24

        result_df.index = list(range(24))
        result_df.index.name = "Hour"

        result_df.to_excel(writer, sheet_name=month_name)

print("✅16")