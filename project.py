# 1.Importing Libraries
import pandas as pd
import numpy as np

# 2.Reading the Excel File
df = pd.read_excel('Flight.xlsx')

# 3.Setting Column Names and Dropping Rows
df.columns = df.iloc[1]
df = df.drop([0, 1])
df.reset_index(drop=True, inplace=True)

# 4.Displaying Information about the DataFrame
df.info()

# 5.Displaying Head and Tail of the DataFrame
print('Head of the DataFrame: ')
print(df.head())
print('\nTail of the DataFrame: ')
print(df.tail())

# 6.Counting Separators in the DataFrame
separator_count_comma = df.map(lambda x: isinstance(x, str) and ',' in x).sum().sum()
separator_count_semicolon = df.map(lambda x: isinstance(x, str) and ';' in x).sum().sum()
print(f"Number of ',' separators: {separator_count_comma}")
print(f"Number of ';' separators: {separator_count_semicolon}")

# 7.Checking if a Header Exists
header_exists = pd.read_excel("Flight.xlsx", header=None).map(lambda x: isinstance(x, str)).sum().sum()
print(f"Header exists: {header_exists > 0}")

# 8.Counting NAs in the DataFrame
na_count = df.isna().sum().sum()
print(f"Number of NAs in the dataset: {na_count}")

# 9.Checking for Duplicate Columns
duplicate_columns = df.columns[df.columns.duplicated()]
if duplicate_columns.empty:
    print("There are no duplicate column names.")
else:
    print(f"Duplicate column names: {duplicate_columns}")

# 10.Displaying Unique Values in Each Column
for column in df.columns:
    unique_values = df[column].unique()
    print(f"\nUnique values in column '{column}':\n{unique_values}")

# 11.Checking for Variables Stored in Rows
variables_in_rows = df.map(lambda x: isinstance(x, str) and x.lower() == 'na').sum().sum()
if variables_in_rows > 0:
    print(f"There are {variables_in_rows} instances where variables are stored in rows.")
else:
    print("There are no variables stored in rows.")

# 12.Checking for Duplicate Rows
duplicate_rows = df[df.duplicated()]
if not duplicate_rows.empty:
    print("Duplicate rows found. Review the following:")
    print(duplicate_rows)
else:
    print("There are no duplicate rows.")


# 13.Creating a Tidy DataFrame
tidy_df = df.drop_duplicates()
tidy_df.columns = tidy_df.columns.str.strip().str.upper()
tidy_df.columns = tidy_df.columns.str.replace('"', '').str.replace('_', ' ')
tidy_df.columns = tidy_df.columns.str.replace('#', '')
print(tidy_df.head())

# 14.Cleaning Object Columns in the Tidy DataFrame
object_columns = tidy_df.select_dtypes(include=['object']).columns
for column in object_columns:
    tidy_df.loc[:, column] = tidy_df[column].astype(str).str.replace('"', '').str.replace('_', ' ').str.upper()
print(tidy_df.head())
object_columns = tidy_df.select_dtypes(include=['object']).columns

# 15.Replacing Incorrect Values in Object Columns
for column in object_columns:
    if column in tidy_df.columns:
        print(f"Column: {column}")
        print(tidy_df[column].value_counts())
        print("\n")
        tidy_df.loc[:, column] = tidy_df[column].replace({'incorrect_value': 'correct_value'})
print(tidy_df.head())

# 16.Further Cleaning Specific Columns
tidy_df.loc[:, 'CABIN TYPE'] = tidy_df['CABIN TYPE'].replace({'S': 'STANDARD', 'E': 'ECONOMY', 'L': 'LUXURY'})
tidy_df.loc[:, 'DISTANCE KM'] = pd.to_numeric(tidy_df['DISTANCE KM'], errors='coerce')
tidy_df.loc[:, 'DISTANCE KM'] = tidy_df['DISTANCE KM'].apply(lambda x: abs(x) if not pd.isna(x) else x)

# 17.Handling Numeric Columns
numeric_columns = ['FLIGHT ID', 'DISTANCE KM', 'DURATION HOURS', 'PRICE', 'DELAY MINUTES', 'THE NUMBER OF PASSENGERS', 'BAGGAGE FEE']
tidy_df[numeric_columns] = tidy_df[numeric_columns].replace(',', '.', regex=True)
tidy_df[numeric_columns] = tidy_df[numeric_columns].astype(float)

# 18.Handling Categorical Columns:
categorical_columns = ['CABIN TYPE', 'DAY OF WEEK', 'AIRLINE', 'SEAT PLACE', 'DESTINATION AIRPORT', 'SOURCE AIRPORT']
tidy_df[categorical_columns] = tidy_df[categorical_columns].astype(object)

# 19.Converting 'DATE' Column to Datetime
tidy_df['DATE'] = pd.to_datetime(tidy_df['DATE'])
print(tidy_df.dtypes)

# 20.Handling Missing Values in Categorical Columns
tidy_df['CABIN TYPE'] = tidy_df['CABIN TYPE'].replace({'NAN': np.nan})
categorical_columns = tidy_df.select_dtypes(include=['object']).columns
for col in categorical_columns:
    mode_value = tidy_df[col].mode()[0]
    tidy_df[col].fillna(mode_value, inplace=True)
cabin_type_mode = tidy_df['CABIN TYPE'].mode()[0]
tidy_df['CABIN TYPE'] = tidy_df['CABIN TYPE'].fillna(cabin_type_mode)
tidy_df.replace('NAN', np.nan, inplace=True)
day_of_week_mode = tidy_df['DAY OF WEEK'].mode()[0]
tidy_df['DAY OF WEEK'].fillna(day_of_week_mode, inplace=True)

# 21.Handling Missing Values in Numeric Columns:
numeric_columns = tidy_df.select_dtypes(include=['number']).columns
for col in numeric_columns:
    mean_value = round(tidy_df[col].mean())
    tidy_df[col].fillna(mean_value, inplace=True)

# 22.Formatting 'DATE' Column and Fixing 'CABIN TYPE'
tidy_df['DATE'] = tidy_df['DATE'].dt.date
tidy_df['CABIN TYPE'] = tidy_df['CABIN TYPE'].str.strip()
unique_cabin_types = tidy_df['CABIN TYPE'].unique()
if 'STANDARD' in unique_cabin_types and 'standard' in unique_cabin_types:
    tidy_df['CABIN TYPE'] = tidy_df['CABIN TYPE'].replace({'standard': 'STANDARD'})
unique_cabin_types_after_fix = tidy_df['CABIN TYPE'].unique()
print(f"Unique values in 'CABIN TYPE' after fixing: {unique_cabin_types_after_fix}")

# 23.Saving the Cleaned DataFrame to a New Excel File
output_file_path = 'cleaned.xlsx'
tidy_df.to_excel(output_file_path, index=False)
print(f'The DataFrame has been saved to {output_file_path}')