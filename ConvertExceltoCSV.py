import pandas as pd
from sqlalchemy import create_engine

# File path
input_file = "FILENAME.xlsx"
output_file = "FILENAME.csv"

# Read the Excel file (all as string to preserve Arabic text and dates)
df = pd.read_excel(input_file, dtype=str)

# Save as CSV with UTF-8 encoding
df.to_csv(output_file, index=False, encoding="utf-8-sig")
print("CSV file created at:", output_file)

# ------------------------------
# Upload same data into SQL Server
# ------------------------------

# SQL Server connection details (update these)
server = "YOUR_SERVER_NAME"
database = "YOUR_DATABASE_NAME"
username = "YOUR_USERNAME"
password = "YOUR_PASSWORD"

# Create SQLAlchemy engine
engine = create_engine(
    f"mssql+pyodbc://{username}:{password}@{server}/{database}?driver=ODBC+Driver+17+for+SQL+Server"
)

# Target table name
table_name = "TABLENAME"

# Upload data (append rows to table)
df.to_sql(table_name, con=engine, if_exists="append", index=False)

print("Data uploaded successfully to SQL Server table:", table_name)



# import pandas as pd

# # File path
# input_file = "filename.Xlsx"
# output_file = "filename.csv"

# # Read the Excel file
# df = pd.read_excel(input_file, dtype=str)  # Read all as string to preserve formatting like Arabic text and dates

# # Save as CSV with UTF-8 encoding
# df.to_csv(output_file, index=False, encoding="utf-8-sig")

# output_file