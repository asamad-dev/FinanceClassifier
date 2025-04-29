import pandas as pd

# Set pandas display options to show full width
pd.set_option('display.max_columns', None)  # Show all columns
pd.set_option('display.width', 0)            # No line width restriction
pd.set_option('display.max_colwidth', None)  # Show full content inside each column

# Load file
file_path = "input.xlsx"
df = pd.read_excel(file_path, sheet_name="Sheet1", header=0)

# Fix unnamed columns
new_columns = []
for idx, col in enumerate(df.columns):
    if "Unnamed" in str(col) or pd.isna(col):
        new_columns.append(f"Column_{idx+1}")
    else:
        new_columns.append(col)

df.columns = new_columns

# Display
print(df.head(50))

# Rows and columns count
print(f"Number of rows: {df.shape[0]}")
print(f"Number of columns: {df.shape[1]}")