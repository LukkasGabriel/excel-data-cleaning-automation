import pandas as pd


# =========================================================
# FILE CONFIGURATION
# =========================================================

# Path of the input Excel dataset (raw data)
arquivo = "dataset_treino_excel_complexo.xlsx"

# Name of the output Excel file after cleaning and processing
output_file = "dataset_tratado.xlsx"


# =========================================================
# LOAD ALL SHEETS FROM EXCEL
# =========================================================

# Read all sheets from the Excel file
# sheet_name=None loads every sheet into a dictionary of DataFrames
dfs = pd.read_excel(arquivo, sheet_name=None)

# Dictionary that will store all cleaned DataFrames
dfs_tratados = {}


# =========================================================
# PROCESS EACH SHEET
# =========================================================

# Loop through every sheet in the Excel file
for nome, df in dfs.items():

    # Process only sheets that contain "Table" in their name
    # This avoids processing sheets like summaries or metadata
    if "Table" in nome:

        # -------------------------------------------------
        # REMOVE INVALID OR INCOMPLETE ROWS
        # -------------------------------------------------

        # Remove rows that are missing critical business fields
        # These fields are required to calculate revenue and analyze sales
        df = df.dropna(subset=["Units", "UnitPrice", "Region", "Product"])

        # Remove duplicated orders based on OrderID
        # This prevents double counting of sales
        df = df.drop_duplicates(subset=["OrderID"])


        # -------------------------------------------------
        # DEFINE COLUMN GROUPS
        # -------------------------------------------------

        # Full list of expected columns
        colunas = [
            "OrderID",
            "Date",
            "Region",
            "Seller",
            "Product",
            "Category",
            "Units",
            "UnitPrice",
            "Revenue",
            "Approved"
        ]

        # Columns that must be converted to float
        colunas_float = ["UnitPrice", "Revenue"]

        # Columns that must be converted to integer
        colunas_int = ["Units"]

        # Columns that should be treated as strings
        colunas_string = ["Region", "Seller", "Product", "Category"]


        # -------------------------------------------------
        # CONVERT DATA TYPES
        # -------------------------------------------------

        # Ensure numeric columns are properly formatted
        for col_float in colunas_float:
            df[col_float] = df[col_float].astype(float)

        for col_int in colunas_int:
            df[col_int] = df[col_int].astype(int)

        # Convert text columns to string type
        for col_str in colunas_string:
            df[col_str] = df[col_str].astype(str)


        # -------------------------------------------------
        # TEXT STANDARDIZATION
        # -------------------------------------------------

        # Columns where text formatting must be standardized
        text_columns = ["Region", "Seller", "Product", "Category"]

        for col in text_columns:
            df[col] = (
                df[col]
                .str.title()      # Capitalize first letter of each word
                .str.strip()      # Remove leading and trailing spaces
                .str.replace(r"\s+", " ", regex=True)  # Normalize multiple spaces
            )


        # -------------------------------------------------
        # APPROVAL STATUS MAPPING
        # -------------------------------------------------

        # Convert numeric approval codes into readable labels
        df["Approved"] = df["Approved"].map({
            1: "Approved",
            0: "Not Approved"
        })


        # -------------------------------------------------
        # REVENUE CALCULATION
        # -------------------------------------------------

        # Recalculate revenue to guarantee data consistency
        # Revenue should always be Units * UnitPrice
        df["Revenue"] = df["Units"] * df["UnitPrice"]


        # -------------------------------------------------
        # HANDLE MISSING STATUS VALUES
        # -------------------------------------------------

        # Replace missing approval values with a default label
        df["Approved"] = df["Approved"].fillna("Not informed")


        # -------------------------------------------------
        # OPTIONAL DATA QUALITY CHECKS
        # -------------------------------------------------

        # Count missing values in important columns
        vazios = df[colunas].isnull().sum()

        # Count duplicated OrderIDs
        duplicados = df["OrderID"].duplicated().sum()

        # Optional debugging prints
        # Uncomment if you want to inspect the dataset quality
        # print(nome)
        # print("Missing values:", vazios)
        # print("Duplicated OrderID:", duplicados)
        # print("===============")


        # -------------------------------------------------
        # STORE CLEANED DATAFRAME
        # -------------------------------------------------

        # Save the cleaned DataFrame in the dictionary
        dfs_tratados[nome] = df



# =========================================================
# CREATE REVENUE SUMMARIES
# =========================================================

# Revenue summary for Table_1 grouped by Region
summary_t1 = (
    dfs_tratados["Table_1"]
    .groupby("Region", as_index=False)["Revenue"]
    .sum()
    .rename(columns={"Revenue" : "Total Revenue (Table_1)"})
)

# Revenue summary for Table_2 grouped by Region
summary_t2 = (
    dfs_tratados["Table_2"]
    .groupby("Region", as_index=False)["Revenue"]
    .sum()
    .rename(columns={"Revenue" : "Total Revenue (Table_2)"})
)

# Units summary from Table_3 considering only rows with Units > 5
summary_units = (
    dfs_tratados["Table_3"]
    .query("Units > 5")
    .groupby("Region", as_index=False)["Units"]
    .sum()
    .rename(columns={"Units": "Units > 5 (Table_3)"})
)

# Merge summaries together by Region
summary_df = summary_t1.merge(summary_t2, on="Region", how="outer")
summary_df = summary_df.merge(summary_units, on="Region", how="outer")

# Calculate the grand total revenue
summary_df["Grand Total"] = (
    summary_df["Total Revenue (Table_1)"] +
    summary_df["Total Revenue (Table_2)"]
)

# Reorder columns for better readability
summary_df = summary_df[
    [
        "Region",
        "Total Revenue (Table_1)",
        "Total Revenue (Table_2)",
        "Grand Total",
        "Units > 5 (Table_3)"
    ]
]

# Store summary sheet
dfs_tratados["Summary"] = summary_df



# =========================================================
# CREATE PIVOT-STYLE REPORT
# =========================================================

# Combine all table datasets into a single DataFrame
all_tables = pd.concat(
    [dfs_tratados[nome] for nome in dfs_tratados if "Table" in nome],
    ignore_index=True
)

# Simulate a pivot table using groupby
pivot_simulated = (
    all_tables
    .groupby(["Region", "Product"], as_index=False)
    .agg({
        "Units": "sum",
        "Revenue": "sum"
    })
    .rename(columns={
        "Units": "Total Units",
        "Revenue": "Total Revenue"
    })
)

# Store pivot-style report
dfs_tratados["Pivot_Simulated"] = pivot_simulated



# =========================================================
# EXPORT CLEAN DATA TO NEW EXCEL FILE
# =========================================================

# Columns that should receive currency formatting
colunas_formatar = [
    "Revenue",
    "Grand Total",
    "Total Revenue",
    "UnitPrice",
    "Total Revenue (Table_1)",
    "Total Revenue (Table_2)"
]

# Create Excel writer with formatting support
with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:

    workbook = writer.book

    # Define numeric format for financial values
    formato_global = workbook.add_format({'num_format': '#,##0.00'})

    # Export each cleaned DataFrame as a separate sheet
    for nome, df in dfs_tratados.items():

        df.to_excel(writer, sheet_name=nome, index=False)

        worksheet = writer.sheets[nome]

        # Apply formatting to financial columns
        for idx, col in enumerate(df.columns):

            if col in colunas_formatar:
                worksheet.set_column(idx, idx, 18, formato_global)
