import pandas as pd


# =========================================================
# FILE CONFIGURATION
# =========================================================

# Input dataset
arquivo = "dataset_treino_excel_complexo.xlsx"

# Output dataset
output_file = "dataset_tratado.xlsx"


# =========================================================
# LOAD ALL SHEETS FROM EXCEL
# =========================================================

# Read every sheet into a dictionary of DataFrames
dfs = pd.read_excel(arquivo, sheet_name=None)

# Dictionary that will store processed DataFrames
dfs_tratados = {}


# =========================================================
# PROCESS EACH SHEET
# =========================================================

for nome, df in dfs.items():

    # Process only sheets containing "Table"
    if "Table" in nome:

        # -------------------------------------------------
        # REMOVE INVALID OR INCOMPLETE ROWS
        # -------------------------------------------------

        # Remove rows missing critical business data
        df = df.dropna(subset=["Units", "UnitPrice", "Region", "Product"])

        # Remove duplicated orders
        df = df.drop_duplicates(subset=["OrderID"])


        # -------------------------------------------------
        # DEFINE COLUMN GROUPS
        # -------------------------------------------------

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

        colunas_float = ["UnitPrice", "Revenue"]
        colunas_int = ["Units"]
        colunas_string = ["Region", "Seller", "Product", "Category"]

        # -------------------------------------------------
        # CONVERT DATA TYPES
        # -------------------------------------------------

        for col_float in colunas_float:
            df[col_float] = df[col_float].astype(float)

        for col_int in colunas_int:
            df[col_int] = df[col_int].astype(int)

        for col_str in colunas_string:
            df[col_str] = df[col_str].astype(str)


        # -------------------------------------------------
        # TEXT STANDARDIZATION
        # -------------------------------------------------

        text_columns = ["Region", "Seller", "Product", "Category"]

        for col in text_columns:
            df[col] = (
                df[col]
                .str.title()      # Standard capitalization
                .str.strip()      # Remove leading/trailing spaces
                .str.replace(r"\s+", " ", regex=True)  # Normalize spacing
            )


        # -------------------------------------------------
        # APPROVAL STATUS MAPPING
        # -------------------------------------------------

        # Convert numeric status into readable labels
        df["Approved"] = df["Approved"].map({
            1: "Approved",
            0: "Not Approved"
        })


        # -------------------------------------------------
        # REVENUE CALCULATION
        # -------------------------------------------------

        # Recalculate revenue to guarantee data consistency
        df["Revenue"] = df["Units"] * df["UnitPrice"]


        # -------------------------------------------------
        # HANDLE MISSING STATUS VALUES
        # -------------------------------------------------

        df["Approved"] = df["Approved"].fillna("Not informed")


        # -------------------------------------------------
        # OPTIONAL DATA QUALITY CHECKS
        # -------------------------------------------------

        # Count missing values
        vazios = df[colunas].isnull().sum()

        # Count duplicated OrderID
        duplicados = df["OrderID"].duplicated().sum()

        # (Optional debugging prints)
        # print(nome)
        # print("Missing values:", vazios)
        # print("Duplicated OrderID:", duplicados)
        # print("===============")


        # -------------------------------------------------
        # STORE CLEANED DATAFRAME
        # -------------------------------------------------

        dfs_tratados[nome] = df



summary_t1 = (
    dfs_tratados["Table_1"].groupby("Region", as_index=False)["Revenue"].sum()
    .rename(columns={"Revenue" : "Total Revenue (Table_1)"})
)

summary_t2 = (
    dfs_tratados["Table_2"]
    .groupby("Region", as_index=False)["Revenue"].sum()
    .rename(columns={"Revenue" : "Total Revenue (Table_2)"})
)


summary_units = (
    dfs_tratados["Table_3"]
    .query("Units > 5")
    .groupby("Region", as_index=False)["Units"]
    .sum()
    .rename(columns={"Units": "Units > 5 (Table_3)"})
)

summary_df = summary_t1.merge(summary_t2, on="Region", how="outer")
summary_df = summary_df.merge(summary_units, on="Region", how="outer")

summary_df["Grand Total"] = (
    summary_df["Total Revenue (Table_1)"] +
    summary_df["Total Revenue (Table_2)"]

)

summary_df = summary_df[
    [
        "Region",
        "Total Revenue (Table_1)",
        "Total Revenue (Table_2)",
        "Grand Total",
        "Units > 5 (Table_3)"
    ]
]

dfs_tratados["Summary"] = summary_df



all_tables = pd.concat(
    [dfs_tratados[nome] for nome in dfs_tratados if "Table" in nome],
    ignore_index=True
)

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

dfs_tratados["Pivot_Simulated"] = pivot_simulated



# =========================================================
# EXPORT CLEAN DATA TO NEW EXCEL FILE
# =========================================================
colunas_formatar = ["Revenue", "Grand Total", "Total Revenue", "UnitPrice","Total Revenue (Table_1)", "Total Revenue (Table_2)"]

with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:

    workbook = writer.book
    formato_global = workbook.add_format({'num_format': '#,##0.00'})



    for nome, df in dfs_tratados.items():

        df.to_excel(writer, sheet_name=nome, index=False)

        worksheet = writer.sheets[nome]

        for idx, col in enumerate(df.columns):

            if col in colunas_formatar:
                worksheet.set_column(idx, idx, 18, formato_global)