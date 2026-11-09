import pandas as pd


# STEP 1 — BANKING SYSTEM SIZE


# CBB Excel file
file_path = "Jun-2026.xlsx"

# Load the Excel workbook
excel_file = pd.ExcelFile(file_path)

# Check the sheets available in the workbook
print("Sheets in the CBB workbook:")
print(excel_file.sheet_names)



# TABLE 13 — TOTAL BANKING SYSTEM ASSETS


table_13 = pd.read_excel(
    file_path,
    sheet_name="13",
    header=None
)

# Table 13:
# Column 0 = End of Period / Year
# Column 8 = Total Assets
#
# Rows 12:22 = Annual observations 2016–2025

annual_assets = table_13.iloc[12:22, [0, 8]].copy()

annual_assets.columns = [
    "Year",
    "Total_Assets_USD_Million"
]

print("\nTable 13 — Total Banking System Assets")
print(annual_assets)\


# YEAR-OVER-YEAR ASSET GROWTH


annual_assets["Asset_Growth_YoY_%"] = (
    annual_assets["Total_Assets_USD_Million"].pct_change() * 100
)

print("\nBanking System Asset Growth (YoY)")
print(annual_assets)


# TABLE 14 — RETAIL BANK ASSETS


table_14 = pd.read_excel(
    file_path,
    sheet_name="14",
    header=None
)

# Table 14:
# Column 0 = Year
# Column 11 = Total Assets
#
# Rows 12:22 = Annual observations 2016–2025

retail_assets = table_14.iloc[12:22, [0, 11]].copy()

retail_assets.columns = [
    "Year",
    "Retail_Bank_Assets_BD_Million"
]

print("\nTable 14 — Retail Bank Assets")
print(retail_assets)


# RETAIL BANK ASSET SHARE


asset_data = annual_assets.merge(
    retail_assets,
    on="Year",
    how="inner"
)

asset_data["Retail_Asset_Share_%"] = (
    asset_data["Retail_Bank_Assets_BD_Million"]
    / (asset_data["Total_Assets_USD_Million"] * 0.376)
    * 100
)

print("\nRetail Banks' Share of Total Banking System Assets")
print(asset_data)

# ==========================================
# STEP 2 — DEPOSITS
# ==========================================

# TABLE 19 — RETAIL BANK DEPOSIT LIABILITIES


table_19 = pd.read_excel(
    file_path,
    sheet_name="19",
    header=None
)

annual_deposits = table_19.iloc[
    12:22,
    [0, 2, 4, 6, 8, 10, 12, 14, 15, 16]
].copy()

annual_deposits.columns = [
    "Year",
    "Government_Deposits_BD_Million",
    "Demand_Deposits_BD_Million",
    "Savings_Deposits_BD_Million",
    "Time_Deposits_BD_Million",
    "Domestic_Deposits_BD_Million",
    "Foreign_Deposits_BD_Million",
    "Total_Deposits_BD_Million",
    "Total_Deposits_FC_Million",
    "Total_Deposits_All_Currencies_BD_Million"
]

print("\nTable 19 — Annual Deposit Data")
print(annual_deposits)


# PRIVATE-SECTOR BD DEPOSITS


annual_deposits["Private_Sector_BD_Deposits"] = (
    annual_deposits["Demand_Deposits_BD_Million"]
    + annual_deposits["Savings_Deposits_BD_Million"]
    + annual_deposits["Time_Deposits_BD_Million"]
)

# Reconstruct domestic BD deposits
annual_deposits["Calculated_Domestic_Deposits_BD_Million"] = (
    annual_deposits["Government_Deposits_BD_Million"]
    + annual_deposits["Private_Sector_BD_Deposits"]
)

# Compare our calculated domestic deposits with the CBB domestic total
annual_deposits["Domestic_Deposit_Check_Difference"] = (
    annual_deposits["Calculated_Domestic_Deposits_BD_Million"]
    - annual_deposits["Domestic_Deposits_BD_Million"]
)

print("\nDeposit Reconciliation Check")
print(
    annual_deposits[
        [
            "Year",
            "Government_Deposits_BD_Million",
            "Private_Sector_BD_Deposits",
            "Domestic_Deposits_BD_Million",
            "Calculated_Domestic_Deposits_BD_Million",
            "Domestic_Deposit_Check_Difference"
        ]
    ]
)

# TOTAL DEPOSIT GROWTH — YEAR OVER YEAR


annual_deposits["Deposit_Growth_YoY_%"] = (
    annual_deposits["Total_Deposits_All_Currencies_BD_Million"]
    .pct_change()
    * 100
)

print("\nTotal Deposit Growth (YoY)")
print(
    annual_deposits[
        [
            "Year",
            "Total_Deposits_All_Currencies_BD_Million",
            "Deposit_Growth_YoY_%"
        ]
    ]
)

# PRIVATE-SECTOR DEPOSIT COMPOSITION


annual_deposits["Demand_Share_%"] = (
    annual_deposits["Demand_Deposits_BD_Million"]
    / annual_deposits["Private_Sector_BD_Deposits"]
    * 100
)

annual_deposits["Savings_Share_%"] = (
    annual_deposits["Savings_Deposits_BD_Million"]
    / annual_deposits["Private_Sector_BD_Deposits"]
    * 100
)

annual_deposits["Time_Share_%"] = (
    annual_deposits["Time_Deposits_BD_Million"]
    / annual_deposits["Private_Sector_BD_Deposits"]
    * 100
)

# Check that the three categories add up to 100%
annual_deposits["Composition_Check_%"] = (
    annual_deposits["Demand_Share_%"]
    + annual_deposits["Savings_Share_%"]
    + annual_deposits["Time_Share_%"]
)

print("\nPrivate-Sector Deposit Composition")
print(
    annual_deposits[
        [
            "Year",
            "Private_Sector_BD_Deposits",
            "Demand_Share_%",
            "Savings_Share_%",
            "Time_Share_%",
            "Composition_Check_%"
        ]
    ]
)

# GOVERNMENT VS PRIVATE-SECTOR BD DEPOSITS


annual_deposits["Government_Share_%"] = (
    annual_deposits["Government_Deposits_BD_Million"]
    / annual_deposits["Domestic_Deposits_BD_Million"]
    * 100
)

annual_deposits["Private_Sector_Share_%"] = (
    annual_deposits["Private_Sector_BD_Deposits"]
    / annual_deposits["Domestic_Deposits_BD_Million"]
    * 100
)

# Check that the two shares add up to 100%
annual_deposits["Depositor_Share_Check_%"] = (
    annual_deposits["Government_Share_%"]
    + annual_deposits["Private_Sector_Share_%"]
)

print("\nGovernment vs Private-Sector Deposit Share")
print(
    annual_deposits[
        [
            "Year",
            "Government_Deposits_BD_Million",
            "Private_Sector_BD_Deposits",
            "Government_Share_%",
            "Private_Sector_Share_%",
            "Depositor_Share_Check_%"
        ]
    ]
)

# ==========================================
# STEP 3 — LENDING
# ==========================================

# TABLE 20 — OUTSTANDING LOANS AND ADVANCES


table_20 = pd.read_excel(
    file_path,
    sheet_name="20",
    header=None
)

annual_lending = table_20.iloc[
    12:22,
    [0, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11]
].copy()

annual_lending.columns = [
    "Year",
    "Business_Loans_BD_Million",
    "Government_Loans_BD_Million",
    "Mortgages_BD_Million",
    "Vehicle_Loans_BD_Million",
    "Loans_Against_Deposits_BD_Million",
    "Salary_Assignment_Loans_BD_Million",
    "Credit_Card_Loans_BD_Million",
    "Other_Personal_Loans_BD_Million",
    "Personal_Loans_BD_Million",
    "Total_Lending_BD_Million"
]

print("\nTable 20 — Annual Lending Data")
print(annual_lending)

# LENDING RECONCILIATION CHECK


# Check 1:
# Business + Government + Personal = Total Lending

annual_lending["Calculated_Total_Lending_BD_Million"] = (
    annual_lending["Business_Loans_BD_Million"]
    + annual_lending["Government_Loans_BD_Million"]
    + annual_lending["Personal_Loans_BD_Million"]
)

annual_lending["Total_Lending_Check_Difference"] = (
    annual_lending["Calculated_Total_Lending_BD_Million"]
    - annual_lending["Total_Lending_BD_Million"]
)


# Check 2:
# All personal-loan categories = Personal Sector Total

annual_lending["Calculated_Personal_Lending_BD_Million"] = (
    annual_lending["Mortgages_BD_Million"]
    + annual_lending["Vehicle_Loans_BD_Million"]
    + annual_lending["Loans_Against_Deposits_BD_Million"]
    + annual_lending["Salary_Assignment_Loans_BD_Million"]
    + annual_lending["Credit_Card_Loans_BD_Million"]
    + annual_lending["Other_Personal_Loans_BD_Million"]
)

annual_lending["Personal_Lending_Check_Difference"] = (
    annual_lending["Calculated_Personal_Lending_BD_Million"]
    - annual_lending["Personal_Loans_BD_Million"]
)

print("\nLending Reconciliation Check")

print(
    annual_lending[
        [
            "Year",
            "Total_Lending_BD_Million",
            "Calculated_Total_Lending_BD_Million",
            "Total_Lending_Check_Difference",
            "Personal_Loans_BD_Million",
            "Calculated_Personal_Lending_BD_Million",
            "Personal_Lending_Check_Difference"
        ]
    ]
)

# TOTAL LENDING GROWTH — YEAR OVER YEAR


annual_lending["Lending_Growth_YoY_%"] = (
    annual_lending["Total_Lending_BD_Million"].pct_change()
    * 100
)

print("\nTotal Lending Growth (YoY)")
print(
    annual_lending[
        [
            "Year",
            "Total_Lending_BD_Million",
            "Lending_Growth_YoY_%"
        ]
    ]
)

# LENDING COMPOSITION


annual_lending["Business_Share_%"] = (
    annual_lending["Business_Loans_BD_Million"]
    / annual_lending["Total_Lending_BD_Million"]
    * 100
)

annual_lending["Government_Share_%"] = (
    annual_lending["Government_Loans_BD_Million"]
    / annual_lending["Total_Lending_BD_Million"]
    * 100
)

annual_lending["Personal_Share_%"] = (
    annual_lending["Personal_Loans_BD_Million"]
    / annual_lending["Total_Lending_BD_Million"]
    * 100
)

# Check that the three categories add up to 100%
annual_lending["Lending_Composition_Check_%"] = (
    annual_lending["Business_Share_%"]
    + annual_lending["Government_Share_%"]
    + annual_lending["Personal_Share_%"]
)

print("\nLending Composition")
print(
    annual_lending[
        [
            "Year",
            "Business_Share_%",
            "Government_Share_%",
            "Personal_Share_%",
            "Lending_Composition_Check_%"
        ]
    ]
)

# PERSONAL LENDING COMPOSITION


personal_categories = [
    "Mortgages_BD_Million",
    "Vehicle_Loans_BD_Million",
    "Loans_Against_Deposits_BD_Million",
    "Salary_Assignment_Loans_BD_Million",
    "Credit_Card_Loans_BD_Million",
    "Other_Personal_Loans_BD_Million"
]

for column in personal_categories:
    share_column = column.replace(
        "_BD_Million",
        "_Share_%"
    )

    annual_lending[share_column] = (
        annual_lending[column]
        / annual_lending["Personal_Loans_BD_Million"]
        * 100
    )

# Check that all personal lending categories add up to 100%

annual_lending["Personal_Lending_Composition_Check_%"] = (
    annual_lending["Mortgages_Share_%"]
    + annual_lending["Vehicle_Loans_Share_%"]
    + annual_lending["Loans_Against_Deposits_Share_%"]
    + annual_lending["Salary_Assignment_Loans_Share_%"]
    + annual_lending["Credit_Card_Loans_Share_%"]
    + annual_lending["Other_Personal_Loans_Share_%"]
)

print("\nPersonal Lending Composition")

print(
    annual_lending[
        [
            "Year",
            "Mortgages_Share_%",
            "Vehicle_Loans_Share_%",
            "Loans_Against_Deposits_Share_%",
            "Salary_Assignment_Loans_Share_%",
            "Credit_Card_Loans_Share_%",
            "Other_Personal_Loans_Share_%",
            "Personal_Lending_Composition_Check_%"
        ]
    ]
)

# ==========================================
# STEP 3D — LENDING RATES
# ==========================================

table_7 = pd.read_excel(
    file_path,
    sheet_name="7",
    header=None
)

# Table 7:
# Column 0  = Year
# Column 10 = Business Loans:
#              Total (Excludes overdraft approvals)
# Column 17 = Personal Loans:
#              Total
# Column 18 = Credit Cards
#
# Rows 12:22 = Annual observations 2016–2025

annual_lending_rates = table_7.iloc[
    12:22,
    [0, 10, 17, 18]
].copy()

annual_lending_rates.columns = [
    "Year",
    "Business_Loan_Rate_Excl_Overdraft_%",
    "Personal_Loan_Rate_%",
    "Credit_Card_Rate_%"
]

print("\nTable 7 — Annual Lending Rates")
print(annual_lending_rates)

# ==========================================
# STEP 4 — PROFITABILITY
# ==========================================


# TABLE 37 — ROA & ROE


table_37 = pd.read_excel(
    file_path,
    sheet_name="37",
    header=None
)

annual_profitability = table_37.iloc[
    14:24,
    [0, 6, 7]
].copy()

annual_profitability.columns = [
    "Year",
    "ROA_%",
    "ROE_%"
]

print("\nTable 37 — Profitability")
print(annual_profitability)


# STEP 4B — ASSET QUALITY


annual_asset_quality = table_37.iloc[
    14:24,
    [0, 4, 5]
].copy()

annual_asset_quality.columns = [
    "Year",
    "NPL_Ratio_%",
    "Specific_Provisions_%"
]

print("\nTable 37 — Asset Quality")
print(annual_asset_quality)


# STEP 4C — LIQUIDITY

annual_liquidity = table_37.iloc[
    14:24,
    [0, 8, 9]
].copy()

annual_liquidity.columns = [
    "Year",
    "Liquid_Assets_Ratio_%",
    "Loans_Deposit_Ratio_%"
]

print("\nTable 37 — Liquidity")
print(annual_liquidity)


# STEP 4D — CAPITAL ADEQUACY


annual_capital = table_37.iloc[
    14:24,
    [0, 2, 3]
].copy()

annual_capital.columns = [
    "Year",
    "Total_Capital_Adequacy_Ratio_%",
    "Tier1_Capital_Adequacy_Ratio_%"
]

print("\nTable 37 — Capital Adequacy")
print(annual_capital)

# ==========================================
# STEP 5 — BANKING STRUCTURE
# ==========================================

# TABLE 13 — ASSET & LIABILITY STRUCTURE


annual_structure = table_13.iloc[
    12:22,
    [0, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14]
].copy()

annual_structure.columns = [
    "Year",

    "Assets_Banks_USD_Million",
    "Assets_Private_NonBanks_USD_Million",
    "Assets_Government_USD_Million",
    "Assets_Other_USD_Million",
    "Assets_Domestic_Total_USD_Million",
    "Assets_Foreign_USD_Million",
    "Total_Assets_USD_Million",

    "Liabilities_Banks_USD_Million",
    "Liabilities_Private_NonBanks_USD_Million",
    "Liabilities_Government_USD_Million",
    "Liabilities_Other_USD_Million",
    "Liabilities_Domestic_Total_USD_Million",
    "Liabilities_Foreign_USD_Million"
]

print("\nTable 13 — Banking Structure")
print(annual_structure)


# ASSET STRUCTURE RECONCILIATION CHECK


annual_structure["Calculated_Total_Assets_USD_Million"] = (
    annual_structure["Assets_Banks_USD_Million"]
    + annual_structure["Assets_Private_NonBanks_USD_Million"]
    + annual_structure["Assets_Government_USD_Million"]
    + annual_structure["Assets_Other_USD_Million"]
    + annual_structure["Assets_Foreign_USD_Million"]
)

annual_structure["Asset_Check_Difference"] = (
    annual_structure["Calculated_Total_Assets_USD_Million"]
    - annual_structure["Total_Assets_USD_Million"]
)

print("\nAsset Structure Reconciliation Check")
print(
    annual_structure[
        [
            "Year",
            "Total_Assets_USD_Million",
            "Calculated_Total_Assets_USD_Million",
            "Asset_Check_Difference"
        ]
    ]
)


# ASSET COMPOSITION


annual_structure["Banks_Asset_Share_%"] = (
    annual_structure["Assets_Banks_USD_Million"]
    / annual_structure["Total_Assets_USD_Million"]
    * 100
)

annual_structure["Private_NonBanks_Asset_Share_%"] = (
    annual_structure["Assets_Private_NonBanks_USD_Million"]
    / annual_structure["Total_Assets_USD_Million"]
    * 100
)

annual_structure["Government_Asset_Share_%"] = (
    annual_structure["Assets_Government_USD_Million"]
    / annual_structure["Total_Assets_USD_Million"]
    * 100
)

annual_structure["Other_Asset_Share_%"] = (
    annual_structure["Assets_Other_USD_Million"]
    / annual_structure["Total_Assets_USD_Million"]
    * 100
)

annual_structure["Foreign_Asset_Share_%"] = (
    annual_structure["Assets_Foreign_USD_Million"]
    / annual_structure["Total_Assets_USD_Million"]
    * 100
)

# Check that the five categories add up to approximately 100%

annual_structure["Asset_Composition_Check_%"] = (
    annual_structure["Banks_Asset_Share_%"]
    + annual_structure["Private_NonBanks_Asset_Share_%"]
    + annual_structure["Government_Asset_Share_%"]
    + annual_structure["Other_Asset_Share_%"]
    + annual_structure["Foreign_Asset_Share_%"]
)

print("\nAsset Composition")
print(
    annual_structure[
        [
            "Year",
            "Banks_Asset_Share_%",
            "Private_NonBanks_Asset_Share_%",
            "Government_Asset_Share_%",
            "Other_Asset_Share_%",
            "Foreign_Asset_Share_%",
            "Asset_Composition_Check_%"
        ]
    ]
)


# LIABILITY STRUCTURE RECONCILIATION CHECK


annual_structure["Calculated_Total_Liabilities_USD_Million"] = (
    annual_structure["Liabilities_Banks_USD_Million"]
    + annual_structure["Liabilities_Private_NonBanks_USD_Million"]
    + annual_structure["Liabilities_Government_USD_Million"]
    + annual_structure["Liabilities_Other_USD_Million"]
    + annual_structure["Liabilities_Foreign_USD_Million"]
)

annual_structure["Liability_Check_Difference"] = (
    annual_structure["Calculated_Total_Liabilities_USD_Million"]
    - (
        annual_structure["Liabilities_Domestic_Total_USD_Million"]
        + annual_structure["Liabilities_Foreign_USD_Million"]
    )
)

print("\nLiability Structure Reconciliation Check")
print(
    annual_structure[
        [
            "Year",
            "Liabilities_Domestic_Total_USD_Million",
            "Liabilities_Foreign_USD_Million",
            "Calculated_Total_Liabilities_USD_Million",
            "Liability_Check_Difference"
        ]
    ]
)


# LIABILITY COMPOSITION


annual_structure["Banks_Liability_Share_%"] = (
    annual_structure["Liabilities_Banks_USD_Million"]
    / annual_structure["Calculated_Total_Liabilities_USD_Million"]
    * 100
)

annual_structure["Private_NonBanks_Liability_Share_%"] = (
    annual_structure["Liabilities_Private_NonBanks_USD_Million"]
    / annual_structure["Calculated_Total_Liabilities_USD_Million"]
    * 100
)

annual_structure["Government_Liability_Share_%"] = (
    annual_structure["Liabilities_Government_USD_Million"]
    / annual_structure["Calculated_Total_Liabilities_USD_Million"]
    * 100
)

annual_structure["Other_Liability_Share_%"] = (
    annual_structure["Liabilities_Other_USD_Million"]
    / annual_structure["Calculated_Total_Liabilities_USD_Million"]
    * 100
)

annual_structure["Foreign_Liability_Share_%"] = (
    annual_structure["Liabilities_Foreign_USD_Million"]
    / annual_structure["Calculated_Total_Liabilities_USD_Million"]
    * 100
)

# Check that all categories add up to approximately 100%

annual_structure["Liability_Composition_Check_%"] = (
    annual_structure["Banks_Liability_Share_%"]
    + annual_structure["Private_NonBanks_Liability_Share_%"]
    + annual_structure["Government_Liability_Share_%"]
    + annual_structure["Other_Liability_Share_%"]
    + annual_structure["Foreign_Liability_Share_%"]
)

print("\nLiability Composition")
print(
    annual_structure[
        [
            "Year",
            "Banks_Liability_Share_%",
            "Private_NonBanks_Liability_Share_%",
            "Government_Liability_Share_%",
            "Other_Liability_Share_%",
            "Foreign_Liability_Share_%",
            "Liability_Composition_Check_%"
        ]
    ]
)

# ==========================================
# STEP 6 — CONVENTIONAL VS ISLAMIC BANKING
# ==========================================


# TABLE 38 — CONVENTIONAL BANKS


table_38 = pd.read_excel(
    file_path,
    sheet_name="38",
    header=None
)

annual_conventional = table_38.iloc[
    14:24,
    [0, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17]
].copy()

annual_conventional.columns = [
    "Year",

    "Total_CAR_Retail_%",
    "Total_CAR_Wholesale_%",

    "Tier1_CAR_Retail_%",
    "Tier1_CAR_Wholesale_%",

    "NPL_Retail_%",
    "NPL_Wholesale_%",

    "Specific_Provisions_Retail_%",
    "Specific_Provisions_Wholesale_%",

    "ROA_Retail_%",
    "ROA_Wholesale_%",

    "ROE_Retail_%",
    "ROE_Wholesale_%",

    "Liquid_Assets_Retail_%",
    "Liquid_Assets_Wholesale_%",

    "Loans_Deposit_Retail_%",
    "Loans_Deposit_Wholesale_%"
]

print("\nTable 38 — Conventional Banks")
print(annual_conventional)


# TABLE 39 — ISLAMIC BANKS


table_39 = pd.read_excel(
    file_path,
    sheet_name="39",
    header=None
)

annual_islamic = table_39.iloc[
    14:24,
    [0, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17]
].copy()

annual_islamic.columns = [
    "Year",

    "Total_CAR_Retail_%",
    "Total_CAR_Wholesale_%",

    "Tier1_CAR_Retail_%",
    "Tier1_CAR_Wholesale_%",

    "NPL_Retail_%",
    "NPL_Wholesale_%",

    "Specific_Provisions_Retail_%",
    "Specific_Provisions_Wholesale_%",

    "ROA_Retail_%",
    "ROA_Wholesale_%",

    "ROE_Retail_%",
    "ROE_Wholesale_%",

    "Liquid_Assets_Retail_%",
    "Liquid_Assets_Wholesale_%",

    "Loans_Deposit_Retail_%",
    "Loans_Deposit_Wholesale_%"
]
pd.set_option("display.max_columns", None)
pd.set_option("display.width", 200)
print("\nTable 39 — Islamic Banks")
print(annual_islamic)


# STEP 6B — CAPITAL ADEQUACY COMPARISON


capital_comparison = annual_conventional[
    [
        "Year",
        "Total_CAR_Retail_%",
        "Total_CAR_Wholesale_%"
    ]
].merge(
    annual_islamic[
        [
            "Year",
            "Total_CAR_Retail_%",
            "Total_CAR_Wholesale_%"
        ]
    ],
    on="Year",
    suffixes=("_Conventional", "_Islamic")
)

# Calculate the difference in percentage points
capital_comparison["Retail_CAR_Difference_pp"] = (
    capital_comparison["Total_CAR_Retail_%_Conventional"]
    - capital_comparison["Total_CAR_Retail_%_Islamic"]
)

capital_comparison["Wholesale_CAR_Difference_pp"] = (
    capital_comparison["Total_CAR_Wholesale_%_Conventional"]
    - capital_comparison["Total_CAR_Wholesale_%_Islamic"]
)

print("\nCapital Adequacy — Conventional vs Islamic")
print(capital_comparison)


# STEP 6C — TIER 1 CAPITAL COMPARISON

tier1_comparison = annual_conventional[
    [
        "Year",
        "Tier1_CAR_Retail_%",
        "Tier1_CAR_Wholesale_%"
    ]
].merge(
    annual_islamic[
        [
            "Year",
            "Tier1_CAR_Retail_%",
            "Tier1_CAR_Wholesale_%"
        ]
    ],
    on="Year",
    suffixes=("_Conventional", "_Islamic")
)

# Difference in percentage points

tier1_comparison["Retail_Tier1_Difference_pp"] = (
    tier1_comparison["Tier1_CAR_Retail_%_Conventional"]
    - tier1_comparison["Tier1_CAR_Retail_%_Islamic"]
)

tier1_comparison["Wholesale_Tier1_Difference_pp"] = (
    tier1_comparison["Tier1_CAR_Wholesale_%_Conventional"]
    - tier1_comparison["Tier1_CAR_Wholesale_%_Islamic"]
)

print("\nTier 1 Capital Adequacy — Conventional vs Islamic")
print(tier1_comparison)


# STEP 6D — NPL COMPARISON

npl_comparison = annual_conventional[
    [
        "Year",
        "NPL_Retail_%",
        "NPL_Wholesale_%"
    ]
].merge(
    annual_islamic[
        [
            "Year",
            "NPL_Retail_%",
            "NPL_Wholesale_%"
        ]
    ],
    on="Year",
    suffixes=("_Conventional", "_Islamic")
)

# Difference in percentage points

npl_comparison["Retail_NPL_Difference_pp"] = (
    npl_comparison["NPL_Retail_%_Conventional"]
    - npl_comparison["NPL_Retail_%_Islamic"]
)

npl_comparison["Wholesale_NPL_Difference_pp"] = (
    npl_comparison["NPL_Wholesale_%_Conventional"]
    - npl_comparison["NPL_Wholesale_%_Islamic"]
)

print("\nNPL Ratio — Conventional vs Islamic")
print(npl_comparison)

# STEP 6E — SPECIFIC PROVISIONS COMPARISON


provisions_comparison = annual_conventional[
    [
        "Year",
        "Specific_Provisions_Retail_%",
        "Specific_Provisions_Wholesale_%"
    ]
].merge(
    annual_islamic[
        [
            "Year",
            "Specific_Provisions_Retail_%",
            "Specific_Provisions_Wholesale_%"
        ]
    ],
    on="Year",
    suffixes=("_Conventional", "_Islamic")
)

# Difference in percentage points

provisions_comparison["Retail_Provisions_Difference_pp"] = (
    provisions_comparison["Specific_Provisions_Retail_%_Conventional"]
    - provisions_comparison["Specific_Provisions_Retail_%_Islamic"]
)

provisions_comparison["Wholesale_Provisions_Difference_pp"] = (
    provisions_comparison["Specific_Provisions_Wholesale_%_Conventional"]
    - provisions_comparison["Specific_Provisions_Wholesale_%_Islamic"]
)

print("\nSpecific Provisions — Conventional vs Islamic")
print(provisions_comparison)


# STEP 6F — ROA COMPARISON


roa_comparison = annual_conventional[
    [
        "Year",
        "ROA_Retail_%",
        "ROA_Wholesale_%"
    ]
].merge(
    annual_islamic[
        [
            "Year",
            "ROA_Retail_%",
            "ROA_Wholesale_%"
        ]
    ],
    on="Year",
    suffixes=("_Conventional", "_Islamic")
)

# Difference in percentage points

roa_comparison["Retail_ROA_Difference_pp"] = (
    roa_comparison["ROA_Retail_%_Conventional"]
    - roa_comparison["ROA_Retail_%_Islamic"]
)

roa_comparison["Wholesale_ROA_Difference_pp"] = (
    roa_comparison["ROA_Wholesale_%_Conventional"]
    - roa_comparison["ROA_Wholesale_%_Islamic"]
)

print("\nROA — Conventional vs Islamic")
print(roa_comparison)


# STEP 6G — ROE COMPARISON


roe_comparison = annual_conventional[
    [
        "Year",
        "ROE_Retail_%",
        "ROE_Wholesale_%"
    ]
].merge(
    annual_islamic[
        [
            "Year",
            "ROE_Retail_%",
            "ROE_Wholesale_%"
        ]
    ],
    on="Year",
    suffixes=("_Conventional", "_Islamic")
)

# Difference in percentage points

roe_comparison["Retail_ROE_Difference_pp"] = (
    roe_comparison["ROE_Retail_%_Conventional"]
    - roe_comparison["ROE_Retail_%_Islamic"]
)

roe_comparison["Wholesale_ROE_Difference_pp"] = (
    roe_comparison["ROE_Wholesale_%_Conventional"]
    - roe_comparison["ROE_Wholesale_%_Islamic"]
)

print("\nROE — Conventional vs Islamic")
print(roe_comparison)


# STEP 6H — LIQUID ASSETS RATIO COMPARISON


liquid_assets_comparison = annual_conventional[
    [
        "Year",
        "Liquid_Assets_Retail_%",
        "Liquid_Assets_Wholesale_%"
    ]
].merge(
    annual_islamic[
        [
            "Year",
            "Liquid_Assets_Retail_%",
            "Liquid_Assets_Wholesale_%"
        ]
    ],
    on="Year",
    suffixes=("_Conventional", "_Islamic")
)

# Difference in percentage points

liquid_assets_comparison["Retail_Liquid_Assets_Difference_pp"] = (
    liquid_assets_comparison["Liquid_Assets_Retail_%_Conventional"]
    - liquid_assets_comparison["Liquid_Assets_Retail_%_Islamic"]
)

liquid_assets_comparison["Wholesale_Liquid_Assets_Difference_pp"] = (
    liquid_assets_comparison["Liquid_Assets_Wholesale_%_Conventional"]
    - liquid_assets_comparison["Liquid_Assets_Wholesale_%_Islamic"]
)

print("\nLiquid Assets Ratio — Conventional vs Islamic")
print(liquid_assets_comparison)

# STEP 6I — LOANS / DEPOSIT RATIO COMPARISON

loans_deposit_comparison = annual_conventional[
    [
        "Year",
        "Loans_Deposit_Retail_%",
        "Loans_Deposit_Wholesale_%"
    ]
].merge(
    annual_islamic[
        [
            "Year",
            "Loans_Deposit_Retail_%",
            "Loans_Deposit_Wholesale_%"
        ]
    ],
    on="Year",
    suffixes=("_Conventional", "_Islamic")
)

# Difference in percentage points

loans_deposit_comparison["Retail_Loans_Deposit_Difference_pp"] = (
    loans_deposit_comparison["Loans_Deposit_Retail_%_Conventional"]
    - loans_deposit_comparison["Loans_Deposit_Retail_%_Islamic"]
)

loans_deposit_comparison["Wholesale_Loans_Deposit_Difference_pp"] = (
    loans_deposit_comparison["Loans_Deposit_Wholesale_%_Conventional"]
    - loans_deposit_comparison["Loans_Deposit_Wholesale_%_Islamic"]
)

print("\nLoans / Deposit Ratio — Conventional vs Islamic")
print(loans_deposit_comparison)