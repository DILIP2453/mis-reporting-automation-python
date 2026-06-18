import pandas as pd
from sqlalchemy import create_engine

# 🔴 SQL connection
engine = create_engine(
    "mssql+pyodbc://localhost\\SQLEXPRESS/inventory_analyst?driver=ODBC+Driver+17+for+SQL+Server"
)

# 🔴 Excel file path
file_path = r"C:\Users\ADMIN\Downloads\inventory_dataset.xlsx"

# 🔥 Clean column names
def clean_columns(df):
    df.columns = df.columns.str.strip().str.lower().str.replace(" ", "_")
    return df

# 🔥 Smart date fix (AUTO DETECT)
def fix_dates(df):
    for col in df.columns:
        if 'date' in col:
            if pd.api.types.is_numeric_dtype(df[col]):
                # Excel serial → real date
                df[col] = pd.to_datetime(df[col], origin='1899-12-30', unit='D', errors='coerce')
            else:
                # already date/string → safe convert
                df[col] = pd.to_datetime(df[col], errors='coerce')
    return df

# =========================
# 🔥 PRODUCTS
# =========================
df_products = pd.read_excel(file_path, sheet_name='products')
df_products = clean_columns(df_products)

df_products = df_products[['sku','product_name','category','cost_price','selling_price']]

df_products.to_sql('raw_products', engine, if_exists='replace', index=False)
print("✅ Products Loaded")

# =========================
# 🔥 PURCHASES
# =========================
df_purchases = pd.read_excel(file_path, sheet_name='purchase')
df_purchases = clean_columns(df_purchases)

df_purchases = fix_dates(df_purchases)

df_purchases = df_purchases[['purchase_id','sku','supplier','warehouse','qty','cost','purchase_date','delivery_date']]

df_purchases.to_sql('raw_purchases', engine, if_exists='replace', index=False)
print("✅ Purchases Loaded")

# =========================
# 🔥 SALES
# =========================
df_sales = pd.read_excel(file_path, sheet_name='sales')
df_sales = clean_columns(df_sales)

df_sales = fix_dates(df_sales)

df_sales = df_sales[['sale_id','sku','warehouse','store','qty','selling_price','sale_date']]

df_sales.to_sql('raw_sales', engine, if_exists='replace', index=False)
print("✅ Sales Loaded")

# =========================
# 🔥 INVENTORY BATCHES
# =========================
df_inventory = pd.read_excel(file_path, sheet_name='inventory_batches')
df_inventory = clean_columns(df_inventory)

df_inventory = fix_dates(df_inventory)

df_inventory = df_inventory[['batch_id','sku','warehouse','qty','unit_cost','purchase_date','expiry_date']]

df_inventory.to_sql('raw_inventory_batches', engine, if_exists='replace', index=False)
print("✅ Inventory Loaded")

print("\n🔥 All Sheets Imported Successfully with Correct Dates!")