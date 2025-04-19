import xmlrpc.client
import pandas as pd

# Odoo configurations
url = "https://erp.refaceballos.com"  # Replace with your Odoo URL
db = "erp.refaceballos.com"  # Replace with your database name
username = "admin"  # Replace your Odoo user's email
password = ".Reface123&abc"  # Replace with your Odoo user's password

# Establish connection with Odoo's API
common = xmlrpc.client.ServerProxy(f"{url}/xmlrpc/2/common")
uid = common.authenticate(db, username, password, {})

if not uid:
    print("Authentication failed. Check your credentials.")
    exit()

models = xmlrpc.client.ServerProxy(f"{url}/xmlrpc/2/object")

# Path to the Excel file
file_path = "products.xlsx"

# Read the Excel file into a pandas DataFrame
try:
    # Use pandas to read the .xlsx file
    df = pd.read_excel(file_path)

    # Print the first 5 rows of the DataFrame
    print("Excel file data (first 5 rows):")
    print(df.head())

except FileNotFoundError:
    print(f"The file '{file_path}' does not exist. Please check the path.")
except Exception as e:
    print(f"An error occurred: {e}")

df = pd.read_excel(file_path, sheet_name="Sheet1")  # Replace "Sheet1" with your actual sheet name

for index, row in df.iterrows():

    product_data = {

        "name": row["name"],  # Product Name
        "lst_price": row["lst_price"],  # Sales price
        "standard_price": row["standard_price"],  # Cost
        "type": "consu",  # Product type: 'product', 'consu', or 'service'
        #"uom_id": row["uom_id"],  # Unit of Measure ID
        "default_code": row["default_code"],  # Internal Reference / SKU
    }

    try:
        product_id = models.execute_kw(
            db, uid, password,
            "product.product", "create",
            [product_data]
        )
        print(f"Created product with ID {product_id}: {product_data['name']}")
    except Exception as e:
        print(f"Failed to create product {product_data['name']}: {e}")

