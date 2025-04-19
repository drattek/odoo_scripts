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

        "id": row["id_external"],  # id_external OTRO CAMPO EN BD
        "description": row["Description"],  # Description 
        "name": row["name"],  # Product Name
        "default_code": row["default_code"],  # Internal Reference / SKU
        "x_studio_xcross": row["xcross"],  # xcross OTRO CAMPO EN BD
        "standard_price": row["standard_price"],  # Cost
        "lst_price": row["lst_price"],  # Sales price
        # "uom_id": row["uom_id"],  # ERROR por tipo de campo many2one
        "barcode": row["barcode"],  # barcode
        "type": "consu",  # Product type: 'product', 'consu', or 'service'
        # "taxes_id": row["taxes_id"],  # ERROR por tipo de campo many2many
        # "supplier_taxes_id": row["supplier_taxes_id"],  # ERROR por tipo de campo many2many
        "is_storable": row["is_storable"],  # is_storable
        "invoice_policy": row["invoice_policy"],  # invoice_policy
        # "categ_id": row["categ_id"],  # ERROR por tipo de campo many2one
        # "pos_categ_id": row["pos_categ_id"],  # ERROR por tipo de campo many2many
        "sale_ok": row["sale_ok"],  # sale_ok
        "purchase_ok": row["purcahse_ok"],  # purcahse_ok
        "available_in_pos": row["available_in_pos"],  # available_in_pos
        "is_published": row["is_published"],  # is_published
        "self_order_available": row["self_order_available"],  # self_order_available
        # "route_ids/id": row["`route_ids/id`"],  # ERROR por tipo de campo many2many
        # "categoria": row["categoria"],  # ERROR No hay como tal un campo de categoria
        # "CVE_LINEA": row["CVE_LINEA"],  # ERROR No hay como tal un campo en la BD
        # "cve_linea2": row["cve_linea2"],  # ERROR No hay como tal un campo en la BD
        # "clave_SAT": row["clave_SAT"],  # no se si sea hs_code en BD, que tiene como etiqueta Código SA
        # "UBICACION": row["UBICACION"],  # ERROR de las 3 opciones de ubicacion las 3 son many2one

    }

    try:
        product_id = models.execute_kw(
            db, uid, password,
            "product.product", "create",
            [product_data]
        )
        print(f"Created product with ID {product_id}: {product_data['name']}")
    except Exception as e:
        print(f"Failed to create product {product_data['name']}: {e}\n")

