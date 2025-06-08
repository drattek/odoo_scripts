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

    columns = df.columns  # Define 'columns' as the list of column names in the DataFrame
    for field in columns:
            # Extract the value from the row for the current column
            field_value = row.get(field, None)

            # Handle cases where the column value is missing or empty
            if not field_value:
                print(f"WARNING: The value for column '{field}' is empty or missing. Skipping.")
                data[field] = None
                continue
            # Specify the model name you want to inspect
            model_name = "product.product"
            data = {}
            # Get the fields (columns) of the model
            model_fields = models.execute_kw(db, uid, password, model_name, 'fields_get', [],
                                        {})
            #validates field on Excel file is on model
            field_metadata = model_fields  # Initialize field_metadata with model_fields retrieved earlier

            if field in field_metadata:
                # Get the metadata for the specific field
                metadata = field_metadata[field]

                # Access specific information from the metadata
                field_type = metadata.get('type')  # Retrieve the field type
                field_label = metadata.get('string')  # Retrieve the label of the field
                related_model = metadata.get('relation')  # Retrieve the related model, if any

                print(f"Field Name: {field}")
                print(f"Type: {field_type}")
                print(f"Label: {field_label}")
                print(f"Related Model: {related_model}")

                if field_type == 'many2one' and related_model:
                    # Search for the field by name
                    #model_fields = models.execute_kw(db, uid, password, related_model, 'fields_get', [],{})
                    field_ids = models.execute_kw(
                        db, uid, password,
                        related_model, 'search',
                        [[('name', '=', field_value)]],  # Search criteria
                        {'limit': 1}  # Retrieve only one record
                    )

                    # If the catalog exists, return its ID
                    if field_ids:
                        data[field] = field_ids[0]
                    else:
                        # Otherwise, create the category and return its ID
                        field_id = models.execute_kw(
                            db, uid, password,
                            related_model, 'create',
                            [{'name': field_value}]  # Data for the new category
                        )
                        data[field] = field_id

                elif field_type == 'many2many' and related_model:
                    print(f"The '{field}' field is a many2many relationship with the model: {related_model}")
                    record = models.execute_kw(
                        db, uid, password,
                        related_model,  # The model to query
                        "search_read",  # The Odoo method to fetch records
                        [[[field, "=", field_value]], ["id"]]  # Search domain and fields to fetch
                    )

                    # Map the resolved ID to the column name or set as None if not resolved
                    if record:
                        data[field] = record[0]["id"]  # Found: Assign the ID
                    else:
                        print(f"WARNING: '{field_value}' not found for column '{field}'.")
                        data[field] = None  # Not found: Set as None

                else:
                    data[field] = field_value

    #try:
    #    product_id = models.execute_kw(
    #        db, uid, password,
    #        "product.product", "create",
    #        [data]
    #    )
    #    print(f"Created product with ID {product_id}: {data['name']}")
    #except Exception as e:
    #    print(f"Failed to create product {data['name']}: {e}\n")
