import xmlrpc.client
import pandas as pd

# Odoo configurations
url = "https://pruebas.refaceballos.com"  # Replace with your Odoo URL
db = "reface_tst"  # Replace with your database name
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
file_path = "products_small_tst.xlsx"

# Read the Excel file into a pandas DataFrame
try:
    # Use pandas to read the .xlsx file
    df = pd.read_excel(file_path, sheet_name="Sheet1")  # Default to "Sheet1", change if necessary

    # Print the first 5 rows of the DataFrame
    print("Excel file data (first 5 rows):")
    print(df.head())

except FileNotFoundError:
    print(f"The file '{file_path}' does not exist. Please check the path.")
    exit()
except Exception as e:
    print(f"An error occurred: {e}")
    exit()

df = pd.read_excel(file_path, sheet_name="Sheet1")  # Replace "Sheet1" with your actual sheet name

for index, row in df.iterrows():
    columns = df.columns  # Define 'columns' as the list of column names in the DataFrame
    data = {}
    data_supplierinfo = {}

    for field in columns:
        field_value = row.get(field, None)

        if pd.isna(field_value) or not field_value:
            print(f"WARNING: The value for column '{field}' is empty or missing. Skipping.")
            continue

        try:
            # Get the fields (columns) of the model
            model_name = "product.supplierinfo"
            model_fields = models.execute_kw(db, uid, password, model_name, 'fields_get', [])

            if field in model_fields:
                field_metadata = model_fields[field]
                field_type = field_metadata.get('type')
                related_model = field_metadata.get('relation')

                if field_type == 'many2one' and related_model:
                    print(f'Handling many2one relationship for field \'{field}\' with value \'{field_value}\'.')

                    field_ids = models.execute_kw(
                        db, uid, password,
                        related_model, 'search',
                        [[('name', '=', field_value)]],
                        {'limit': 1}
                    )

                    if field_ids:
                        print(f'Field already exists: \'{field}\' with value \'{field_value}\'..\'')
                        data_supplierinfo[field] = field_ids[0]
                    else:
                        # Create the related record if not found
                        print(f'Field not found, creating record: \'{field}\' with value \'{field_value}\'..\'')
                        field_id = models.execute_kw(
                            db, uid, password,
                            related_model, 'create',
                            [{'name': field_value}]
                        )
                        data_supplierinfo[field] = field_id
                elif field_type == 'many2many' and related_model:
                    # Many-to-many resolution
                    print(f"Handling many2many relationship for field '{field}' with value '{field_value}'..'")

                    # Normalize field_value into a list of values
                    if isinstance(field_value, str):
                        # Split comma-separated string into a list
                        values_to_search = [v.strip() for v in field_value.split(',')]
                    elif isinstance(field_value, list):
                        # Already a list
                        values_to_search = field_value
                    else:
                        # Handle unexpected type: Log warning, skip, or raise an error
                        print(f"Unexpected type for field_value: {type(field_value)}")
                        values_to_search = []
                    # Iterate over each value and perform the search
                    field_ids = []
                    for value in values_to_search:
                        print(f"Searching for value '{value}' in related model '{related_model}'")
                        ids = models.execute_kw(
                            db, uid, password,
                            related_model, 'search',
                            [[('name', '=', value)]],
                            {'context': {'lang': 'es_MX'}}
                        )
                        field_ids.extend(ids)  # Collect all IDs

                    print(f"Resolved field_ids for field '{field}': {field_ids}")

                    if field_ids:
                        #if not isinstance(field_ids[0], list):
                        #    field_ids[0] = [field_ids[0]] if field_ids[0] is not None else []
                        data_supplierinfo[field] = field_ids
                    else:

                        # Create the related record if not found generic
                        field_id = models.execute_kw(
                            db, uid, password,
                            related_model, 'create',
                            [{'name': field_value}]
                        )

                        # Ensure field_id is always a list for many2many relationships
                        field_id = [field_id] if not isinstance(field_id, list) else field_id
                        data_supplierinfo[field] = field_id
                else:
                    # Default value assignment
                    data_supplierinfo[field] = field_value


        except Exception as e:
            print(f"Failed to process field '{field}': {e}")

        try:
            # Extract field metadata from the model
            model_fields = models.execute_kw(db, uid, password, "product.product", 'fields_get', [])
            if field in model_fields:
                metadata = model_fields[field]
                field_type = metadata.get('type')
                related_model = metadata.get('relation')

                if field_type == 'many2one' and related_model:




                    print(f'Handling many2one relationship for field \'{field}\' with value \'{field_value}\'..\'')
                    if field == 'categ_id':
                        parent_categ_id = row.get('parent_categ_id', None)
                        parent_categ_id_value = None

                        categ_id = row.get('categ_id', None)
                        categ_id_value = None

                        if parent_categ_id:
                            # Search for the parent category by name
                            parent_categ_ids = models.execute_kw(
                                db, uid, password,
                                related_model, 'search',
                                [[('name', '=', parent_categ_id)]],
                                {'limit': 1}
                            )
                            if parent_categ_ids:
                                parent_categ_id_value = parent_categ_ids[0]
                            else:
                                # Create the parent category if not found
                                parent_categ_id_value = models.execute_kw(
                                    db, uid, password,
                                    related_model, 'create',
                                    [{'name': parent_categ_id}]
                                )
                        if categ_id:
                            # Search for the parent category by name
                            categ_ids = models.execute_kw(
                                db, uid, password,
                                related_model, 'search',
                                [[('name', '=', categ_id)]],
                                {'limit': 1}
                            )
                            if categ_ids:
                                categ_id_value = categ_ids[0]
                                # Update the existing categ_id, if applicable
                                models.execute_kw(
                                    db, uid, password,
                                    related_model, 'write',
                                    [[categ_id_value], {'name': categ_id, 'parent_id': parent_categ_id_value}]
                                )
                                data[field] = categ_id_value
                            else:
                                # Create the category with parent_id set if applicable
                                categ_id_value = models.execute_kw(
                                    db, uid, password,
                                    related_model, 'create',
                                    [{'name': categ_id, 'parent_id': parent_categ_id_value}]
                                )
                                data[field] = categ_id_value
                    elif field == 'unspsc_code_id':
                        # Handle specific search condition for unspsc_code_id
                        search_field = 'code' if field == 'unspsc_code_id' else 'name'
                        field_ids = models.execute_kw(
                            db, uid, password,
                            related_model, 'search',
                            [[(search_field, '=', field_value)]],  # Search criteria
                            {'limit': 1,
                             'context': {'lang': 'es_MX'}  # Language specified here (if needed)
                             }
                        )
                    elif field.startswith('x_'):
                        # Handle specific search condition for fields starting with 'x'
                        # Search for the field by name

                        field_ids = models.execute_kw(
                            db, uid, password,
                            related_model, 'search',
                            [[('x_name', '=', field_value)]],
                            {'limit': 1}
                        )

                        if field_ids:
                            print(f'Field already exist.. \'{field}\' with value \'{field_value}\'..\'')
                            data[field] = field_ids[0]
                        else:
                            # Create the related record if not found
                            print(f'Field not found, creating record.. \'{field}\' with value \'{field_value}\'..\'')
                            field_id = models.execute_kw(
                                db, uid, password,
                                related_model, 'create',
                                [{'x_name': field_value}]
                            )
                            data[field] = field_id

                    else:
                        # Search for the field by name

                        field_ids = models.execute_kw(
                            db, uid, password,
                            related_model, 'search',
                            [[('x_name', '=', field_value)]],
                            {'limit': 1}
                        )

                        if field_ids:
                            print(f'Field already exist.. \'{field}\' with value \'{field_value}\'..\'')
                            data[field] = field_ids[0]
                        else:
                            # Create the related record if not found
                            print(f'Field not found, creating record.. \'{field}\' with value \'{field_value}\'..\'')
                            field_id = models.execute_kw(
                                db, uid, password,
                                related_model, 'create',
                                [{'x_name': field_value}]
                            )
                            data[field] = field_id

                elif field_type == 'many2many' and related_model:
                    # Many-to-many resolution
                    print(f"Handling many2many relationship for field '{field}' with value '{field_value}'..'")


                    # Normalize field_value into a list of values
                    if isinstance(field_value, str):
                        # Split comma-separated string into a list
                        values_to_search = [v.strip() for v in field_value.split(',')]
                    elif isinstance(field_value, list):
                        # Already a list
                        values_to_search = field_value
                    else:
                        # Handle unexpected type: Log warning, skip, or raise an error
                        print(f"Unexpected type for field_value: {type(field_value)}")
                        values_to_search = []
                        # Iterate over each value and perform the search
                        field_ids = []
                        for value in values_to_search:
                            print(f"Searching for value '{value}' in related model '{related_model}'")
                            ids = models.execute_kw(
                                db, uid, password,
                                related_model, 'search',
                                [[['name', '=', value]]],  # Apply search criterion for the current value
                                {
                                    'context': {'lang': 'es_MX'}  # Language specified here (if needed)
                                }
                            )
                            field_ids.extend(ids)  # Collect all IDs

                        print(f"Resolved field_ids for field '{field}': {field_ids}")

                        if field_ids:
                            #if not isinstance(field_ids[0], list):
                            #    field_ids[0] = [field_ids[0]] if field_ids[0] is not None else []
                            data[field] = field_ids
                        else:

                            # Create the related record if not found generic
                            field_id = models.execute_kw(
                                db, uid, password,
                                related_model, 'create',
                                [{'name': field_value}]
                            )

                            # Ensure field_id is always a list
                            if not isinstance(field_id, list):
                                field_id = [field_id] if field_id is not None else []
                            data[field] = field_id
                else:
                    data[field] = field_value


        except Exception as e:
            print(f"Failed to process field '{field}': {e}")

    # Create the product
    print(f"Data for creation: {data}")
    try:
        # Check if the product already exists by name or relevant criteria
        existing_product_ids = models.execute_kw(
            db, uid, password,
            "product.product", "search",
            [[("name", "=", data.get("name"))]],  # Adjust the search field if necessary
            {"limit": 1}
        )

        if existing_product_ids:
            product_id = existing_product_ids[0]
            print(f"Product already exists with ID {product_id}: {data} .. updating data instead")
            models.execute_kw(
                db, uid, password,
                "product.product", "write",
                [[product_id], data]
            )
        else:
            product_id = models.execute_kw(
                db, uid, password,
                "product.product", "create",
                [data]
            )
            print(f"Created product with ID {product_id}: {data}")
    except Exception as e:
        print(f"Failed to create product {data.get('name', 'unknown')}: {e}")

    # Create the product supplierinfo
    data_supplierinfo['product_id'] = product_id
    print(f"Data for creation: {data_supplierinfo}")
    try:
        product_supplierinfo_id = models.execute_kw(
            db, uid, password,
            "product.supplierinfo", "create",
            [data_supplierinfo]
        )
        print(f"Created product suppliferinfo with ID {product_supplierinfo_id}: {data_supplierinfo}")
    except Exception as e:
        print(f"Failed to create product supplierinfo {data_supplierinfo.get('name', 'unknown')}: {e}")
