import stripe
import pyodbc
import pandas as pd


#making connections to the database
server = '192.168.17.201'
database = 'Integra'
username = 'sa'
password = 'sa123'
port = 63578
driver = 'ODBC Driver 17 for SQL Server'   # Driver name for SQL Server
# driver = 'SQL Server'  # Driver name for SQL Server
# Construct the connection string with the driver name
try:
    connection_string = f"DRIVER={{{driver}}};SERVER={server},{port};DATABASE={database};UID={username};PWD={password}"
    # Establish a connection
    connection = pyodbc.connect(connection_string)

    # data = pd.read_sql("SELECT * from patients", conn)
except Exception as e:
    print('Exception occurred', e)



def find_customers_with_null_shipping():
    try:
        # List all customers
        customers = stripe.Customer.list()

        # Filter customers with null shipping details and collect their IDs and metadata
        customers_with_null_shipping = []
        for customer in customers.data:
            shipping = customer.shipping
            if shipping is None or any(value is None for value in shipping.address.values()):
                customer_info = {
                    "id": customer.id,
                    "metadata": customer.metadata,
                    "name" : customer.name
                }
                customers_with_null_shipping.append(customer_info)

        return customers_with_null_shipping

    except stripe.error.StripeError as e:
        # Handle any Stripe API errors
        print(f"Error: {e}")
        return None


def search_customers_by_metadata(metadata_value):
    try:
        # Create the query string dynamically using the metadata value from the data frame
        metadata_query = 'metadata["PPSN"] : "' + str(metadata_value) + '"'
        # Use the search method with the constructed query
        customers = stripe.Customer.search(query=metadata_query)

        # Return the list of customers matching the query
        return customers.data

    except stripe.error.StripeError as e:
        # Handle any Stripe API errors
        print(f"Error: {e}")
        return None


#reading data from the excel files
df_beneavin=pd.read_excel("Beneavin_Data_July_2023.xlsx")
df_integra=pd.read_excel("integra_data.xlsx")


#merging the two dataframes
merged_df = pd.merge(df_integra,df_beneavin, left_on = 'ppsn_number', right_on = 'PPSN_NUMBER', how = 'inner')

merged_df=merged_df.drop_duplicates("PPSN_NUMBER")

merged_df.dropna(subset=['ppsn_number'], inplace=True)
#sdasdas


#Using stripe to create customers
stripe.api_key = "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"

for i in merged_df.index:
    try:
        stripe.Customer.create(
                name=str(merged_df['First Name'][i] + ' ' + merged_df['SURNAME_x'][i]),
                id=str(merged_df['Random_ID'][i]),
                metadata = {
                    'id': str(merged_df['PPSN_NUMBER'][i])
                }
                
        )
    except Exception as e:
        print('Error occurred ',e)
