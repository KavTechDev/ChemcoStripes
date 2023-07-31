import stripe
import pyodbc
import pandas as pd


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


#Using Stripe Invoice APi to create invoices for the customers
for i in final_merged_df.index:
    print('i =', i)
    stripe.api_key = "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"

    try:
        customer_id = final_merged_df.loc[i, 'PPSN_NUMBER_x']
        april_amount = str(int (final_merged_df.loc[i, 'person_total_april']* 100 ))
        may_amount = str(int(final_merged_df.loc[i, 'person_total_may'] * 100 ))
       
        april_description = 'April Prescription Charges'
        may_description = 'May Prescription Charges'

        invoice = stripe.Invoice.create(
            customer=customer_id
        )

        april_invoice_item = stripe.InvoiceItem.create(
            customer=customer_id,
            invoice=invoice.id,
            amount=april_amount,
            currency='eur',
            description=april_description
        )

        may_invoice_item = stripe.InvoiceItem.create(
            customer=customer_id,
            invoice=invoice.id,
            amount=may_amount,
            currency='eur',
            description=may_description
        )
    except stripe.error.StripeError as e:
        print('Error occurred:', str(e))

