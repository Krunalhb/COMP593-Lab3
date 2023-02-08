from sys import argv, exit
import os
from datetime import date
import pandas as pd
import re
import xlsxwriter


def get_sales_csv():
  
    #check whether a command line parameter was provided
    if len(argv) >= 2:
        sales_csv = argv[1]

    #check whether the csv file path exists
        if os.path.isfile(sales_csv):
            return sales_csv
        else:
            print('ERROR: csv file path does not exist')
            exit('scrpit execution aborted')

        
    else:
        print('ERROR: no csv file path has been provided')
        exit('Scrpit Execution Aborted')

def get_order_dir(sales_csv):

    # get directory path of sales data csv file
    sales_dir = os.path.dirname(sales_csv)

    # determine orders directory names (Orders_YYYY-MM-DD)
    todays_date = date.today().isoformat()
    order_dir_name = "Orders_" + todays_date

    # build the full path of the orders directory
    Order_dir = os.path.join(sales_dir, order_dir_name)

    # make the orders directory if it doesn't already exists
    if not os.path.exists(Order_dir):
        os.makedirs(Order_dir)    

    return Order_dir

def split_sales_into_orders(sales_csv, order_dir):
    
    # read data from sales data csv in the dataframe
    sales_df = pd.read_csv(sales_csv)

    # inserting a new column for total price
    sales_df.insert(7, 'TOTAL PRICE', sales_df['ITEM QUANTITY'] * sales_df['ITEM PRICE'])

    # droping unwanted columns
    sales_df.drop( columns=['ADDRESS', 'CITY', 'STATE', 'POSTAL CODE', 'COUNTRY'], inplace= True)

    for order_id, order_df in sales_df.groupby('ORDER ID'):
        
        # droping the order id column
        order_df.drop( columns= ['ORDER ID'], inplace= True)

        # sort the order by item number
        order_df.sort_values(by= ['ITEM NUMBER'], inplace= True)

        # add grand total row at the bottom
        grand_total = order_df['TOTAL PRICE'].sum()
        grand_total_df = pd.DataFrame({'ITEM PRICE': ['GRAND TOTAL:'], 'TOTAL PRICE': [grand_total]})
        order_df = pd.concat([order_df, grand_total_df])

        # determine the path of the order file
        customer_name = order_df['CUSTOMER NAME'].values[0]
        customer_name = re.sub(r'\W', '', customer_name)
        order_file_name = 'Order' + str(order_id) + '_' + customer_name + '.xlsx'
        order_file_path = os.path.join(order_dir, order_file_name)

        # save the ordered information to an excel spreadsheet
        sheet_name = 'Order #' + str(order_id)
        # order_df.to_excel(order_file_path, index= False, sheet_name= sheet_name)
        
        
        writer = pd.ExcelWriter(order_file_path, engine= 'xlsxwriter')
        order_df.to_excel(writer, index= False, sheet_name = sheet_name)

        workbook = writer.book
        worksheet = writer.sheets[sheet_name]
        
        # format the money values
        money_fmt = workbook.add_format({'num_format': '$#,##0.00', 'bold': True})

        # re-sizing the column width
        worksheet.set_column('A:E', 15)
        worksheet.set_column('F:G', 15, money_fmt)
        worksheet.set_column('H:J', 15)

        writer.save()

sales_csv = get_sales_csv()
order_dir = get_order_dir(sales_csv)
split_sales_into_orders(sales_csv, order_dir)