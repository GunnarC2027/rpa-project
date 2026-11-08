# Insert your python code here


# INSERT ANSWER FOR QUESTION 1 HERE

while True:
    directory = select_file('Select any file in the directory.')
    directory = directory.split('/')
    del directory[-1]
    directory = '/'.join(directory) + '/'

    filenames = os.listdir(directory)
    if(
        'ProductData.xlsx' in filenames and
        'MasterInvoice.pdf' in filenames and
        'Order1.xlsx' in filenames and
        'Order2.xlsx' in filenames and
        'Order3.xlsx' in filenames and
        'Order4.xlsx' in filenames):
        break
    else:
        show_message("Error", "Please re-select directory.")

# INSERT ANSWER FOR QUESTION 2 HERE

product_descriptions = {}
product_prices = {}
df_description = pd.read_excel(directory+'ProductData.xlsx', sheet_name='Descriptions')
df_prices = pd.read_excel(directory+'ProductData.xlsx', sheet_name='Prices')

for index,row in df_description.iterrows():
    Description = row['Description']
    Unit = row['Unit number']
    product_descriptions[Unit] = Description
for index,row in df_prices.iterrows():
    Price = row['Price per unit']
    Unit = row['Unit number']
    product_prices[Unit] = Price

print(product_descriptions)
print(product_prices)

# INSERT ANSWER FOR QUESTION 3 HERE

def get_order_data(invoice_num, filename, directory):
    update_pdf = {}
    template  = openpyxl.load_workbook(directory + filename)
    selected_sheet = template['OrderForm']
    name = selected_sheet['C14'].value
    email = selected_sheet['C13'].value
    date = selected_sheet['C9'].value
    date  = date.strftime('%m/%d/%Y')
    address = selected_sheet['C16'].value
    city = selected_sheet['C17'].value
    state = selected_sheet['C18'].value
    zipcode = selected_sheet['C19'].value
    company_name = selected_sheet['C15'].value
    amount0 = 0
    amount1 = 0
    amount2 = 0
    amount3 = 0
    amount4 = 0

    unit_number0 = selected_sheet['B24'].value
    if pd.isnull(unit_number0) == False:
        description0 = product_descriptions[unit_number0]
        price0 = product_prices[unit_number0]
        quantity0 = selected_sheet['C24'].value
        amount0 = quantity0 * price0
        

        update_pdf['Description0'] = description0
        update_pdf['Quantity0'] = quantity0
        update_pdf['Price0'] = price0
        amount0_round = f'${amount0:,.2f}'
        update_pdf['Amount0'] = amount0_round

    unit_number1 = selected_sheet['B25'].value
    if pd.isnull(unit_number1) == False:
        description1 = product_descriptions[unit_number1]
        price1 = product_prices[unit_number1]
        quantity1 = selected_sheet['C25'].value
        amount1 = quantity1 * price1
        

        update_pdf['Description1'] = description1
        update_pdf['Quantity1'] = quantity1
        update_pdf['Price1'] = price1
        amount1_round = f'${amount1:,.2f}'
        update_pdf['Amount1'] = amount1_round
   
    unit_number2 = selected_sheet['B26'].value
    if pd.isnull(unit_number2) == False:
        description2 = product_descriptions[unit_number2]
        price2 = product_prices[unit_number2]
        quantity2 = selected_sheet['C26'].value
        amount2 = quantity2 * price2
        

        update_pdf['Description2'] = description2
        update_pdf['Quantity2'] = quantity2
        update_pdf['Price2'] = price2
        amount2_round = f'${amount2:,.2f}'
        update_pdf['Amount2'] = amount2_round

    unit_number3 = selected_sheet['B27'].value
    if pd.isnull(unit_number3) == False:
        description3 = product_descriptions[unit_number3]
        price3 = product_prices[unit_number3]
        quantity3 = selected_sheet['C27'].value
        amount3 = quantity3 * price3
        

        update_pdf['Description3'] = description3
        update_pdf['Quantity3'] = quantity3
        update_pdf['Price3'] = price3
        amount3_round = f'${amount3:,.2f}'
        update_pdf['Amount3'] = amount3_round
   
    unit_number4 = selected_sheet['B28'].value
    if pd.isnull(unit_number4) == False:
        description4 = product_descriptions[unit_number4]
        price4 = product_prices[unit_number4]
        quantity4 = selected_sheet['C28'].value
        amount4 = quantity4 * price4
        

        update_pdf['Description4'] = description4
        update_pdf['Quantity4'] = quantity4
        update_pdf['Price4'] = price4
        amount4_round = f'${amount4:,.2f}'
        update_pdf['Amount4'] = amount4_round

    update_pdf['InvoiceNum'] = invoice_num
    update_pdf['Date'] = date
    Total = amount0 + amount1 + amount2 + amount3 + amount4
    Total = f'${Total:,.2f}'
    update_pdf['Total'] = Total
    update_pdf['CustomerAddress1'] = address
    update_pdf['CustomerName'] = name
    update_pdf['CustomerCompany'] = company_name
    update_pdf['CustomerAddress2'] = f'{city}, {state} {zipcode}'
    return update_pdf, name, email, date

order_data, name, email, date = get_order_data('8515', 'Order1.xlsx', directory)
print(order_data)
print(name, email, date)


# INSERT ANSWER FOR QUESTION 4 HERE

def create_invoice(invoice_num, order_data, directory): 
    invoice_file = f'Invoice {invoice_num}.pdf'
    create_pdf(directory+'MasterInvoice.pdf', directory+invoice_file, order_data)
    
create_invoice('8515', order_data, directory)

# INSERT ANSWER FOR QUESTION 5 HERE

def send_email(invoice_num, name, email, date, directory):

    gmail_user = "gunnar.chris13@gmail.com"
    gmail_password_path = select_file('Please select the gmail password file')
    gmail_password = open(gmail_password_path).read()

    recipients = [f'{email}']

    subject = f'Your Order from PythonTech Co. - Invoice {invoice_num}'

    # Body of the email

    name = f'{name}'

    body = f"""Dear {name},
        
    Attached is your invoice dated {date}.

    Thank you for your business.

    PythonTech Co.
    10 Python Drive
    Pleasant Grove, UT 84062
    Phone: (801)555-2372
    Email: sales@pythontechco.com"""

    # Initialize the email message

    message = EmailMessage()
    message['From'] = gmail_user
    message['To'] = recipients
    message['Subject'] = subject
    message.set_content(body)

    invoice_file = f'Invoice {invoice_num}.pdf'
    path_to_file = directory + invoice_file # MAC
    message = add_attachment(path_to_file, message) # This function is available from the analyticskit package
            
    # Log in to the Gmail server

    mail_server = smtplib.SMTP_SSL('smtp.gmail.com')
    mail_server.login(gmail_user, gmail_password)

    # Send the email

    mail_server.send_message(message)
    mail_server.quit()

send_email(8515, 'Thomas Holt', 'gunnar.chris13@gmail.com', '11/13/2023', directory)


filenames = os.listdir(directory)
filenames.sort()
count = 8515
for filename in filenames: 
    if 'Order' in filename and 'xlsx' in filename: 
        order_data, name, email, date = get_order_data(count, filename, directory)
        create_invoice(count, order_data, directory)
        send_email(count, name, email, date, directory)
        count = count + 1
