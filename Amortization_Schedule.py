import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.chart import BarChart, Reference, Series
from openpyxl.styles import NamedStyle

# get user input
loan_amount = float(input("Enter the loan amount: "))
interest_rate = float(input("Enter the interest rate (as a percentage): "))
loan_term = int(input("Enter the loan term (in years): "))
payments_per_year = int(input("Enter the number of payments per year: "))
start_date_str = input("Enter the start date of the loan (in mm/dd/yyyy format): ")
extra_payment = float(input("Enter the extra monthly payment amount (optional, press enter to skip): ") or 0)

# convert start date to datetime object
start_date = pd.to_datetime(start_date_str, format='%m/%d/%Y')

# calculate loan payment
r = interest_rate / 100 / payments_per_year  # interest rate per payment period
n = loan_term * payments_per_year  # total number of payments
P = loan_amount  # principal
payment = (r * P) / (1 - (1 + r) ** -n)  # payment amount

def get_monthly_payment(loan_amount, interest_rate, n):
    """Calculate the monthly payment for a loan."""
    r = interest_rate / 12
    payment = (r * loan_amount) / (1 - (1 + r)**(-n))
    return payment

def create_amortization_schedule(loan_amount, interest_rate, n, start_date, payments_per_year=12, extra_payment=None):
    """Create an amortization schedule for a loan."""
    schedule = pd.DataFrame(columns=['Payment', 'Payment Date', 'Payment Amount', 'Interest Paid', 'Principal Paid', 'Extra Payment', 'Balance'])
    balance = loan_amount
    r = interest_rate / 12
    payment = get_monthly_payment(loan_amount, interest_rate, n)
    for i in range(n):
        interest_paid = balance * r
        principal_paid = payment - interest_paid
        if extra_payment and balance > extra_payment:
            principal_paid += extra_payment
        balance -= principal_paid
        if balance <= 0:
            # stop amortization when balance reaches 0
            balance = 0
            row_data = [i+1, start_date + pd.DateOffset(months=i), payment, interest_paid, principal_paid-extra_payment, extra_payment, balance]
            schedule = schedule.append(pd.Series(row_data, index=schedule.columns), ignore_index=True)
            break
        payment_date = start_date + pd.DateOffset(months=i)
        if payments_per_year == 12:
            # monthly payments
            row_data = [i+1, payment_date.strftime('%m/%d/%Y'), payment, interest_paid, principal_paid-extra_payment, extra_payment, balance]
            schedule = schedule.append(pd.Series(row_data, index=schedule.columns), ignore_index=True)
        elif i % (12//payments_per_year) == 0:
            # quarterly, semi-annual, or annual payments
            row_data = [i+1, payment_date.strftime('%m/%d/%Y'), payment, interest_paid, principal_paid-extra_payment, extra_payment, balance]
            schedule = schedule.append(pd.Series(row_data, index=schedule.columns), ignore_index=True)
            
    extra_schedule = pd.DataFrame(columns=['Payment', 'Payment Date', 'Payment Amount', 'Interest Paid', 'Principal Paid', 'Extra Payment', 'Balance'])
    balance = loan_amount
    r = interest_rate / 12
    for i in range(n):
        interest_paid = balance * r
        principal_paid = 0
        if extra_payment and balance > extra_payment:
            principal_paid = extra_payment
        balance -= principal_paid
        if balance <= 0:
            # stop amortization when balance reaches 0
            balance = 0
            row_data = [i+1, start_date + pd.DateOffset(months=i), extra_payment, interest_paid, principal_paid, extra_payment, balance]
            extra_schedule = extra_schedule.append(pd.Series(row_data, index=extra_schedule.columns), ignore_index=True)
            break
        payment_date = start_date + pd.DateOffset(months=i)
        row_data = [i+1, payment_date.strftime('%m/%d/%Y)
   
# create workbook and add schedule to first sheet
wb = Workbook()
ws1 = wb.active
ws1.title = "Amortization Schedule"
schedule = create_amortization_schedule(loan_amount, interest_rate, n, start_date, payments_per_year, extra_payment)
for r in dataframe_to_rows(create_amortization_schedule, index=False, header=True):
    ws1.append(r)

# create extra schedule and add to second sheet
extra_schedule = create_amortization_schedule(loan_amount, interest_rate, n, start_date, payments_per_year, extra_payment)
ws2 = wb.create_sheet("Extra Payment Schedule")
for r in dataframe_to_rows(extra_schedule, index=False, header=True):
    ws2.append(r)

# save workbook to a file
wb.save('amortization_schedule.xlsx')
