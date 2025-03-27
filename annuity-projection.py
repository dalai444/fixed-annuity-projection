# Creating list of variables (will be inputs from excel sheet eventually)
# Slightly unnessary since variables could be declared when gathering values from excel

principal = 10000 # The initial amount
interest = 0.05 # Interest collected every year
frequency = 12 # Number of payments per year
years = 2 # Number of Years annuity will be paid out
payment = 0 # How much will be paid each period

paymentType = "Due" # Will either be Ordinary (payments at end) or Due (payments at beginning)

periodic_interest = interest / frequency
total_periods = frequency * years


table = []

# We will use the annuity formula to calculate the payment amounts
if paymentType == "Ordinary":
    payment = principal * ( (periodic_interest*pow(1+periodic_interest, total_periods)) / (pow(1+periodic_interest, total_periods) - 1) )
elif paymentType == "Due":
    payment = principal * ( (periodic_interest*pow(1+periodic_interest, total_periods)) / (pow(1+periodic_interest, total_periods) - 1) ) * (1 + periodic_interest)

begin_balance = principal
total_payment_amount = 0
total_interest_collected = 0


for i in range(total_periods):
    interestCollected = begin_balance*periodic_interest
    balance =  begin_balance + interestCollected - payment

    # Set all of our values for later plotting in excel
    table_row = [i, principal, interestCollected, payment, balance]
    table.append(table_row)

    total_payment_amount += payment
    total_interest_collected += interestCollected

    if(balance <= 0):
        print("Account hits 0 before all payment periods are completed")
        exit

    begin_balance = balance # change the balance to the final balance that was calculated to continue loop
    

for row in table:
    print("Period: " , row[0], " Principal: " , row[1], " Interest Collected: ", row[2], " Payment: ", row[3], " Balance: ", row[4])

