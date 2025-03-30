import pandas as pd
import numpy as np
import os

# Load customer data from Excel file
input_file = "Core.xlsx"  # Ensure this file exists in the same directory
output_folder = "Generated_Stock_Statements"

# Create an output folder if it doesn't exist
os.makedirs(output_folder, exist_ok=True)

# Read customer data from the Excel file
df_customers = pd.read_excel(input_file)

# Function to generate stock statement
def generate_stock_statement(customer_name, limit):
    months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", 
              "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
    
    data = []
    for month in months:
        # Randomly vary DP for each month (between 115% to 120% of limit)
        dp = np.random.randint(int(limit * 1.15), int(limit * 1.2))

        # Randomly generate Finished Goods (40% to 60% of DP)
        finished_goods = np.random.randint(dp * 0.4, dp * 0.6)

        # Adjust Book Debts to match the randomly chosen DP
        book_debts = dp - finished_goods  

        # Store the month's data
        data.append([month, finished_goods, book_debts, dp])

    # Convert data to DataFrame
    df = pd.DataFrame(data, columns=["Month", "Finished Goods", "Book Debts", "DP"])

    # Save to Excel with a single sheet
    file_name = os.path.join(output_folder, f"{customer_name}.xlsx")
    df.to_excel(file_name, sheet_name="Stock Statement", index=False)

    print(f"Generated: {file_name}")

# Loop through each customer and generate stock statements
for _, row in df_customers.iterrows():
    generate_stock_statement(row["Unit Name"], row["Limit"])
