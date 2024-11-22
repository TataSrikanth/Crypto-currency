import requests
import pandas as pd
import xlwings as xw
import time

# Function to fetch cryptocurrency data
def fetch_crypto_data():
    url = "https://api.coingecko.com/api/v3/coins/markets"
    params = {
        'vs_currency': 'usd',
        'order': 'market_cap_desc',
        'per_page': 50,
        'page': 1
    }
    try:
        response = requests.get(url, params=params)
        response.raise_for_status()
        data = response.json()
        return data
    except requests.exceptions.RequestException as e:
        print(f"Error fetching data: {e}")
        return []

# Function to analyze data
def analyze_data(data):
    df = pd.DataFrame(data)
    
    # Fill NaN values for price change percentage
    df['price_change_percentage_24h'] = df['price_change_percentage_24h'].fillna(0)
    
    # Top 5 cryptocurrencies by Market Capitalization
    top_5 = df.nlargest(5, 'market_cap')[['name', 'symbol', 'market_cap']]
    
    # Average price of the top 50 cryptocurrencies
    avg_price = df['current_price'].mean()
    
    # Cryptocurrency with the highest and lowest 24h percentage price change
    highest_change = df.loc[df['price_change_percentage_24h'].idxmax()]
    lowest_change = df.loc[df['price_change_percentage_24h'].idxmin()]
    
    return top_5, avg_price, highest_change, lowest_change

# Function to update Excel with live data in the same file
def update_excel(file_path, data, sheet_name="Live Crypto Data"):
    try:
        # Open the existing workbook
        wb = xw.Book(file_path)
        
        # Check if the sheet exists, if not, add a new sheet
        if sheet_name in [sheet.name for sheet in wb.sheets]:
            sheet = wb.sheets[sheet_name]
        else:
            sheet = wb.sheets.add(sheet_name)
        
        # Clear the existing content in the sheet
        sheet.clear()
        
        # Convert data to DataFrame
        df = pd.DataFrame(data)
        
        # Write new data to the Excel sheet
        sheet.range("A1").value = ["Cryptocurrency Name", "Symbol", "Current Price (USD)", 
                                   "Market Cap", "24h Trading Volume", "24h Price Change (%)"]
        sheet.range("A2").value = df[["name", "symbol", "current_price", 
                                      "market_cap", "total_volume", "price_change_percentage_24h"]].values.tolist()
        
        print("Excel file updated with live data.")
    except Exception as e:
        print(f"Error updating Excel: {e}")

# Main function
if __name__ == "__main__":
    try:
        # Specify the path to your existing Excel file
        excel_file_path = "Book1.xlsx"  # Replace with your file path
        
        print("Fetching live cryptocurrency data...")
        while True:
            # Fetch live data
            crypto_data = fetch_crypto_data()
            if not crypto_data:
                time.sleep(300)
                continue
            
            # Analyze the data
            top_5, avg_price, highest_change, lowest_change = analyze_data(crypto_data)
            
            # Print the analysis to the console
            print("\n--- Cryptocurrency Analysis ---")
            print("Top 5 Cryptocurrencies by Market Cap:")
            print(top_5)
            print(f"Average Price of Top 50 Cryptocurrencies: ${avg_price:.2f}")
            print(f"Highest 24h % Change: {highest_change['name']} ({highest_change['price_change_percentage_24h']:.2f}%)")
            print(f"Lowest 24h % Change: {lowest_change['name']} ({lowest_change['price_change_percentage_24h']:.2f}%)")
            
            # Update Excel with live data
            update_excel(excel_file_path, crypto_data)
            
            # Wait 5 minutes before fetching the next update
            time.sleep(300)
    except KeyboardInterrupt:
        print("\nLive updating stopped.")
