# Cryptocurrency Live Data Analysis and Excel Integration

This project fetches live cryptocurrency data for the top 50 cryptocurrencies, analyzes the data, and presents the results in a live-updating Excel sheet. The script is designed to continuously update the Excel file with the latest cryptocurrency prices and key metrics every 5 minutes.

---

## **Features**
- Fetches live cryptocurrency data from the [CoinGecko API](https://www.coingecko.com/en/api).
- Displays the following key metrics:
  - Cryptocurrency Name
  - Symbol
  - Current Price (USD)
  - Market Capitalization
  - 24-hour Trading Volume
  - 24-hour Percentage Price Change
- Performs data analysis:
  - Identifies the top 5 cryptocurrencies by market cap.
  - Calculates the average price of the top 50 cryptocurrencies.
  - Finds the highest and lowest 24-hour percentage price changes.
- Updates data in an Excel sheet every 5 minutes.

---

## **Setup Instructions**

### **Prerequisites**
- Python 3.7 or higher
- Install the required Python libraries:
  ```bash
  pip install requests pandas xlwings
