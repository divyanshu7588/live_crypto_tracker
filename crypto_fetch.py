import requests
import pandas as pd
import time
from openpyxl import load_workbook

#  API ka URL aur parameters define kar rahe hain


url = "https://api.coingecko.com/api/v3/coins/markets"
params = {
    "vs_currency": "usd",   # USD currency me data fetch karna hai
    "order": "market_cap_desc",  # Market Cap ke basis pe sorting
    "per_page": 50,  # Top 50 cryptocurrencies fetch karni hain
    "page": 1
}


#  Excel file ka naam
file_path = "crypto_dataa.xlsx"

def fetch_crypto_data():
    """ API se cryptocurrency data fetch karne ka function"""
    response = requests.get(url, params=params)


    # Agar API ka response theek nahi hai toh error print karega
    if response.status_code != 200:
        print(" API Error:", response.status_code)
        return None

    data = response.json()
    
    
    # Agar API ka response khali hai toh None return karega
    if not data:
        print(" API Response Empty!")
        return None

    #  API Response thik se aagaya hai
    df = pd.DataFrame(data)


    #  DataFrame ke andar required columns hone chahiye
    required_columns = {'name', 'symbol', 'current_price', 'market_cap', 'total_volume', 'price_change_percentage_24h'}
    if not required_columns.issubset(df.columns):
        print(" Kuch Columns missing hain, API response check karo!")
        return None


    # Sirf required columns ko select kar rahe hain
    return df[['name', 'symbol', 'current_price', 'market_cap', 'total_volume', 'price_change_percentage_24h']]

def save_to_excel(df):
    """ Excel file update karne ka function"""
    try:
        with pd.ExcelWriter(file_path, mode="w", engine="openpyxl") as writer:
            df.to_excel(writer, index=False)
        print(" Excel file update ho chuki hai!")
    except Exception as e:
        print(" Excel likhne me error aaya:", e)


#  Ye loop bar bar data fetch karega aur Excel update karega
while True:
    df = fetch_crypto_data()

    if df is not None:
        save_to_excel(df)  
    

    time.sleep(15)  #  (15 seconds) me update karega
