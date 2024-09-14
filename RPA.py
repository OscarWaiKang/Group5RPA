import streamlit as st
import pandas as pd
import time
import random
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# Title of the app
st.title("RPA Assignment")

# File uploader
uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx", "xls"])

if uploaded_file is not None:
    # Read the Excel file
    try:
        df = pd.read_excel(uploaded_file)

        # Check if 'Required Date' column exists
        if 'Required Date' in df.columns:
            # Convert 'Required Date' to datetime
            df['Required Date'] = pd.to_datetime(df['Required Date'])

            # Sort the DataFrame by 'Required Date'
            sorted_requisition = df.sort_values(by='Required Date')
            st.write("Sorted Requisition Data by Date:")
            st.dataframe(sorted_requisition)
        else:
            st.error("The 'Required Date' column is not found in the uploaded file.")

    except Exception as e:
        st.error(f"Error reading the file: {e}")
        print(e)  # Print the error to console for debugging

# Set up Chrome options for headless mode
options = Options()
options.headless = True  # Run in headless mode
options.add_argument("--no-sandbox")  # Bypass OS security model
options.add_argument("--disable-dev-shm-usage")  # Overcome limited resource problems

# Set up ChromeDriver using webdriver-manager for automatic management
service = Service(ChromeDriverManager().install())

# Initialize the Chrome driver
driver = webdriver.Chrome(service=service, options=options)

def get_prices_ebay(product_name):
    driver.get(f"https://www.ebay.com/sch/i.html?_nkw={product_name}")
    try:
        # Wait for the product listings to load
        item_elements = WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located((By.CSS_SELECTOR, '.s-item'))
        )

        prices = []
        ratings = []

        for item in item_elements:
            try:
                # Extract price
                price_element = item.find_element(By.CSS_SELECTOR, '.s-item__price')
                price = price_element.text.strip()
                prices.append(price)

                # Extract rating (if available)
                try:
                    rating_element = item.find_element(By.CSS_SELECTOR, '.s-item__reviews span.clipped')
                    rating_text = rating_element.text.strip()
                    numeric_rating = float(rating_text.split()[0])  # Get the numeric part
                    ratings.append(numeric_rating)
                except Exception:
                    ratings.append(0.0)

            except Exception as e:
                print(f"Error parsing item: {e}")
                continue

        return prices, ratings

    except Exception as e:
        print(f"Error retrieving prices from eBay: {e}")
        return [], []

def get_prices_alibaba(product_name):
    driver.get(f"https://www.alibaba.com/trade/search?fsb=y&IndexArea=product_en&CatId=&SearchText={product_name}")

    try:
        price_elements = WebDriverWait(driver, 30).until(
            EC.presence_of_all_elements_located((By.CSS_SELECTOR, '.search-card-e-price-main')))
        rating_elements = WebDriverWait(driver, 30).until(
            EC.presence_of_all_elements_located((By.CSS_SELECTOR, '.search-card-e-review')))
        
        prices = [price.text.strip() for price in price_elements if price.text.strip()]
        ratings = [float(rating.text.strip().split()[0]) for rating in rating_elements if rating.text.strip()]

        return prices, ratings
    except Exception as e:
        print(f"Error retrieving prices from Alibaba: {str(e)}")
        return [], []

# Example product name (replace with your actual DataFrame)
if 'sorted_requisition' in locals():
    product_name = sorted_requisition.iloc[0]['Product Name'] 

    # Collect prices and ratings from both vendors
    prices_ebay, ratings_ebay = get_prices_ebay(product_name)
    prices_alibaba, ratings_alibaba = get_prices_alibaba(product_name)

    # Combine prices, ratings, and create vendor list if there are prices
    if prices_ebay or prices_alibaba:
        prices = prices_ebay + prices_alibaba
        ratings = ratings_ebay + ratings_alibaba
        vendors = ['eBay'] * len(prices_ebay) + ['Alibaba'] * len(prices_alibaba)

        # Ensure all lists are the same length
        min_length = min(len(prices), len(ratings), len(vendors))
        prices = prices[:min_length]
        ratings = ratings[:min_length]
        vendors = vendors[:min_length]

        comparison_df = pd.DataFrame({'Vendor': vendors, 'Price': prices, 'Rating': ratings})

        # Clean and convert prices to numeric
        def clean_price(price):
            if "Tap item to see current price" in price or "See price" in price:
                return None
            
            price = price.replace('US$', '').replace('$', '').replace(',', '').strip()

            if 'to' in price:
                price = price.split('to')[0].strip()
            
            if '-' in price:
                price = price.split('-')[0].strip()

            try:
                return float(price)
            except ValueError:
                return None  # Return None if conversion fails

        # Apply price cleaning
        comparison_df['Price'] = comparison_df['Price'].apply(clean_price)
        
        # Drop rows with invalid prices
        comparison_df = comparison_df.dropna(subset=['Price'])

        if not comparison_df.empty:
            # Save the comparison table to an Excel file
            comparison_df.to_excel('comparison_table.xlsx', index=False)

            # Get the highest rating
            max_rating = comparison_df['Rating'].max()

            # Filter the DataFrame to only include rows with the highest rating
            best_rated_df = comparison_df[comparison_df['Rating'] == max_rating]

            # Find the one with the lowest price
            best_combined_row = best_rated_df.loc[best_rated_df['Price'].idxmin()]

            best_combined_price = best_combined_row['Price']
            best_combined_vendor = best_combined_row['Vendor']
            best_combined_rating = best_combined_row['Rating']

            # Generate minutes for the purchase committee based on best combination
            minutes = f"""
[ Product : {product_name} ]
[ Price Comparison ]
{comparison_df.to_string(index=False)}

[ Lowest Price with Best Rating (5.0★) ]
Vendor: {best_combined_vendor}
Price: ${best_combined_price:.2f}
Rating: {best_combined_rating}
"""

            print(minutes)
        else:
            print("No valid price data to display.")
    else:
        print("No data to save.")

driver.quit()

import requests
from bs4 import BeautifulSoup

product_name = sorted_requisition.iloc[0]['Product Name'] 

def get_prices_ebay(product_name):
    url = f"https://www.ebay.com/sch/i.html?_nkw={product_name}"
    response = requests.get(url)
    soup = BeautifulSoup(response.text, 'html.parser')

    prices = []
    ratings = []
    captions = []
    sources = []
    
    items = soup.select('.s-item')
    for item in items:
        price = item.select_one('.s-item__price').text.strip()
        prices.append(price)

        rating_element = item.select_one('.s-item__reviews span.clipped')
        if rating_element:
            rating_text = rating_element.text.strip()
            numeric_rating = float(rating_text.split()[0])
            ratings.append(numeric_rating)
        else:
            ratings.append(0.0)

        caption_element = item.select_one('.s-item__title')
        if caption_element:
            caption = caption_element.text.strip()
            captions.append(caption)
        else:
            captions.append("No caption available")
            
        sources.append("Ebay")
    return prices, ratings, captions, sources

def rating_to_stars(rating):
    full_stars = int(rating)  # Number of full stars
    if rating % 1 >= 0.75:
        return '★' * (full_stars + 1) + '☆' * (5 - full_stars - 1)  # Full star for >= 0.75
    elif rating % 1 >= 0.25:
        return '★' * full_stars + '½' + '☆' * (5 - full_stars - 1)  # Half star for 0.25 to < 0.75
    else:
        return '★' * full_stars + '☆' * (5 - full_stars)  # No half star for < 0.25

prices_ebay, ratings_ebay, captions_ebay, sources_ebay = get_prices_ebay(product_name)

data = {
    'Sources': sources_ebay,
    'Caption': captions_ebay,
    'Price': prices_ebay,
    'Rating': [rating_to_stars(r) for r in ratings_ebay]
}

# Create DataFrame
results_df = pd.DataFrame(data)

# Print product name and DataFrame
print(f"\nProduct: {product_name}\n")
print(results_df.to_string(index=False))

# Save the DataFrame to an Excel file
results_df.to_excel('BScomparison_table(ebay).xlsx', index=False)

# Confirm save
print("\nData has been saved to 'BScomparison_table(ebay).xlsx'.")

import requests
from bs4 import BeautifulSoup

product_name = sorted_requisition.iloc[0]['Product Name'] 

def get_prices_alibaba(product_name):
    url = f"https://www.alibaba.com/trade/search?SearchText={product_name}"
    response = requests.get(url)
    soup = BeautifulSoup(response.text, 'html.parser')

    prices = []
    ratings = []
    captions = []
    sources = []
    
    items = soup.select('.search-card')
    for item in items:
        price = item.select_one('.search-card-e-price-main').text.strip()
        prices.append(price)

        rating_element = item.select_one('.x-star-rating')
        if rating_element:
            rating_text = rating_element.text.strip()
            numeric_rating = float(rating_text.split()[0])
            ratings.append(numeric_rating)
        else:
            ratings.append(0.0)

        caption_element = item.select_one('.search-e-card-title')
        if caption_element:
            caption = caption_element.text.strip()
            captions.append(caption)
        else:
            captions.append("No caption available")
            
        sources.append("Alibaba")
    
    return prices, ratings, captions, sources

def rating_to_stars(rating):
    full_stars = int(rating)
    if rating % 1 >= 0.75:
        return '★' * (full_stars + 1) + '☆' * (5 - full_stars - 1) 
    elif rating % 1 >= 0.25:
        return '★' * full_stars + '½' + '☆' * (5 - full_stars - 1)  
    else:
        return '★' * full_stars + '☆' * (5 - full_stars)

prices_alibaba, ratings_alibaba, captions_alibaba, sources_alibaba= get_prices_alibaba(product_name)

data = {
    'Sources': sources_alibaba,
    'Caption': captions_alibaba,
    'Price': prices_alibaba,
    'Rating': [rating_to_stars(r) for r in ratings_alibaba]
}

# Create DataFrame
results_df = pd.DataFrame(data)

# Print product name and DataFrame
print(f"\nProduct: {product_name}\n")
print(results_df.to_string(index=False))

# Save the DataFrame to an Excel file
results_df.to_excel('BScomparison_table(alibaba).xlsx', index=False)

import pandas as pd

alibaba_df = pd.read_excel('BScomparison_table(alibaba).xlsx')
ebay_df = pd.read_excel('BScomparison_table(ebay).xlsx')

alibaba_df['Source'] = 'Alibaba'
ebay_df['Source'] = 'eBay'

combined_df = pd.concat([alibaba_df, ebay_df], ignore_index=True)
combined_df['Rating'] = combined_df['Rating'].astype(str).replace({'★★★★★': 5, '5.0': 5})

def extract_lowest_price(price_str):
    price_str = price_str.replace('$', '').replace(',', '')
    prices = price_str.split(' to ')
    return float(prices[0])  # Return the first price as float

combined_df['Price'] = combined_df['Price'].apply(extract_lowest_price)
filtered_df = combined_df[combined_df['Rating'] == 5]

if not filtered_df.empty:
    lowest_price_row = filtered_df.loc[filtered_df['Price'].idxmin()]
    
    def format_rating(rating):
        return '★' * int(rating) + '☆' * (5 - int(rating))

    caption = lowest_price_row['Caption']
    price = f"${lowest_price_row['Price']:.2f}"  # Format price with dollar sign and two decimal places
    rating = format_rating(lowest_price_row['Rating'])  # Convert rating to star format
    source = lowest_price_row['Source']  # Get the source from the added column

    print("\nLowest Price with Highest Rating (5 stars):")
    print(f"Caption: {caption}")
    print(f"Price: {price}")
    print(f"Rating: {rating}")
    print(f"Source: {source}")
else:
    print("No products with a 5-star rating found.")


import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

# Function to read Excel files and find the product with the lowest price and highest rating
def generate_report(file1, file2, output_pdf):
    # Read the Excel files
    df1 = pd.read_excel(file1)
    df2 = pd.read_excel(file2)

    combined_df = pd.concat([df1, df2], ignore_index=True)

    combined_df['Price'] = pd.to_numeric(combined_df['Price'].replace({'\$': '', ',': ''}, regex=True), errors='coerce')

    unique_ratings = combined_df['Rating'].unique()
    print("Unique Ratings:", unique_ratings)

    best_product = combined_df.loc[combined_df['Rating'] == '★★★★★'].nsmallest(1, 'Price')

    if not best_product.empty:
        source = best_product['Sources'].values[0]  # Get the source
        product_name = best_product['Caption'].values[0]
        lowest_price = best_product['Price'].values[0]
        rating = best_product['Rating'].values[0]
        
        c = canvas.Canvas(output_pdf, pagesize=letter)
        c.drawString(100, 750, "Product Report")
        c.drawString(100, 670, f"Source: {source}")
        c.drawString(100, 730, f"Product Name: {product_name}")
        c.drawString(100, 710, f"Lowest Price: ${lowest_price:.2f}")
        c.drawString(100, 690, f"Rating: {rating}")
        c.save()
        
        print(f"Report generated successfully.")
    else:
        print("No products with a 5-star rating found.")

file1 = r'C:/Users/user/Downloads/Ai Assignment (1)-20240910T124230Z-001/Ai Assignment (1)/BScomparison_table(ebay).xlsx'  # First Excel file
file2 = r'C:/Users/user/Downloads/Ai Assignment (1)-20240910T124230Z-001/Ai Assignment (1)/BScomparison_table(alibaba).xlsx'  # Second Excel file
output_pdf = r'C:/Users/user/Downloads/Ai Assignment (1)-20240910T124230Z-001/Ai Assignment (1)/product_report.pdf'  # Output PDF file

generate_report(file1, file2, output_pdf)
