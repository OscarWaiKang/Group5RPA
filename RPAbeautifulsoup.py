import streamlit as st
import pandas as pd
import requests
from bs4 import BeautifulSoup
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

# Title of the app
st.title("RPA Assignment")

# File uploader
uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx", "xls"])

if uploaded_file is not None:
    # Read the Excel file
    try:
        df = pd.read_excel(uploaded_file)

        # Check for 'Required Date' column
        if 'Required Date' in df.columns:
            # Convert and sort the DataFrame
            df['Required Date'] = pd.to_datetime(df['Required Date'])
            sorted_requisition = df.sort_values(by='Required Date')
            st.write("Sorted Requisition Data by Date:")
            st.dataframe(sorted_requisition)

            product_name = sorted_requisition.iloc[0]['Product Name']

            prices_ebay, ratings_ebay, captions_ebay = get_prices_ebay(product_name)
            prices_alibaba, ratings_alibaba, captions_alibaba = get_prices_alibaba(product_name)

            # Create and save DataFrames
            ebay_df = create_dataframe(prices_ebay, ratings_ebay, captions_ebay, "eBay")
            alibaba_df = create_dataframe(prices_alibaba, ratings_alibaba, captions_alibaba, "Alibaba")

            # Combine DataFrames and process prices
            combined_df = combine_dataframes(alibaba_df, ebay_df)
            display_best_product(combined_df)

            # Generate report
            generate_report('BScomparison_table(alibaba).xlsx', 'BScomparison_table(ebay).xlsx', 'product_report.pdf')

        else:
            st.error("The 'Required Date' column is not found in the uploaded file.")

    except Exception as e:
        st.error(f"Error reading the file: {e}")

def get_prices_ebay(product_name):
    url = f"https://www.ebay.com/sch/i.html?_nkw={product_name.replace(' ', '+')}"
    response = requests.get(url)
    soup = BeautifulSoup(response.text, 'html.parser')
    return extract_prices(soup, 's-item', '.s-item__price', '.s-item__title', '.s-item__reviews span.clipped')

def get_prices_alibaba(product_name):
    url = f"https://www.alibaba.com/trade/search?SearchText={product_name.replace(' ', '+')}"
    response = requests.get(url)
    soup = BeautifulSoup(response.text, 'html.parser')
    return extract_prices(soup, '.search-card', '.search-card-e-price-main', '.search-e-card-title', '.x-star-rating')

def extract_prices(soup, item_selector, price_selector, title_selector, rating_selector):
    prices, ratings, captions, sources = [], [], [], []
    items = soup.select(item_selector)
    for item in items:
        prices.append(get_text(item, price_selector, "N/A"))
        ratings.append(get_numeric_rating(item, rating_selector))
        captions.append(get_text(item, title_selector, "No caption available"))
        sources.append("Source")
    return prices, ratings, captions, sources

def get_text(item, selector, default):
    element = item.select_one(selector)
    return element.text.strip() if element else default

def get_numeric_rating(item, selector):
    rating_element = item.select_one(selector)
    if rating_element:
        rating_text = rating_element.text.strip()
        try:
            return float(rating_text.split()[0])
        except (ValueError, IndexError):
            return 0.0
    return 0.0

def create_dataframe(prices, ratings, captions, source_name):
    data = {
        'Sources': [source_name] * len(prices),
        'Caption': captions,
        'Price': prices,
        'Rating': [rating_to_stars(r) for r in ratings]
    }
    df = pd.DataFrame(data)
    df.to_excel(f'BScomparison_table({source_name.lower()}).xlsx', index=False)
    return df

def combine_dataframes(alibaba_df, ebay_df):
    alibaba_df['Source'] = 'Alibaba'
    ebay_df['Source'] = 'eBay'
    combined_df = pd.concat([alibaba_df, ebay_df], ignore_index=True)
    combined_df['Price'] = combined_df['Price'].fillna('0')
    combined_df['Price'] = combined_df['Price'].apply(extract_lowest_price)
    return combined_df

def extract_lowest_price(price_str):
    try:
        price_str = price_str.replace('$', '').replace(',', '')
        prices = price_str.split(' to ')
        return float(prices[0])
    except (ValueError, IndexError):
        return float('inf')

def display_best_product(combined_df):
    max_rating = combined_df['Rating'].max()
    highest_rated_products = combined_df[combined_df['Rating'] == max_rating]

    if not highest_rated_products.empty:
        lowest_price_row = highest_rated_products.loc[highest_rated_products['Price'].idxmin()]
        st.write("\nLowest Price with Highest Rating:")
        st.write(f"Caption: {lowest_price_row['Caption']}")
        st.write(f"Price: ${lowest_price_row['Price']:.2f}")
        st.write(f"Rating: {lowest_price_row['Rating']}")
        st.write(f"Source: {lowest_price_row['Source']}")
    else:
        st.write("No products with a valid rating found.")

def generate_report(file1, file2, output_pdf):
    df1 = pd.read_excel(file1)
    df2 = pd.read_excel(file2)
    combined_df = pd.concat([df1, df2], ignore_index=True)
    combined_df['Price'] = pd.to_numeric(combined_df['Price'].replace({'\$': '', ',': ''}, regex=True), errors='coerce')
    best_product = combined_df.loc[combined_df['Rating'] == combined_df['Rating'].max()].nsmallest(1, 'Price')

    if not best_product.empty:
        create_pdf_report(best_product, output_pdf)
        st.write("Report generated successfully.")
    else:
        st.write("No products with a valid rating found.")

def create_pdf_report(best_product, output_pdf):
    source = best_product['Sources'].values[0]
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
