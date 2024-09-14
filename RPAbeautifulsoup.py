import streamlit as st
import pandas as pd
import requests
from bs4 import BeautifulSoup

# Title of the app
st.title("RPA Assignment - Beautiful Soup")

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

if 'sorted_requisition' in locals():
    product_name = sorted_requisition.iloc[0]['Product Name'] 

    def get_prices_ebay(product_name):
        url = f"https://www.ebay.com/sch/i.html?_nkw={product_name}"
        response = requests.get(url)
        soup = BeautifulSoup(response.text, 'html.parser')

        prices, ratings, captions, sources = [], [], [], []
        items = soup.select('.s-item')
        for item in items:
            price = item.select_one('.s-item__price')
            if price:
                prices.append(price.text.strip())
            else:
                prices.append("N/A")

            rating_element = item.select_one('.s-item__reviews span.clipped')
            if rating_element:
                rating_text = rating_element.text.strip()
                numeric_rating = float(rating_text.split()[0])
                ratings.append(numeric_rating)
            else:
                ratings.append(0.0)

            caption_element = item.select_one('.s-item__title')
            if caption_element:
                captions.append(caption_element.text.strip())
            else:
                captions.append("No caption available")
                
            sources.append("eBay")
        return prices, ratings, captions, sources

    def format_rating(rating):
    try:
        rating = float(rating)  # Ensure rating is a float
        full_stars = int(rating)
        if rating % 1 >= 0.75:
            return '★' * (full_stars + 1) + '☆' * (5 - full_stars - 1)
        elif rating % 1 >= 0.25:
            return '★' * full_stars + '½' + '☆' * (5 - full_stars - 1)
        else:
            return '★' * full_stars + '☆' * (5 - full_stars)
    except ValueError:
        return 'No Rating'  # Handle cases where rating can't be converted
        
    prices_ebay, ratings_ebay, captions_ebay, sources_ebay = get_prices_ebay(product_name)

    data = {
        'Sources': sources_ebay,
        'Caption': captions_ebay,
        'Price': prices_ebay,
        'Rating': [rating_to_stars(r) for r in ratings_ebay]
    }

    # Create DataFrame
    results_df = pd.DataFrame(data)
    results_df.to_excel('BScomparison_table(ebay).xlsx', index=False)

    # Alibaba price retrieval
    def get_prices_alibaba(product_name):
        url = f"https://www.alibaba.com/trade/search?SearchText={product_name}"
        response = requests.get(url)
        soup = BeautifulSoup(response.text, 'html.parser')

        prices, ratings, captions, sources = [], [], [], []
        items = soup.select('.search-card')
        for item in items:
            price = item.select_one('.search-card-e-price-main')
            if price:
                prices.append(price.text.strip())
            else:
                prices.append("N/A")

            rating_element = item.select_one('.x-star-rating')
            if rating_element:
                rating_text = rating_element.text.strip()
                numeric_rating = float(rating_text.split()[0])
                ratings.append(numeric_rating)
            else:
                ratings.append(0.0)

            caption_element = item.select_one('.search-e-card-title')
            if caption_element:
                captions.append(caption_element.text.strip())
            else:
                captions.append("No caption available")

            sources.append("Alibaba")
        
        return prices, ratings, captions, sources

    prices_alibaba, ratings_alibaba, captions_alibaba, sources_alibaba = get_prices_alibaba(product_name)

    data = {
        'Sources': sources_alibaba,
        'Caption': captions_alibaba,
        'Price': prices_alibaba,
        'Rating': [rating_to_stars(r) for r in ratings_alibaba]
    }

    # Create DataFrame
    results_df = pd.DataFrame(data)
    results_df.to_excel('BScomparison_table(alibaba).xlsx', index=False)

    # Combine DataFrames
    alibaba_df = pd.read_excel('BScomparison_table(alibaba).xlsx')
    ebay_df = pd.read_excel('BScomparison_table(ebay).xlsx')

    alibaba_df['Source'] = 'Alibaba'
    ebay_df['Source'] = 'eBay'

    combined_df = pd.concat([alibaba_df, ebay_df], ignore_index=True)

    # Clean Price column and handle NaN values
    combined_df['Price'] = combined_df['Price'].fillna('0')

    def extract_lowest_price(price_str):
        try:
            price_str = price_str.replace('$', '').replace(',', '')
            prices = price_str.split(' to ')
            return float(prices[0])
        except (ValueError, IndexError):
            return float('inf')  # Return a large number if conversion fails

    combined_df['Price'] = combined_df['Price'].apply(extract_lowest_price)
    filtered_df = combined_df[combined_df['Rating'] == '★★★★★']

    if not filtered_df.empty:
    lowest_price_row = filtered_df.loc[filtered_df['Price'].idxmin()]
    
    caption = lowest_price_row['Caption']
    price = f"${lowest_price_row['Price']:.2f}"
    rating = format_rating(lowest_price_row['Rating'])  # This is where the updated function is called
    source = lowest_price_row['Source']

    st.write("\nLowest Price with Highest Rating (5 stars):")
    st.write(f"Caption: {caption}")
    st.write(f"Price: {price}")
    st.write(f"Rating: {rating}")
    st.write(f"Source: {source}")
else:
    st.write("No products with a 5-star rating found.")

    # Report generation (optional)
    from reportlab.lib.pagesizes import letter
    from reportlab.pdfgen import canvas

    def generate_report(file1, file2, output_pdf):
        df1 = pd.read_excel(file1)
        df2 = pd.read_excel(file2)

        combined_df = pd.concat([df1, df2], ignore_index=True)
        combined_df['Price'] = pd.to_numeric(combined_df['Price'].replace({'\$': '', ',': ''}, regex=True), errors='coerce')

        best_product = combined_df.loc[combined_df['Rating'] == '★★★★★'].nsmallest(1, 'Price')

        if not best_product.empty:
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
            st.write("Report generated successfully.")
        else:
            st.write("No products with a 5-star rating found.")

    generate_report('BScomparison_table(alibaba).xlsx', 'BScomparison_table(ebay).xlsx', 'product_report.pdf')
