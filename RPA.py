import streamlit as st
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager

# Set up Chrome options
options = Options()
options.headless = True  # You can set this to False for debugging
options.add_argument("--no-sandbox")  # Add this line
options.add_argument("--disable-dev-shm-usage")  # Add this line

# Initialize the Chrome driver
service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=options)

# Title of the app
st.title("RPA Assignment - Selenium")

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
