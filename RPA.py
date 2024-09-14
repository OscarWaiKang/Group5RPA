{
 "cells": [
  {
   "cell_type": "markdown",
   "id": "8b75765e",
   "metadata": {},
   "source": [
    "# RPA assignment"
   ]
  },
  {
   "cell_type": "markdown",
   "id": "840988d1-f9c4-408f-b735-d1ab48895687",
   "metadata": {},
   "source": [
    "# Install Selenium"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": null,
   "id": "005f4935",
   "metadata": {},
   "outputs": [],
   "source": [
    "pip install pandas selenium openpyxl fpdf"
   ]
  },
  {
   "cell_type": "markdown",
   "id": "78fba90a",
   "metadata": {},
   "source": [
    "# Read excel file input by user"
   ]
  },
  {
   "cell_type": "raw",
   "id": "6fc05d86-4e52-45e3-bc09-edf3bd57940e",
   "metadata": {},
   "source": [
    "(requisition form and sort by date)"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 23,
   "id": "097fe707-853f-4aa9-8a8c-da914601a5e4",
   "metadata": {},
   "outputs": [
    {
     "name": "stderr",
     "output_type": "stream",
     "text": [
      "2024-09-14 17:56:20.896 \n",
      "  \u001b[33m\u001b[1mWarning:\u001b[0m to view this Streamlit app on a browser, run it with the following\n",
      "  command:\n",
      "\n",
      "    streamlit run C:\\Users\\user\\anaconda3\\Lib\\site-packages\\ipykernel_launcher.py [ARGUMENTS]\n"
     ]
    }
   ],
   "source": [
    "import streamlit as st\n",
    "import pandas as pd\n",
    "\n",
    "# Title of the app\n",
    "st.title(\"RPA Assignment\")\n",
    "\n",
    "# File uploader\n",
    "uploaded_file = st.file_uploader(\"Choose an Excel file\", type=[\"xlsx\", \"xls\"])\n",
    "\n",
    "if uploaded_file is not None:\n",
    "    # Read the Excel file\n",
    "    try:\n",
    "        df = pd.read_excel(uploaded_file)\n",
    "\n",
    "        # Check if 'Required Date' column exists\n",
    "        if 'Required Date' in df.columns:\n",
    "            # Convert 'Required Date' to datetime\n",
    "            df['Required Date'] = pd.to_datetime(df['Required Date'])\n",
    "\n",
    "            # Sort the DataFrame by 'Required Date'\n",
    "            sorted_requisition = df.sort_values(by='Required Date')\n",
    "            st.write(\"Sorted Requisition Data by Date:\")\n",
    "            st.dataframe(sorted_requisition)\n",
    "        else:\n",
    "            st.error(\"The 'Required Date' column is not found in the uploaded file.\")\n",
    "\n",
    "    except Exception as e:\n",
    "        st.error(f\"Error reading the file: {e}\")"
   ]
  },
  {
   "cell_type": "markdown",
   "id": "05655c24",
   "metadata": {},
   "source": [
    "# Browse Websites and Collect the lowest prices compared with the best user's rating"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": null,
   "id": "b41a25cc-2522-4be1-8dd2-d472822532fb",
   "metadata": {
    "scrolled": true
   },
   "outputs": [],
   "source": [
    "import time\n",
    "import random\n",
    "from selenium import webdriver\n",
    "from selenium.webdriver.common.by import By\n",
    "from selenium.webdriver.support.ui import WebDriverWait\n",
    "from selenium.webdriver.support import expected_conditions as EC\n",
    "\n",
    "driver = webdriver.Chrome()\n",
    "\n",
    "def get_prices_ebay(product_name):\n",
    "    driver.get(f\"https://www.ebay.com/sch/i.html?_nkw={product_name}\")\n",
    "    try:\n",
    "        # Wait for the product listings to load\n",
    "        item_elements = WebDriverWait(driver, 10).until(\n",
    "            EC.presence_of_all_elements_located((By.CSS_SELECTOR, '.s-item'))\n",
    "        )\n",
    "\n",
    "        prices = []\n",
    "        ratings = []\n",
    "\n",
    "        for item in item_elements:\n",
    "            try:\n",
    "                # Extract price\n",
    "                price_element = item.find_element(By.CSS_SELECTOR, '.s-item__price')\n",
    "                price = price_element.text.strip()\n",
    "                prices.append(price)\n",
    "\n",
    "                # Extract rating (if available)\n",
    "                try:\n",
    "                    rating_element = item.find_element(By.CSS_SELECTOR, '.s-item__reviews span.clipped')\n",
    "                    rating_text = rating_element.text.strip()\n",
    "                    numeric_rating = float(rating_text.split()[0])  # Get the numeric part\n",
    "                    ratings.append(numeric_rating)\n",
    "                except Exception as e:\n",
    "                    # No rating available, append 0.0 or None\n",
    "                    ratings.append(0.0)\n",
    "\n",
    "            except Exception as e:\n",
    "                print(f\"Error parsing item: {e}\")\n",
    "                continue\n",
    "\n",
    "        # Debugging output\n",
    "        print(f\"eBay Prices: {prices}\")\n",
    "        print(f\"eBay Ratings: {ratings}\")\n",
    "\n",
    "        # Return both lists\n",
    "        return prices, ratings\n",
    "\n",
    "    except Exception as e:\n",
    "        print(f\"Error retrieving prices from eBay: {e}\")\n",
    "        return [], []\n",
    "        \n",
    "def print_star_rating(rating):\n",
    "    full_stars = int(rating)\n",
    "    half_star = 1 if rating % 1 >= 0.5 else 0\n",
    "    stars = '★' * full_stars + '½' * half_star + '☆' * (5 - full_stars - half_star)\n",
    "    return stars\n",
    "\n",
    "def get_prices_alibaba(product_name):\n",
    "    driver.get(f\"https://www.alibaba.com/trade/search?fsb=y&IndexArea=product_en&CatId=&SearchText={product_name}\")\n",
    "\n",
    "    try:\n",
    "        price_elements = WebDriverWait(driver, 30).until(\n",
    "            EC.presence_of_all_elements_located((By.CSS_SELECTOR, '.search-card-e-price-main')))\n",
    "        rating_elements = WebDriverWait(driver, 30).until(\n",
    "            EC.presence_of_all_elements_located((By.CSS_SELECTOR, '.search-card-e-review')))\n",
    "        \n",
    "        prices = [price.text.strip() for price in price_elements if price.text.strip()]\n",
    "        ratings = [float(rating.text.strip().split()[0]) for rating in rating_elements if rating.text.strip()]\n",
    "\n",
    "        print(f\"Alibaba Prices: {prices}\")  # Debug statement\n",
    "        print(f\"Alibaba Ratings: {ratings}\")  # Debug statement\n",
    "        \n",
    "        return prices, ratings\n",
    "    except Exception as e:\n",
    "        print(f\"Error retrieving prices from Alibaba: {str(e)}\")\n",
    "        return [], []\n",
    "\n",
    "def get_prices_alibaba(product_name):\n",
    "    driver.get(f\"https://www.alibaba.com/trade/search?fsb=y&IndexArea=product_en&CatId=&SearchText={product_name}\")\n",
    "\n",
    "    try:\n",
    "        price_elements = WebDriverWait(driver, 30).until(\n",
    "            EC.presence_of_all_elements_located((By.CSS_SELECTOR, '.search-card-e-price-main')))\n",
    "        rating_elements = WebDriverWait(driver, 30).until(\n",
    "            EC.presence_of_all_elements_located((By.CSS_SELECTOR, '.search-card-e-review')))\n",
    "        \n",
    "        prices = [price.text.strip() for price in price_elements if price.text.strip()]\n",
    "        ratings = [float(rating.text.strip().split()[0]) for rating in rating_elements if rating.text.strip()]\n",
    "\n",
    "        print(f\"Alibaba Prices: {prices}\")  # Debug statement\n",
    "        print(f\"Alibaba Ratings: {ratings}\")  # Debug statement\n",
    "        \n",
    "        return prices, ratings\n",
    "    except Exception as e:\n",
    "        print(f\"Error retrieving prices from Alibaba: {str(e)}\")\n",
    "        return [], []\n",
    "\n",
    "# Example product name (replace with your actual DataFrame)\n",
    "product_name = sorted_requisition.iloc[0]['Product Name'] \n",
    "\n",
    "# Collect prices and ratings from both vendors\n",
    "prices_ebay, ratings_ebay = get_prices_ebay(product_name)\n",
    "prices_alibaba, ratings_alibaba = get_prices_alibaba(product_name)\n",
    "\n",
    "# Integrity checks\n",
    "if not prices_ebay:\n",
    "    print(\"No prices retrieved from eBay.\")\n",
    "if not ratings_ebay:\n",
    "    print(\"No ratings retrieved from eBay.\")"
   ]
  },
  {
   "cell_type": "markdown",
   "id": "570d861e",
   "metadata": {},
   "source": [
    "# Compare the lowest price with the best rating"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": null,
   "id": "a753ef59-1358-4710-b449-fea4a32d7e5b",
   "metadata": {},
   "outputs": [],
   "source": [
    "# Combine prices, ratings, and create vendor list if there are prices\n",
    "if prices_ebay or prices_alibaba:\n",
    "    prices = prices_ebay + prices_alibaba\n",
    "    ratings = ratings_ebay + ratings_alibaba\n",
    "    vendors = ['eBay'] * len(prices_ebay) + ['Alibaba'] * len(prices_alibaba)\n",
    "\n",
    "    print(\"Number of vendors:\", len(vendors))\n",
    "    print(\"Number of prices:\", len(prices))\n",
    "\n",
    "    # Ensure all lists are the same length\n",
    "    min_length = min(len(prices), len(ratings), len(vendors))\n",
    "    prices = prices[:min_length]\n",
    "    ratings = ratings[:min_length]\n",
    "    vendors = vendors[:min_length]\n",
    "\n",
    "    comparison_df = pd.DataFrame({'Vendor': vendors, 'Price': prices, 'Rating': ratings})\n",
    "\n",
    "    # Clean and convert prices to numeric\n",
    "    def clean_price(price):\n",
    "        if \"Tap item to see current price\" in price or \"See price\" in price:\n",
    "            return None\n",
    "        \n",
    "        price = price.replace('US$', '').replace('$', '').replace(',', '').strip()\n",
    "\n",
    "        if 'to' in price:\n",
    "            price = price.split('to')[0].strip()\n",
    "        \n",
    "        if '-' in price:\n",
    "            price = price.split('-')[0].strip()\n",
    "\n",
    "        try:\n",
    "            return float(price)\n",
    "        except ValueError:\n",
    "            return None  # Return None if conversion fails\n",
    "\n",
    "    # Apply price cleaning\n",
    "    comparison_df['Price'] = comparison_df['Price'].apply(clean_price)\n",
    "    \n",
    "    # Drop rows with invalid prices\n",
    "    comparison_df = comparison_df.dropna(subset=['Price'])\n",
    "\n",
    "    if not comparison_df.empty:\n",
    "        # Save the comparison table to an Excel file\n",
    "        comparison_df.to_excel('comparison_table.xlsx', index=False)\n",
    "\n",
    "        # Step 1: Get the highest rating\n",
    "        max_rating = comparison_df['Rating'].max()\n",
    "\n",
    "        # Step 2: Filter the DataFrame to only include rows with the highest rating\n",
    "        best_rated_df = comparison_df[comparison_df['Rating'] == max_rating]\n",
    "\n",
    "        # Step 3: From the highest-rated rows, find the one with the lowest price\n",
    "        best_combined_row = best_rated_df.loc[best_rated_df['Price'].idxmin()]\n",
    "\n",
    "        best_combined_price = best_combined_row['Price']\n",
    "        best_combined_vendor = best_combined_row['Vendor']\n",
    "        best_combined_rating = best_combined_row['Rating']\n",
    "\n",
    "        # Generate minutes for the purchase committee based on best combination\n",
    "        minutes = f\"\"\"\n",
    "\n",
    "[ Product : {product_name} ]\n",
    "[ Price Comparison ]\n",
    "{comparison_df.to_string(index=False)}\n",
    "\n",
    "[ Lowest Price with Best Rating (5.0★) ]\n",
    "Vendor: {best_combined_vendor}\n",
    "Price: ${best_combined_price:.2f}\n",
    "Rating: {best_combined_rating}\n",
    "\"\"\"\n",
    "\n",
    "        print(minutes)\n",
    "    else:\n",
    "        print(\"No valid price data to display.\")\n",
    "else:\n",
    "    print(\"No data to save.\")\n",
    "\n",
    "driver.quit()"
   ]
  },
  {
   "cell_type": "markdown",
   "id": "2110934a",
   "metadata": {},
   "source": [
    "# Using Beautiful Soup"
   ]
  },
  {
   "cell_type": "markdown",
   "id": "c34a1cba-0eb5-47ed-a603-b3f17d1efcc0",
   "metadata": {},
   "source": [
    "# Install require package"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 1,
   "id": "239292c8-310d-418c-8f58-d83dadbc36c8",
   "metadata": {},
   "outputs": [
    {
     "name": "stdout",
     "output_type": "stream",
     "text": [
      "Requirement already satisfied: pandas in c:\\users\\user\\anaconda3\\lib\\site-packages (2.2.2)\n",
      "Requirement already satisfied: requests in c:\\users\\user\\anaconda3\\lib\\site-packages (2.32.2)\n",
      "Requirement already satisfied: beautifulsoup4 in c:\\users\\user\\anaconda3\\lib\\site-packages (4.12.3)\n",
      "Requirement already satisfied: openpyxl in c:\\users\\user\\anaconda3\\lib\\site-packages (3.1.2)\n",
      "Requirement already satisfied: fpdf in c:\\users\\user\\anaconda3\\lib\\site-packages (1.7.2)\n",
      "Requirement already satisfied: numpy>=1.26.0 in c:\\users\\user\\anaconda3\\lib\\site-packages (from pandas) (1.26.4)\n",
      "Requirement already satisfied: python-dateutil>=2.8.2 in c:\\users\\user\\anaconda3\\lib\\site-packages (from pandas) (2.9.0.post0)\n",
      "Requirement already satisfied: pytz>=2020.1 in c:\\users\\user\\anaconda3\\lib\\site-packages (from pandas) (2024.1)\n",
      "Requirement already satisfied: tzdata>=2022.7 in c:\\users\\user\\anaconda3\\lib\\site-packages (from pandas) (2023.3)\n",
      "Requirement already satisfied: charset-normalizer<4,>=2 in c:\\users\\user\\anaconda3\\lib\\site-packages (from requests) (2.0.4)\n",
      "Requirement already satisfied: idna<4,>=2.5 in c:\\users\\user\\anaconda3\\lib\\site-packages (from requests) (3.7)\n",
      "Requirement already satisfied: urllib3<3,>=1.21.1 in c:\\users\\user\\anaconda3\\lib\\site-packages (from requests) (2.2.2)\n",
      "Requirement already satisfied: certifi>=2017.4.17 in c:\\users\\user\\anaconda3\\lib\\site-packages (from requests) (2024.6.2)\n",
      "Requirement already satisfied: soupsieve>1.2 in c:\\users\\user\\anaconda3\\lib\\site-packages (from beautifulsoup4) (2.5)\n",
      "Requirement already satisfied: et-xmlfile in c:\\users\\user\\anaconda3\\lib\\site-packages (from openpyxl) (1.1.0)\n",
      "Requirement already satisfied: six>=1.5 in c:\\users\\user\\anaconda3\\lib\\site-packages (from python-dateutil>=2.8.2->pandas) (1.16.0)\n",
      "Note: you may need to restart the kernel to use updated packages.\n"
     ]
    }
   ],
   "source": [
    "pip install pandas requests beautifulsoup4 openpyxl fpdf"
   ]
  },
  {
   "cell_type": "markdown",
   "id": "bcec591b-a220-4abb-9f83-4c945fe38a10",
   "metadata": {},
   "source": [
    "# Read the requisition form"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 25,
   "id": "6fe95ec6-3e5e-4aaf-9c92-e6c0dcdb3c06",
   "metadata": {},
   "outputs": [],
   "source": [
    "import streamlit as st\n",
    "import pandas as pd\n",
    "\n",
    "# Title of the app\n",
    "st.title(\"RPA Assignment\")\n",
    "\n",
    "# File uploader\n",
    "uploaded_file = st.file_uploader(\"Choose an Excel file\", type=[\"xlsx\", \"xls\"])\n",
    "\n",
    "if uploaded_file is not None:\n",
    "    # Read the Excel file\n",
    "    try:\n",
    "        df = pd.read_excel(uploaded_file)\n",
    "\n",
    "        # Check if 'Required Date' column exists\n",
    "        if 'Required Date' in df.columns:\n",
    "            # Convert 'Required Date' to datetime\n",
    "            df['Required Date'] = pd.to_datetime(df['Required Date'])\n",
    "\n",
    "            # Sort the DataFrame by 'Required Date'\n",
    "            sorted_requisition = df.sort_values(by='Required Date')\n",
    "            st.write(\"Sorted Requisition Data by Date:\")\n",
    "            st.dataframe(sorted_requisition)\n",
    "        else:\n",
    "            st.error(\"The 'Required Date' column is not found in the uploaded file.\")\n",
    "\n",
    "    except Exception as e:\n",
    "        st.error(f\"Error reading the file: {e}\")"
   ]
  },
  {
   "cell_type": "markdown",
   "id": "da1a3e85-d0d5-4f9f-80e5-444c72ea098b",
   "metadata": {},
   "source": [
    "# Retrieve data"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 42,
   "id": "93ea3b06-a55b-4b05-bdb4-b732bb02ff06",
   "metadata": {},
   "outputs": [],
   "source": [
    "import requests\n",
    "from bs4 import BeautifulSoup\n",
    "\n",
    "product_name = sorted_requisition.iloc[0]['Product Name'] \n",
    "\n",
    "def get_prices_ebay(product_name):\n",
    "    url = f\"https://www.ebay.com/sch/i.html?_nkw={product_name}\"\n",
    "    response = requests.get(url)\n",
    "    soup = BeautifulSoup(response.text, 'html.parser')\n",
    "\n",
    "    prices = []\n",
    "    ratings = []\n",
    "    captions = []\n",
    "    sources = []\n",
    "    \n",
    "    items = soup.select('.s-item')\n",
    "    for item in items:\n",
    "        price = item.select_one('.s-item__price').text.strip()\n",
    "        prices.append(price)\n",
    "\n",
    "        rating_element = item.select_one('.s-item__reviews span.clipped')\n",
    "        if rating_element:\n",
    "            rating_text = rating_element.text.strip()\n",
    "            numeric_rating = float(rating_text.split()[0])\n",
    "            ratings.append(numeric_rating)\n",
    "        else:\n",
    "            ratings.append(0.0)\n",
    "\n",
    "        caption_element = item.select_one('.s-item__title')\n",
    "        if caption_element:\n",
    "            caption = caption_element.text.strip()\n",
    "            captions.append(caption)\n",
    "        else:\n",
    "            captions.append(\"No caption available\")\n",
    "            \n",
    "        sources.append(\"Ebay\")\n",
    "    return prices, ratings, captions, sources"
   ]
  },
  {
   "cell_type": "markdown",
   "id": "eea01046-2b21-4471-96cf-2cd14da88d43",
   "metadata": {},
   "source": [
    "# Convert retrieved rating to stars"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 43,
   "id": "6a64e22e-bbb9-47e2-956c-442a462c5dcc",
   "metadata": {},
   "outputs": [],
   "source": [
    "def rating_to_stars(rating):\n",
    "    full_stars = int(rating)  # Number of full stars\n",
    "    if rating % 1 >= 0.75:\n",
    "        return '★' * (full_stars + 1) + '☆' * (5 - full_stars - 1)  # Full star for >= 0.75\n",
    "    elif rating % 1 >= 0.25:\n",
    "        return '★' * full_stars + '½' + '☆' * (5 - full_stars - 1)  # Half star for 0.25 to < 0.75\n",
    "    else:\n",
    "        return '★' * full_stars + '☆' * (5 - full_stars)  # No half star for < 0.25"
   ]
  },
  {
   "cell_type": "markdown",
   "id": "c0955aae-09d1-418e-98c9-77e4b20d0e96",
   "metadata": {},
   "source": [
    "# Call the functin with product name"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 44,
   "id": "1d9c54d0-0fb0-4a5e-a452-d6cd3bbedbca",
   "metadata": {},
   "outputs": [],
   "source": [
    "prices_ebay, ratings_ebay, captions_ebay, sources_ebay = get_prices_ebay(product_name)"
   ]
  },
  {
   "cell_type": "markdown",
   "id": "2ef7f39f-a529-42bd-ba83-30921e1b1f58",
   "metadata": {},
   "source": [
    "# Data Frame"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 45,
   "id": "2377c4f0-7cdd-4928-8e09-f0bbae5df123",
   "metadata": {},
   "outputs": [],
   "source": [
    "data = {\n",
    "    'Sources': sources_ebay,\n",
    "    'Caption': captions_ebay,\n",
    "    'Price': prices_ebay,\n",
    "    'Rating': [rating_to_stars(r) for r in ratings_ebay]\n",
    "}"
   ]
  },
  {
   "cell_type": "markdown",
   "id": "107c01f8-55f0-4c10-807a-5fcbc683d3a2",
   "metadata": {},
   "source": [
    "# Display the data frame in table format"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 46,
   "id": "e3196540-b259-4467-83f7-f9ebad1574f9",
   "metadata": {},
   "outputs": [
    {
     "name": "stdout",
     "output_type": "stream",
     "text": [
      "\n",
      "Product: External Hard Drive\n",
      "\n",
      "Sources                                                                                     Caption             Price Rating\n",
      "   Ebay                                                                                Shop on eBay            $20.00  ☆☆☆☆☆\n",
      "   Ebay                                                                                Shop on eBay            $20.00  ☆☆☆☆☆\n",
      "   Ebay            WD 5TB Certified Refurbished Elements, External Hard Drive - RWDBU6Y0050BBK-WESN            $89.99  ☆☆☆☆☆\n",
      "   Ebay            WD 6TB Certified Refurbished Elements, External Hard Drive - RWDBWLG0060HBK-NESN            $99.99  ☆☆☆☆☆\n",
      "   Ebay                        WD 14TB Certified Refurbished My Book, Desktop External Hard Drive -           $189.99  ☆☆☆☆☆\n",
      "   Ebay            WD 3TB Certified Refurbished My Book, Desktop External HDD - RWDBBGB0030HBK-NESN            $49.99  ☆☆☆☆☆\n",
      "   Ebay                 BRAND NEW WD Black P10 5TB Portable External Game Hard Drive for ps5 SEALED            $51.00  ☆☆☆☆☆\n",
      "   Ebay                BRAND NEW WD Black P10 5TB Portable External Game Hard Drive for ps5 ps  box            $48.00  ☆☆☆☆☆\n",
      "   Ebay                Toshiba Canvio Advance Plus - 4TB External Hard Drive  USB 3.2 Gen 1 - Green            $37.00  ☆☆☆☆☆\n",
      "   Ebay                    Toshiba Canvio Advance Plus - 4TB External Hard Drive  USB 3.2 Gen 1 NEW            $35.00  ☆☆☆☆☆\n",
      "   Ebay               WD External Hard Drive My Passport 2 TB Password Protection Stockage Portable            $15.50  ☆☆☆☆☆\n",
      "   Ebay                     Avolusion 1 TB, External, 2.5\" (HD250U3-X1-1TB-XBOX) Hard Drive - Black            $10.50  ☆☆☆☆☆\n",
      "   Ebay                               Lacie 4TB HDD Rugged LRD0TU7 Thunderbolt USB-C Portable Drive            $15.00  ☆☆☆☆☆\n",
      "   Ebay     New ListingBRAND NEW WD Black P10 5TB Portable External Game Hard Drive for ps5 ps  box            $48.00  ☆☆☆☆☆\n",
      "   Ebay            WD 2TB Certified Refurbished Elements, External Hard Drive - RWDBU6Y0020BBK-WESN            $49.99  ☆☆☆☆☆\n",
      "   Ebay            WD 2TB Certified Refurbished Elements, External Hard Drive - RWDBUZG0020BBK-WESN            $49.99  ☆☆☆☆☆\n",
      "   Ebay              External Hard Drive Case 2TB USB 3.0 Portable Disk Enclosure 2.5\" HDD Sata SSD            $12.99  ☆☆☆☆☆\n",
      "   Ebay     New ListingToshiba Canvio Advance Plus - 4TB External Hard Drive  USB 3.2 Gen 1 - Green            $37.00  ★★★★½\n",
      "   Ebay            Open Box: WD_BLACK 2TB P10 Game Drive - Portable External Hard Drive HDD, Compat            $44.99  ★★★★★\n",
      "   Ebay            Cirago External Hard Drive USB 3.0 for PC, MacBook, Xbox One, Xbox 360, PS4, Mac  $16.99 to $58.99  ☆☆☆☆☆\n",
      "   Ebay                              Lacie Rugged RUFWU3B Fire Wire USB 3.0 1TB External Hard Drive            $21.99  ☆☆☆☆☆\n",
      "   Ebay                       WD 1TB Elements SE, Portable External Hard Drive - WDBEPK0010BBK-WESN            $54.99  ★★★★★\n",
      "   Ebay            WD 4TB Certified Refurbished Elements, External Hard Drive - RWDBU6Y0040BBK-WESN            $79.99  ☆☆☆☆☆\n",
      "   Ebay            ORICO 2/4/5 Bay USB 3.0 External Hard Drive Enclosure For 2.5/3.5\" SATA HDD 80TB $83.99 to $159.99  ☆☆☆☆☆\n",
      "   Ebay             WD 3TB Certified Refurbished Elements Portable External HDD RWDBU6Y0030BBK-WESN            $59.99  ☆☆☆☆☆\n",
      "   Ebay                    WD 5TB Certified Refurbished Elements SE, Portable External Hard Drive -            $89.99  ☆☆☆☆☆\n",
      "   Ebay                  Seagate Expansion 14TB Hard Drive External HDD - USB 3.0, w/ Data Recovery           $194.99  ★★★★★\n",
      "   Ebay                                       WD MY PASSPORT WDBYFT0020BBK-0B 2TB HARD DRIVE W/Cord            $20.00  ☆☆☆☆☆\n",
      "   Ebay             Seagate Expansion 500G 1TB 2TB External USB 3.0 Portable Hard Drive STEA1000400  $29.98 to $49.98  ☆☆☆☆☆\n",
      "   Ebay                         Seagate STKP14000400 Expansion 14TB External Hard Drive HDD USB 3.0           $194.99  ★★★★★\n",
      "   Ebay            HardDisk Extreme 500G 1T 2TB 256TB Portable USB 3.1 External Hard for Laptop PS5  $19.99 to $54.99  ☆☆☆☆☆\n",
      "   Ebay             Western Digital WDBABV5000ABK-00 500GB External Hard Drive/HDD | Tested Working            $19.89  ☆☆☆☆☆\n",
      "   Ebay             Toshiba Canvio Basics Portable External Hard Drive Black-500GB/1TB/2TB/4TB- NEW  $32.99 to $59.99  ☆☆☆☆☆\n",
      "   Ebay                 6TB EXTERNAL HARD DRIVE 2.5\" HDD BASIC STORAGE USB 3.0 *BRAND NEW OPEN BOX*            $44.99  ☆☆☆☆☆\n",
      "   Ebay                              Western Digital 8TB Certified Refurbished My Book Hard Drive -           $124.99  ☆☆☆☆☆\n",
      "   Ebay             WD My Passport 2TB Portable External Hard Disk Drive Storage HDD for Windows PC            $48.00  ★★★★★\n",
      "   Ebay               WD_BLACK 5TB P10 Game Drive for Xbox, Certified Refurbished Portable External            $99.99  ☆☆☆☆☆\n",
      "   Ebay             Avolusion HDDGEAR PRO X 8TB External Gaming Hard Drive for PS5 Game Console HDD           $109.99  ☆☆☆☆☆\n",
      "   Ebay            Avolusion PRO-Z Series 12TB USB 3.0 External Hard Drive for WindowsOS PC, Laptop           $129.99  ☆☆☆☆☆\n",
      "   Ebay            SanDisk Professional 2TB G-DRIVE ArmorATD External Hard Drive SDPH81G-002T-GBA1D           $104.99  ☆☆☆☆☆\n",
      "   Ebay               WD 2TB My Passport for Mac, Portable External Hard Drive - WDBA2D0020BBL-WESN            $79.99  ★★★★½\n",
      "   Ebay                 Western Digital WD WDBEPK5000ABK-WESN 500GB USB 3.0 External Hard Drive NEW            $39.99  ★★★★★\n",
      "   Ebay               Avolusion 1TB USB 3.0 Portable External Gaming Hard Drive for XBOX Series X|S            $48.99  ★★★★½\n",
      "   Ebay            Western Digital 2TB Certified Refurbished My Passport Ultra- RWDBC3C0020BSL-WESN            $59.99  ★★★★★\n",
      "   Ebay            New ListingWD 4TB My Passport - Portable External Hard Drive -Black- NEW SEALED!            $89.99  ★★★★★\n",
      "   Ebay                WD_BLACK 2TB Certified Refurbished P10 Game Drive, HDD - RWDBA2W0020BBK-WESN            $59.99  ☆☆☆☆☆\n",
      "   Ebay                WD 2TB My Passport, Portable External Hard Drive, White - WDBYVG0020BWT-WESN            $69.99  ☆☆☆☆☆\n",
      "   Ebay                WD 4TB My Passport, Portable External Hard Drive, Black - WDBPKJ0040BBK-WESN           $109.99  ★★★★★\n",
      "   Ebay                          WD 14TB Elements Desktop, External Hard Drive - WDBWLG0140HBK-NESN           $259.99  ★★★★★\n",
      "   Ebay             Avolusion 500GB USB 3.0 Portable External Gaming Hard Drive for XBOX Series X|S            $26.99  ☆☆☆☆☆\n",
      "   Ebay                                  Seagate 5 TB External Hard Drive, STGX5000400 ~ NEW SEALED            $99.99  ★★★★★\n",
      "   Ebay                Avolusion HD250U3-Z1-PRO 1TB USB 3.0 Portable External Gaming PS4 Hard Drive            $39.99  ☆☆☆☆☆\n",
      "   Ebay                            New ListingSeagate STKR6000400 Expansion 6TB External Hard Drive            $59.00  ★★★★★\n",
      "   Ebay                WD Elements 20TB USB 3.0 3.5\" Desktop External Hard Drive WDBWLG0200HBK-NESN           $295.00  ☆☆☆☆☆\n",
      "   Ebay                  WD NEW 1TB 5TB My Passport Portable External Hard Drive RED with Tracking# $85.87 to $182.54  ☆☆☆☆☆\n",
      "   Ebay               Hot External Hard Drive 2TB Capacity USB 3.0 450Mbps High Speed Plug And Play            $31.62  ☆☆☆☆☆\n",
      "   Ebay             WD My Passport 1TB Portable External Hard Disk Drive Storage HDD for Windows PC            $32.00  ★★★★★\n",
      "   Ebay               WD External Hard Drive My Passport 2 TB Password Protection Stockage Portable            $15.50  ☆☆☆☆☆\n",
      "   Ebay             USB 3.0 2TB SATA SSD External Hard Drive Portable Desktop Mobile Hard Disk Case            $13.05  ☆☆☆☆☆\n",
      "   Ebay              Western Digital 500GB WD5000LMVW-11VEDS0 2.5\" USB 3.0 External Hard Disk Drive            $24.50  ☆☆☆☆☆\n",
      "   Ebay                Western Digital My Book Hard Drive ENCLOSURE ONLY SEE BELOW wdbbgb0180hbk-nb            $21.99  ☆☆☆☆☆\n",
      "   Ebay             WD My Passport 2TB Portable External Hard Drive USB 3.2 Gen 1 WHITE + Tracking#           $107.22  ☆☆☆☆☆\n",
      "   Ebay                                                     LaCie d2 Quadra External 1TB Hard Drive            $25.00  ☆☆☆☆☆\n",
      "   Ebay    New ListingWD Western Digital Elements 2TB External Portable HDD WDBU6Y0020BBK-01 (READ)            $18.00  ☆☆☆☆☆\n",
      "   Ebay New ListingSeagate Portable 2N1AP5-500 2TB 2000GB USB 3.0 External Hard Drive HDD -NO POWER            $12.00  ★★★★★\n",
      "   Ebay                 LaCie 1TB, Rugged, Thunderbolt 2 / USB Micro-B, External Hard Drive LRD0TU1            $34.95  ☆☆☆☆☆\n",
      "   Ebay     New ListingSeagate - 2TB External USB 3.0 Portable Hard Drive with Rescue Data Recovery            $36.00  ★★★★★\n",
      "   Ebay                        WD Elements Desktop External Hard Drive 14TB USB 3.0 - WDBWLG0140HBK           $139.99  ☆☆☆☆☆\n",
      "   Ebay                           LaCie Rugged External Hard Drive HDD - 500 GB  - USB 3 - 100% --3            $29.00  ☆☆☆☆☆\n",
      "   Ebay                           LaCie Rugged External Hard Drive HDD - 500 GB  - USB 3 - 100% --4            $29.00  ☆☆☆☆☆\n",
      "   Ebay                                  Seagate External Hard Drive HDD - 2TB  - USB 3 - 100% --18            $49.00  ☆☆☆☆☆\n",
      "   Ebay                      Seagate 8TB External Desktop Drive USB 3.0 Black | Tested Works Great!            $50.00  ★★★★★\n",
      "   Ebay                          Western Digital External Hard Drive HDD - 2TB  - USB 3 - 100% --19            $49.00  ☆☆☆☆☆\n",
      "   Ebay            Seagate Backup Expansion Portable Drive 1.0 TB External Hard Drive SRD00F1 Black            $26.99  ☆☆☆☆☆\n",
      "   Ebay                                       New ListingSeagate Game Drive 2TB Portable Hard Drive            $12.99  ☆☆☆☆☆\n",
      "   Ebay            Seagate SRD0NF1 Expansion Portable 1TB External Hard Drive 2N1AP4-500 Minor Wear            $34.95  ☆☆☆☆☆\n",
      "   Ebay             New ListingThree 3 Terabyte Western Digital 3TB My Passport External Hard Drive            $38.00  ☆☆☆☆☆\n",
      "   Ebay                              LaCie Rugged External Hard Drive HDD - 2TB  - USB 3 - 100% --7            $49.00  ☆☆☆☆☆\n",
      "   Ebay                                  Seagate External Hard Drive HDD - 1TB  - USB 3 - 100% --10            $39.00  ☆☆☆☆☆\n",
      "   Ebay                          Western Digital External Hard Drive HDD - 2TB  - USB 3 - 100% --24            $49.00  ☆☆☆☆☆\n",
      "   Ebay                                     iomega 2TB External Hard Drive LDHD-UP2 31882300 Tested             $9.99  ☆☆☆☆☆\n",
      "   Ebay                                   LaCie Rugged Mini 2TB USB 3.0 External Hard Drive RUGU3M2            $44.95  ☆☆☆☆☆\n",
      "   Ebay            WD 2TB My Passport Portable External Hard Drive, Silicon Grey WDBWML0020BGY-WESN            $84.99  ☆☆☆☆☆\n",
      "   Ebay                       WD 44TB My Book Duo, Desktop External Hard Drive - WDBFBE0440JBK-NESN         $1,199.99  ☆☆☆☆☆\n",
      "\n",
      "Data has been saved to 'BScomparison_table(ebay).xlsx'.\n"
     ]
    }
   ],
   "source": [
    "# Create DataFrame\n",
    "results_df = pd.DataFrame(data)\n",
    "\n",
    "# Print product name and DataFrame\n",
    "print(f\"\\nProduct: {product_name}\\n\")\n",
    "print(results_df.to_string(index=False))\n",
    "\n",
    "# Save the DataFrame to an Excel file\n",
    "results_df.to_excel('BScomparison_table(ebay).xlsx', index=False)\n",
    "\n",
    "# Confirm save\n",
    "print(\"\\nData has been saved to 'BScomparison_table(ebay).xlsx'.\")"
   ]
  },
  {
   "cell_type": "markdown",
   "id": "f9928b15-f557-492e-9a1e-213ab0d41338",
   "metadata": {},
   "source": [
    "# Retrieve data from Alibaba"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 36,
   "id": "24f08221-b6c3-4f13-a470-3c325fcdcf4c",
   "metadata": {},
   "outputs": [],
   "source": [
    "import requests\n",
    "from bs4 import BeautifulSoup\n",
    "\n",
    "product_name = sorted_requisition.iloc[0]['Product Name'] \n",
    "\n",
    "def get_prices_alibaba(product_name):\n",
    "    url = f\"https://www.alibaba.com/trade/search?SearchText={product_name}\"\n",
    "    response = requests.get(url)\n",
    "    soup = BeautifulSoup(response.text, 'html.parser')\n",
    "\n",
    "    prices = []\n",
    "    ratings = []\n",
    "    captions = []\n",
    "    sources = []\n",
    "    \n",
    "    items = soup.select('.search-card')\n",
    "    for item in items:\n",
    "        price = item.select_one('.search-card-e-price-main').text.strip()\n",
    "        prices.append(price)\n",
    "\n",
    "        rating_element = item.select_one('.x-star-rating')\n",
    "        if rating_element:\n",
    "            rating_text = rating_element.text.strip()\n",
    "            numeric_rating = float(rating_text.split()[0])\n",
    "            ratings.append(numeric_rating)\n",
    "        else:\n",
    "            ratings.append(0.0)\n",
    "\n",
    "        caption_element = item.select_one('.search-e-card-title')\n",
    "        if caption_element:\n",
    "            caption = caption_element.text.strip()\n",
    "            captions.append(caption)\n",
    "        else:\n",
    "            captions.append(\"No caption available\")\n",
    "            \n",
    "        sources.append(\"Alibaba\")\n",
    "    \n",
    "    return prices, ratings, captions, sources"
   ]
  },
  {
   "cell_type": "markdown",
   "id": "9971d923-ae58-456a-8f44-fb227d7af60b",
   "metadata": {},
   "source": [
    "# Convert retrieved rating to stars"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 20,
   "id": "48784611-970f-4d2d-b535-d1a8e0db8a9c",
   "metadata": {},
   "outputs": [],
   "source": [
    "def rating_to_stars(rating):\n",
    "    full_stars = int(rating)\n",
    "    if rating % 1 >= 0.75:\n",
    "        return '★' * (full_stars + 1) + '☆' * (5 - full_stars - 1) \n",
    "    elif rating % 1 >= 0.25:\n",
    "        return '★' * full_stars + '½' + '☆' * (5 - full_stars - 1)  \n",
    "    else:\n",
    "        return '★' * full_stars + '☆' * (5 - full_stars)"
   ]
  },
  {
   "cell_type": "markdown",
   "id": "73cf3e98-2ca2-40ac-b360-757bfff324a7",
   "metadata": {},
   "source": [
    "# Call function with product name"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 37,
   "id": "28b8d167-bc1a-4c4d-897a-ee5d21eaa5bc",
   "metadata": {},
   "outputs": [],
   "source": [
    "prices_alibaba, ratings_alibaba, captions_alibaba, sources_alibaba= get_prices_alibaba(product_name)"
   ]
  },
  {
   "cell_type": "markdown",
   "id": "71984696-7629-40d6-846c-be43bf1fd8ef",
   "metadata": {},
   "source": [
    "# Data Frame"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 38,
   "id": "739f3fc3-5ec4-465b-b021-56161b3e03dd",
   "metadata": {},
   "outputs": [],
   "source": [
    "data = {\n",
    "    'Sources': sources_alibaba,\n",
    "    'Caption': captions_alibaba,\n",
    "    'Price': prices_alibaba,\n",
    "    'Rating': [rating_to_stars(r) for r in ratings_alibaba]\n",
    "}"
   ]
  },
  {
   "cell_type": "markdown",
   "id": "2bd727ce-c79e-4e63-b0a3-3a867f56430d",
   "metadata": {},
   "source": [
    "# Display the data frame in table format"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 39,
   "id": "1d0bf573-e996-4653-a363-b0ddb87e5ddc",
   "metadata": {
    "scrolled": true
   },
   "outputs": [
    {
     "name": "stdout",
     "output_type": "stream",
     "text": [
      "\n",
      "Product: External Hard Drive\n",
      "\n",
      "Empty DataFrame\n",
      "Columns: [Sources, Caption, Price, Rating]\n",
      "Index: []\n"
     ]
    }
   ],
   "source": [
    "# Create DataFrame\n",
    "results_df = pd.DataFrame(data)\n",
    "\n",
    "# Print product name and DataFrame\n",
    "print(f\"\\nProduct: {product_name}\\n\")\n",
    "print(results_df.to_string(index=False))\n",
    "\n",
    "# Save the DataFrame to an Excel file\n",
    "results_df.to_excel('BScomparison_table(alibaba).xlsx', index=False)"
   ]
  },
  {
   "cell_type": "markdown",
   "id": "79618f3c-3c55-4153-bd87-d765e2c2d6f3",
   "metadata": {},
   "source": [
    "# Compare the prices and ratings between the sources"
   ]
  },
  {
   "cell_type": "raw",
   "id": "b81142b8-4344-4fc1-a0a7-a656d2cc339b",
   "metadata": {},
   "source": [
    "Select the best products (Lowest price with highest rating)"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 7,
   "id": "b7b621f1-735c-44b5-804d-08fcac6e3e9d",
   "metadata": {},
   "outputs": [
    {
     "name": "stdout",
     "output_type": "stream",
     "text": [
      "\n",
      "Lowest Price with Highest Rating (5 stars):\n",
      "Caption: New ListingSeagate Portable 2N1AP5-500 2TB 2000GB USB 3.0 External Hard Drive HDD -NO POWER\n",
      "Price: $12.00\n",
      "Rating: ★★★★★\n",
      "Source: eBay\n"
     ]
    }
   ],
   "source": [
    "import pandas as pd\n",
    "\n",
    "alibaba_df = pd.read_excel('BScomparison_table(alibaba).xlsx')\n",
    "ebay_df = pd.read_excel('BScomparison_table(ebay).xlsx')\n",
    "\n",
    "alibaba_df['Source'] = 'Alibaba'\n",
    "ebay_df['Source'] = 'eBay'\n",
    "\n",
    "combined_df = pd.concat([alibaba_df, ebay_df], ignore_index=True)\n",
    "combined_df['Rating'] = combined_df['Rating'].astype(str).replace({'★★★★★': 5, '5.0': 5})\n",
    "\n",
    "def extract_lowest_price(price_str):\n",
    "    price_str = price_str.replace('$', '').replace(',', '')\n",
    "    prices = price_str.split(' to ')\n",
    "    return float(prices[0])  # Return the first price as float\n",
    "\n",
    "combined_df['Price'] = combined_df['Price'].apply(extract_lowest_price)\n",
    "filtered_df = combined_df[combined_df['Rating'] == 5]\n",
    "\n",
    "if not filtered_df.empty:\n",
    "    lowest_price_row = filtered_df.loc[filtered_df['Price'].idxmin()]\n",
    "    \n",
    "    def format_rating(rating):\n",
    "        return '★' * int(rating) + '☆' * (5 - int(rating))\n",
    "\n",
    "    caption = lowest_price_row['Caption']\n",
    "    price = f\"${lowest_price_row['Price']:.2f}\"  # Format price with dollar sign and two decimal places\n",
    "    rating = format_rating(lowest_price_row['Rating'])  # Convert rating to star format\n",
    "    source = lowest_price_row['Source']  # Get the source from the added column\n",
    "\n",
    "    print(\"\\nLowest Price with Highest Rating (5 stars):\")\n",
    "    print(f\"Caption: {caption}\")\n",
    "    print(f\"Price: {price}\")\n",
    "    print(f\"Rating: {rating}\")\n",
    "    print(f\"Source: {source}\")\n",
    "else:\n",
    "    print(\"No products with a 5-star rating found.\")"
   ]
  },
  {
   "cell_type": "markdown",
   "id": "2c7771e7-20ec-40e2-8f24-fb8c654d0a60",
   "metadata": {},
   "source": [
    "# Create a pdf to send to manager"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 56,
   "id": "250dc81a-290d-421b-bcb3-d9cc5e3cbf0b",
   "metadata": {},
   "outputs": [
    {
     "name": "stdout",
     "output_type": "stream",
     "text": [
      "Collecting reportlab\n",
      "  Downloading reportlab-4.2.2-py3-none-any.whl.metadata (1.4 kB)\n",
      "Requirement already satisfied: pillow>=9.0.0 in c:\\users\\user\\anaconda3\\lib\\site-packages (from reportlab) (10.3.0)\n",
      "Requirement already satisfied: chardet in c:\\users\\user\\anaconda3\\lib\\site-packages (from reportlab) (4.0.0)\n",
      "Downloading reportlab-4.2.2-py3-none-any.whl (1.9 MB)\n",
      "   ---------------------------------------- 0.0/1.9 MB ? eta -:--:--\n",
      "   ---------------------------------------- 0.0/1.9 MB ? eta -:--:--\n",
      "   ---------------------------------------- 0.0/1.9 MB ? eta -:--:--\n",
      "    --------------------------------------- 0.0/1.9 MB 330.3 kB/s eta 0:00:06\n",
      "   - -------------------------------------- 0.1/1.9 MB 469.7 kB/s eta 0:00:05\n",
      "   ----- ---------------------------------- 0.3/1.9 MB 1.8 MB/s eta 0:00:01\n",
      "   ----------------------- ---------------- 1.1/1.9 MB 5.1 MB/s eta 0:00:01\n",
      "   ---------------------------------------  1.9/1.9 MB 7.7 MB/s eta 0:00:01\n",
      "   ---------------------------------------- 1.9/1.9 MB 7.3 MB/s eta 0:00:00\n",
      "Installing collected packages: reportlab\n",
      "Successfully installed reportlab-4.2.2\n",
      "Note: you may need to restart the kernel to use updated packages.\n"
     ]
    }
   ],
   "source": [
    "pip install reportlab"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 11,
   "id": "6f8ce6c0-db0f-4dde-85eb-f67322eb3ffe",
   "metadata": {},
   "outputs": [
    {
     "name": "stdout",
     "output_type": "stream",
     "text": [
      "Unique Ratings: ['☆☆☆☆☆' '★★★★½' '★★★★★']\n",
      "Report generated successfully.\n"
     ]
    },
    {
     "name": "stderr",
     "output_type": "stream",
     "text": [
      "<>:13: SyntaxWarning: invalid escape sequence '\\$'\n",
      "<>:13: SyntaxWarning: invalid escape sequence '\\$'\n",
      "C:\\Users\\user\\AppData\\Local\\Temp\\ipykernel_3244\\642560357.py:13: SyntaxWarning: invalid escape sequence '\\$'\n",
      "  combined_df['Price'] = pd.to_numeric(combined_df['Price'].replace({'\\$': '', ',': ''}, regex=True), errors='coerce')\n"
     ]
    }
   ],
   "source": [
    "import pandas as pd\n",
    "from reportlab.lib.pagesizes import letter\n",
    "from reportlab.pdfgen import canvas\n",
    "\n",
    "# Function to read Excel files and find the product with the lowest price and highest rating\n",
    "def generate_report(file1, file2, output_pdf):\n",
    "    # Read the Excel files\n",
    "    df1 = pd.read_excel(file1)\n",
    "    df2 = pd.read_excel(file2)\n",
    "\n",
    "    combined_df = pd.concat([df1, df2], ignore_index=True)\n",
    "\n",
    "    combined_df['Price'] = pd.to_numeric(combined_df['Price'].replace({'\\$': '', ',': ''}, regex=True), errors='coerce')\n",
    "\n",
    "    unique_ratings = combined_df['Rating'].unique()\n",
    "    print(\"Unique Ratings:\", unique_ratings)\n",
    "\n",
    "    best_product = combined_df.loc[combined_df['Rating'] == '★★★★★'].nsmallest(1, 'Price')\n",
    "\n",
    "    if not best_product.empty:\n",
    "        source = best_product['Sources'].values[0]  # Get the source\n",
    "        product_name = best_product['Caption'].values[0]\n",
    "        lowest_price = best_product['Price'].values[0]\n",
    "        rating = best_product['Rating'].values[0]\n",
    "        \n",
    "        c = canvas.Canvas(output_pdf, pagesize=letter)\n",
    "        c.drawString(100, 750, \"Product Report\")\n",
    "        c.drawString(100, 670, f\"Source: {source}\")\n",
    "        c.drawString(100, 730, f\"Product Name: {product_name}\")\n",
    "        c.drawString(100, 710, f\"Lowest Price: ${lowest_price:.2f}\")\n",
    "        c.drawString(100, 690, f\"Rating: {rating}\")\n",
    "        c.save()\n",
    "        \n",
    "        print(f\"Report generated successfully.\")\n",
    "    else:\n",
    "        print(\"No products with a 5-star rating found.\")\n",
    "\n",
    "file1 = r'C:/Users/user/Downloads/Ai Assignment (1)-20240910T124230Z-001/Ai Assignment (1)/BScomparison_table(ebay).xlsx'  # First Excel file\n",
    "file2 = r'C:/Users/user/Downloads/Ai Assignment (1)-20240910T124230Z-001/Ai Assignment (1)/BScomparison_table(alibaba).xlsx'  # Second Excel file\n",
    "output_pdf = r'C:/Users/user/Downloads/Ai Assignment (1)-20240910T124230Z-001/Ai Assignment (1)/product_report.pdf'  # Output PDF file\n",
    "\n",
    "generate_report(file1, file2, output_pdf)"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": null,
   "id": "a0946a19-38c4-4a07-862f-f863bbd6237c",
   "metadata": {},
   "outputs": [],
   "source": []
  }
 ],
 "metadata": {
  "kernelspec": {
   "display_name": "Python 3 (ipykernel)",
   "language": "python",
   "name": "python3"
  },
  "language_info": {
   "codemirror_mode": {
    "name": "ipython",
    "version": 3
   },
   "file_extension": ".py",
   "mimetype": "text/x-python",
   "name": "python",
   "nbconvert_exporter": "python",
   "pygments_lexer": "ipython3",
   "version": "3.12.4"
  }
 },
 "nbformat": 4,
 "nbformat_minor": 5
}
