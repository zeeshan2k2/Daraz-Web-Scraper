import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from bs4 import BeautifulSoup
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.chrome.service import Service
import re


# Function to scrape product information from Daraz website
def scrape_daraz(search_query):
    # Instantiate Chrome Options to run the browser in headless mode
    options = Options()
    options.headless = True

    # Start a new WebDriver session
    # driver = webdriver.Chrome(options=options)
    chromedriver_path = '/Users/zeeshanwaheed/Desktop/PycharmProjects1/Daraz Web Scraper/chromedriver'
    service = Service(chromedriver_path)
    driver = webdriver.Chrome(service=service)


    # Navigate to the Daraz search page
    url = f"https://www.daraz.pk/catalog/?q={search_query}&_keyori=ss&from=input&spm=a2a0e.searchlist.search.go.123679b4DeHmAn"
    driver.get(url)

    # Wait until the page is fully loaded
    try:
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "root")))
    except TimeoutException:
        print("Timeout occurred while waiting for the page to load")

    # Get the HTML content of the page
    html_content = driver.page_source
    driver.quit()

    # Parse the HTML content using BeautifulSoup
    soup = BeautifulSoup(html_content, "html.parser")

    # Find all div elements with class name "title-wrapper--IaQ0m" and ID "id-title"
    title_divs = soup.find_all("div", class_="title-wrapper--IaQ0m", id="id-title")

    # Find all div elements with class name "rating--ZI3Ol" and "rating--pwPrV"
    rating_spans = soup.find_all("span",
                                 class_=["ratig-num--KNake rating--pwPrV", "rating__review--ygkUy rating--pwPrV"])

    # Find all div elements with class name "current-price--Jklkc" and inside of it, find span with class "currency--GVKjl"
    price_spans = soup.find_all("div", class_="current-price--Jklkc")

    products = []

    # Extract product information
    for i in range(len(title_divs)):
        # Extract product name
        product_name = title_divs[i].text.strip() if title_divs else "N/A"

        # Extract product rating
        rating_text = rating_spans[i].text.strip() if rating_spans else "N/A"
        rating_value = float(re.sub(r'[^\d.]', '', rating_text))

        # Extract total reviews for the product
        total_reviews_text = rating_spans[i + 1].text.strip() if rating_spans else "N/A"
        total_reviews = int(re.sub(r'\D', '', total_reviews_text))

        # Extract product price
        price_text = price_spans[i].find("span", class_="currency--GVKjl").text.strip() if price_spans else "N/A"
        price = float(re.sub(r'[^\d.]', '', price_text))

        # Append product information to the products list
        products.append({"name": product_name, "rating": rating_value, "reviews": total_reviews, "price": price})

    return products


# Function to perform AI-based analysis on product data
def analyze_products(products):
    # For demonstration purposes, let's assume we're performing sentiment analysis on product names
    # This would require a real sentiment analysis model, but for now, let's just randomly assign sentiments
    sentiments = ['Positive', 'Neutral', 'Negative']
    for i, product in enumerate(products):
        product['sentiment'] = sentiments[i % 3]

    return products


# Main function
def main():
    # Define the search query
    search_query = "laptop"

    # Scrape product information from Daraz website
    products = scrape_daraz(search_query)

    # Perform AI-based analysis on product data
    analyzed_data = analyze_products(products)

    # Convert data to DataFrame
    df = pd.DataFrame(analyzed_data)

    # Format price column with commas for better readability
    df['price'] = df['price'].apply(lambda x: '{:,.2f}'.format(x))

    # Define the file path to save the Excel file
    file_path = r"/Users/zeeshanwaheed/Desktop/PycharmProjects1/Daraz Web Scraper/Daraz Insights.xlsx"

    # Save all product data to an Excel file
    df.to_excel(file_path, index=False)

    # Sort products based on rating and total reviews
    top_products = df.sort_values(by=["rating", "reviews"], ascending=[False, False]).head(5)

    # Save top products to a separate sheet in the Excel file
    with pd.ExcelWriter(file_path, mode='a', engine='openpyxl') as writer:
        top_products.to_excel(writer, sheet_name='Top Products', index=False)

    # Print a success message
    print("Data saved successfully.")


# Check if the script is being run directly
if __name__ == "__main__":
    main()