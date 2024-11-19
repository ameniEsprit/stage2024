import requests
from bs4 import BeautifulSoup
import pandas as pd
import time
import random
from textblob import TextBlob

def scrape_aliexpress_electronics(url, num_pages):
    headers_list = [
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36",
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:89.0) Gecko/20100101 Firefox/89.0",
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36",
        "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.114 Safari/537.36",
    ]

    common_brands = [
        "Apple", "Samsung", "Sony", "Dell", "HP", "Lenovo", "Asus", "Acer",
        "Microsoft", "LG", "Huawei", "Xiaomi", "Nokia", "Google", "Amazon", 
        "Panasonic", "Toshiba", "Philips", "JBL", "Bose"
    ]

    products = []

    for page in range(1, num_pages + 1):
        page_url = f"{url}&page={page}"
        headers = {
            "User-Agent": random.choice(headers_list),
            "Accept-Language": "en-US,en;q=0.9",
            "Accept-Encoding": "gzip, deflate, br",
            "Connection": "keep-alive",
            "DNT": "1"
        }

        try:
            response = requests.get(page_url, headers=headers)
            response.raise_for_status()
            print(f"Page {page} response status: {response.status_code}")
        except requests.exceptions.RequestException as e:
            print(f"Failed to retrieve content from page {page}: {e}")
            continue

        soup = BeautifulSoup(response.content, "html.parser")
        print(f"Page {page} content length: {len(response.content)}")

        items = soup.select(".list-item")
        print(f"Found {len(items)} items on page {page}")

        for item in items:
            title = item.select_one(".item-title")
            price = item.select_one(".price-current")
            link = item.select_one(".item-title a")
            image = item.select_one(".pic a img")

            title_text = title.get_text().strip() if title else "No Title"
            price_text = price.get_text().strip() if price else "No Price"
            product_url = f"https://www.aliexpress.com{link['href']}" if link else "No Link"
            image_url = image['src'] if image else "No Image"

            if product_url == "No Link" or title_text == "No Title":
                print(f"Skipping item with missing title or link on page {page}")
                continue

            # Scrape product details including shipping, items sold, and rating
            try:
                product_response = requests.get(product_url, headers=headers)
                product_response.raise_for_status()
                product_soup = BeautifulSoup(product_response.content, "html.parser")

                shipping = product_soup.select_one(".dynamic-shipping-cost")
                items_sold = product_soup.select_one(".order-num")
                rating = product_soup.select_one(".overview-rating-average")

                shipping_text = shipping.get_text().strip() if shipping else "No Shipping Info"
                items_sold_text = items_sold.get_text().strip() if items_sold else "No Items Sold Info"
                rating_text = rating.get_text().strip() if rating else "No Rating Info"

                reviews = product_soup.select(".feedback-item .feedback-text")
                ratings = product_soup.select(".feedback-item .feedback-star")

                review_texts = [review.get_text().strip() for review in reviews]
                ratings_texts = [rating['title'].split()[0] for rating in ratings if 'title' in rating.attrs]

                # Calculate average rating
                average_rating = None
                if ratings_texts:
                    rating_values = [float(r) for r in ratings_texts]
                    if rating_values:
                        average_rating = sum(rating_values) / len(rating_values)

                # Analyze sentiment of reviews
                sentiments = [TextBlob(review).sentiment for review in review_texts]
                sentiment_polarities = [sentiment.polarity for sentiment in sentiments]
                sentiment_subjectivities = [sentiment.subjectivity for sentiment in sentiments]

                # Average sentiment analysis
                average_sentiment_polarity = sum(sentiment_polarities) / len(sentiment_polarities) if sentiment_polarities else 0
                average_sentiment_subjectivity = sum(sentiment_subjectivities) / len(sentiment_subjectivities) if sentiment_subjectivities else 0

            except requests.exceptions.RequestException as e:
                print(f"Failed to retrieve product details from {product_url}: {e}")
                shipping_text = "No Shipping Info"
                items_sold_text = "No Items Sold Info"
                rating_text = "No Rating Info"
                review_texts = []
                average_rating = None
                average_sentiment_polarity = 0
                average_sentiment_subjectivity = 0

            # Determine brand from title field
            detected_brand = "Unknown"
            for brand in common_brands:
                if brand.lower() in title_text.lower():
                    detected_brand = brand
                    break

            # Determine category based on keywords
            if "laptop" in title_text.lower() or "notebook" in title_text.lower() or "pc" in title_text.lower():
                category_text = "PC"
            elif "smartphone" in title_text.lower() or "phone" in title_text.lower():
                category_text = "Smartphone"
            elif "tablet" in title_text.lower():
                category_text = "Tablet"
            elif "headphone" in title_text.lower() or "earbud" in title_text.lower():
                category_text = "Headphones"
            elif "camera" in title_text.lower():
                category_text = "Camera"
            else:
                category_text = "Other"

            product_info = {
                "Title": title_text,
                "Price": price_text,
                "Brand": detected_brand,
                "Category": category_text,
                "Link": product_url,
                "Image": image_url,
                "Shipping": shipping_text,
                "Items Sold": items_sold_text,
                "Rating": rating_text,
                "Average Rating": average_rating,
                "Average Sentiment Polarity": average_sentiment_polarity,
                "Average Sentiment Subjectivity": average_sentiment_subjectivity
            }
            products.append(product_info)

        time.sleep(random.uniform(3, 5))  # Increased delay

    return products

# URL of the AliExpress electronics page
aliexpress_url = "https://www.aliexpress.com/category/200003482/electronics"

num_pages = 5  # Number of pages to scrape

# Retrieve the data
aliexpress_data = scrape_aliexpress_electronics(aliexpress_url, num_pages)

# Check the retrieved data
print(f"Total products scraped: {len(aliexpress_data)}")
print(aliexpress_data)

# Convert the results to DataFrame
df = pd.DataFrame(aliexpress_data)

# Ensure the file is not open and you have write permissions
output_file = "C:/Users/Amani/Desktop/stage 4eme/aliexpress_electronics.xlsx"
try:
    df.to_excel(output_file, index=False)
    print(f"Data has been successfully scraped and saved to {output_file}")
except PermissionError:
    print(f"Permission denied: Unable to save to {output_file}. Ensure the file is not open and you have write permissions.")
except Exception as e:
    print(f"An error occurred while saving the file: {e}")

