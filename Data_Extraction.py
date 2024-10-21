import requests
from bs4 import BeautifulSoup
import pandas as pd
import os

def extract_article(url):
    try:
        # Send a GET request to the URL
        response = requests.get(url)
        response.raise_for_status()  # Raise an exception for HTTP errors
    except requests.RequestException as e:
        print("Error fetching URL:", e)
        return None, None
    
    # Parse the HTML content of the page
    soup = BeautifulSoup(response.text, 'html.parser')
    
    # Find the title of the article
    title_tag = soup.find('title')
    if title_tag:
        title = title_tag.text.strip()
    else:
        title = "Untitled Article"
    
    # Find the main content of the article
    article_text = ""
    article_content = soup.find_all('p')
    for paragraph in article_content:
        article_text += paragraph.text.strip() + "\n"
    
    return title, article_text

def save_to_file(title, text, filename):
    try:
        with open(filename, "w", encoding="utf-8") as file:
            file.write(title + "\n\n")
            file.write(text)
        print("Article saved to", filename)
    except IOError as e:
        print("Error saving article:", e)

# Read URLs from the CSV file
df = pd.read_excel(r"C:\Users\KIIT\Desktop\Pre-INTERNSHIP BlackCoffer\input.xlsx")

# Create a directory to save text files if it doesn't exist
if not os.path.exists("text_files"):
    os.makedirs("text_files")

# Process each URL and save the extracted data to separate text files
for index, row in df.iterrows():
    url_id = row['URL_ID']
    url = row['URL']
    article_title, article_text = extract_article(url)
    if article_title and article_text:
        filename = f"text_files/blackassign{str(url_id).zfill(4)}.txt"
        save_to_file(article_title, article_text, filename)