import requests
from bs4 import BeautifulSoup
import pandas as pd
from urllib.parse import urlparse
import os


def extract_article_text(url):
    try:
        response = requests.get(url)
        soup = BeautifulSoup(response.content, 'html.parser')

        # Extracting title and main article text
        title = soup.title.text.strip()
        article_div = soup.find('div', class_='td-post-content')
        article_text = article_div.text.strip() if article_div else ''

        return title, article_text
    except Exception as e:
        print(f"Error extracting data from {url}: {e}")
        return None, None


def save_text_to_file(text, filename):
    with open(filename, 'w', encoding='utf-8') as file:
        file.write(text)


def main():
    input_file = 'input.xlsx'
    output_folder = 'extracted_articles'

    # Create output folder if it doesn't exist
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # Read URLs from input excel file
    df = pd.read_excel(input_file, header=None, names=['URL'])

    for index, row in df.iterrows():
        url_id = f"{index}"
        url = row['URL']

        # Extract data from URL
        title, article_text = extract_article_text(url)

        if title and article_text:
            # Save extracted text to a text file
            filename = os.path.join(output_folder, f"{url_id}.txt")
            save_text_to_file(f"{title}\n\n{article_text}", filename)
            print(f"Data extracted and saved for {url} as {filename}")


if __name__ == "__main__":
    main()