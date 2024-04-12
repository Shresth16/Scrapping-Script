import pandas as pd
import requests
from bs4 import BeautifulSoup
import csv


# Function to scrape data from a single URL
def scrape_data(url):
    try:
        response = requests.get(url)
        soup = BeautifulSoup(response.content, 'html.parser')

        # Extract title
        title = soup.title.text.strip() if soup.title else ""

        # Extract paragraph text
        paragraphs = soup.find_all('p')
        text = ' '.join([p.get_text().strip() for p in paragraphs])

        # Extract image URLs
        images = [img['src'] for img in soup.find_all('img', src=True)]

        return {'Title': title, 'Text': text, 'Images': images, 'URL': url}
    except Exception as e:
        print(f"Error scraping {url}: {e}")
        return None


# Function to read links from Excel and scrape data
def scrape_links_from_excel(excel_file):
    try:
        # Read links from Excel file (assuming they are in the first row)
        df = pd.read_excel(excel_file, header=None, nrows=1)
        links = df.values.flatten()

        # Scrape data from each link
        scraped_data = []
        for link in links:
            if isinstance(link, str):  # Check if it's a valid link
                scraped_data.append(scrape_data(link))

        return scraped_data
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return None


# Function to store scraped data in CSV file
def save_to_csv(data, csv_file):
    try:
        with open(csv_file, 'w', newline='', encoding='utf-8') as file:
            writer = csv.DictWriter(file, fieldnames=['Title', 'Text', 'Images', 'URL'])
            writer.writeheader()
            for item in data:
                writer.writerow(item)
        print(f"Scraped data saved to {csv_file}")
    except Exception as e:
        print(f"Error saving to CSV file: {e}")


# Main function
def main():
    excel_file = r'C:\\Users\\shres\\Downloads\\sd.xlsx'  # Provide the path to your Excel file
    csv_file = 'scraped_data2.csv'  # Output CSV file

    # Scrape data from links in Excel file
    scraped_data = scrape_links_from_excel(excel_file)

    if scraped_data:
        # Save scraped data to CSV file
        save_to_csv(scraped_data, csv_file)


if __name__ == "__main__":
    main()
