# extracting Data from the given links in input.xlsx

import pandas as pd
import requests
from bs4 import BeautifulSoup
import os

# Read Input Excel file
input_file = r"C:\Users\khanh\Desktop\learn_j\Blackcoffer\Solution\Input.xlsx"
df = pd.read_excel(input_file)

# Creating Output dir
output_dir = "extracted_articles"
os.makedirs(output_dir, exist_ok=True)

# Itreare over each row of dataFrame
for index, row in df.iterrows():
    url_id = row['URL_ID']
    url = row['URL']
    
    try:
        # Fetch the web page
        response = requests.get(url)
        response.raise_for_status()
        
        # parse the HTML content
        soup = BeautifulSoup(response.text, 'html.parser')
        
        # Extract the article title...
        title = soup.find('title').get_text(strip=True) if soup.find('title') else 'No Title'
        
        # Extract the article text 
        article_text = ""
        for paragraph in soup.find_all('p'):
            article_text += paragraph.get_text(strip=True) + "\n"
        
        # Create the file name
        file_name = f"{url_id}.txt"
        file_path = os.path.join(output_dir, file_name)
        
        # Write the title and article text to the file
        with open(file_path, 'w', encoding='utf-8') as file:
            file.write(f"Title: {title}\n\n")
            file.write(article_text)
        
        print(f"Successfully extracted and saved article for {url_id} to {file_path}")
        
    except requests.RequestException as e:
        print(f"Failed to fetch {url}: {e}")
    except Exception as e:
        print(f"AN error occurred: {e}")
        
        
    