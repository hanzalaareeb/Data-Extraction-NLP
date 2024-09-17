import os
"""_summary_

Loading stop words from a folder containing multiple stop word files.

"""
def load_stop_words(stop_words_folder):
    stop_words = set()
    
    for file_name in os.listdir(stop_words_folder):
        file_path = os.path.join(stop_words_folder, file_name)
        with open(file_path, 'r', encoding='ISO-8859-1') as file:
            for line in file:
                stop_words.add(line.strip().lower())  # Add each stop word to the set
    
    return stop_words


# print(stop_words)


"""_summary_

Cleaning the article text by removing stop words from each article.

"""
def clean_text(text, stop_words):
    words = text.split()
    cleaned_text = ' '.join([word for word in words if word.lower() not in stop_words])
    return cleaned_text

# Load stopword 
stop_words_folder = "StopWords"
stop_words = load_stop_words(stop_words_folder)

# Path to the folder containing the articles
articles_folder = 'extracted_articles'

# Output folder for cleaned articles
cleaned_articles_folder = 'Cleaned_Articles'
os.makedirs(cleaned_articles_folder, exist_ok=True)

# Iterate over each file in the articles folder
for file_name in os.listdir(articles_folder):
    file_path = os.path.join(articles_folder, file_name)
    
    # Read the article text
    with open(file_path, 'r', encoding='utf-8') as file:
        article_text = file.read()
    
    # Clean the article text
    cleaned_text = clean_text(article_text, stop_words)
    
    # Save the cleaned text to a new file
    cleaned_file_path = os.path.join(cleaned_articles_folder, file_name)
    with open(cleaned_file_path, 'w', encoding='utf-8') as cleaned_file:
        cleaned_file.write(cleaned_text)
    
    print(f"Cleaned {file_name} and saved to {cleaned_file_path}")
