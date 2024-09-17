import os
import nltk
import re
import string
from nltk.corpus import stopwords
from nltk.tokenize import word_tokenize, sent_tokenize
import pandas as pd
from openpyxl import load_workbook

# Download NLTK data files
# nltk.download('stopwords')
# nltk.download('punkt')

# Function to load stop words
def load_stopwords(stop_words_folder):
    stop_words = set()
    for file_name in os.listdir(stop_words_folder):
        file_path = os.path.join(stop_words_folder, file_name)
        with open(file_path, 'r', encoding='ISO-8859-1') as file:
            for line in file:
                stop_words.add(line.strip().lower())
    return stop_words

# Function to load positive and negative words from Master Dictionary
def load_master_dictionary(master_dict_folder, stop_words):
    positive_words = set()
    negative_words = set()

    positive_file_path = os.path.join(master_dict_folder, 'positive-words.txt')
    negative_file_path = os.path.join(master_dict_folder, 'negative-words.txt')

    with open(positive_file_path, 'r', encoding='ISO-8859-1') as file:
        for line in file:
            word = line.strip().lower()
            if word not in stop_words:
                positive_words.add(word)

    with open(negative_file_path, 'r', encoding='ISO-8859-1') as file:
        for line in file:
            word = line.strip().lower()
            if word not in stop_words:
                negative_words.add(word)

    return positive_words, negative_words

# Function to calculate the derived variables
def calculate_scores(text, positive_words, negative_words):
    tokens = word_tokenize(text.lower())
    positive_score = sum(1 for word in tokens if word in positive_words)
    negative_score = sum(1 for word in tokens if word in negative_words)

    polarity_score = (positive_score - negative_score) / ((positive_score + negative_score) + 0.000001)
    subjectivity_score = (positive_score + negative_score) / (len(tokens) + 0.000001)

    return positive_score, negative_score, polarity_score, subjectivity_score

# Function to count syllables in a word
def count_syllables(word):
    word = word.lower()
    syllable_count = 0
    vowels = "aeiou"
    
    if word[0] in vowels:
        syllable_count += 1
    
    for index in range(1, len(word)):
        if word[index] in vowels and word[index-1] not in vowels:
            syllable_count += 1
    
    if word.endswith("es") or word.endswith("ed"):
        syllable_count -= 1
    
    if syllable_count == 0:
        syllable_count = 1
    
    return syllable_count

# Function to count complex words (more than 2 syllables)
def complex_word_count(text):
    words = word_tokenize(text)
    complex_words = [word for word in words if count_syllables(word) > 2]
    return len(complex_words)

# Function to calculate readability and related metrics
def analyze_readability(text, stop_words):
    # Tokenize sentences and words
    sentences = sent_tokenize(text)
    words = word_tokenize(text)
    
    # Remove stop words and punctuation
    words_cleaned = [word for word in words if word.lower() not in stop_words and word not in string.punctuation]
    
    # Count metrics
    num_sentences = len(sentences)
    num_words = len(words_cleaned)
    num_complex_words = complex_word_count(' '.join(words_cleaned))
    total_syllables = sum(count_syllables(word) for word in words_cleaned)
    total_characters = sum(len(word) for word in words_cleaned)
    
    # Average Sentence Length
    avg_sent_len = num_words / num_sentences if num_sentences else 0
    
    # Percentage of Complex Words
    percentage_complex_words = num_complex_words / num_words if num_words else 0
    
    # Gunning Fog Index
    fog_index = 0.4 * (avg_sent_len + (percentage_complex_words * 100))
    
    # Average Number of Words Per Sentence
    avg_words_per_sentence = avg_sent_len
    
    # Word Count
    word_count = num_words
    
    # Syllable Count Per Word
    avg_syllables_per_word = total_syllables / num_words if num_words else 0
    
    # Average Word Length
    avg_word_length = total_characters / num_words if num_words else 0
    
    # Personal Pronouns
    personal_pronouns = re.findall(r'\b(I|we|my|ours|us)\b', text, re.IGNORECASE)
    num_personal_pronouns = len(personal_pronouns)
    
    return {
        'Average sentance length': avg_sent_len,
        'FOG index': fog_index,
        'Average Words Per Sentence': avg_words_per_sentence,
        'Complex Word Count': num_complex_words,
        'Number of complex word': num_complex_words,
        'Word Count': word_count,
        'Syllable Count Per Word': avg_syllables_per_word,
        'Average Word Length': avg_word_length,
        'Personal Pronouns Count': num_personal_pronouns,
        'Percentage of complex word': percentage_complex_words
    }

# Load stop words and master dictionary
stopwords_folder = 'StopWords'
master_dict_folder = 'MasterDictionary'
stop_words = load_stopwords(stopwords_folder)
positive_words, negative_words = load_master_dictionary(master_dict_folder, stop_words)

# Load the Excel workbook and select sheet
excel_file_path = 'C:/Users/khanh/Desktop/learn_j/Blackcoffer/Solution/Output Data Solution.xlsx'
workbook = load_workbook(excel_file_path)
sheet = workbook.active


# Load cleaned articles
cleaned_articles_folder = 'Cleaned_Articles'

start_row = 2  # Start from 2nd row

# Iterate over each cleaned article and calculate the scores
for idx, file_name in enumerate(os.listdir(cleaned_articles_folder), start=start_row):
    file_path = os.path.join(cleaned_articles_folder, file_name)
    
    # Read the cleaned article text
    with open(file_path, 'r', encoding='utf-8') as file:
        cleaned_text = file.read()
    
    # Calculate the derived variables
    positive_score, negative_score, polarity_score, subjectivity_score = calculate_scores(
        cleaned_text, positive_words, negative_words
    )
    
    # Calculate Readability
    readability_metrix = analyze_readability(cleaned_text, stop_words)
    
    # Write the data to the Excel sheet
    # 'A' column is for file names and 'B' for Links
    sheet[f'C{idx}'] = positive_score
    sheet[f'D{idx}'] = negative_score
    sheet[f'E{idx}'] = polarity_score
    sheet[f'F{idx}'] = subjectivity_score
    sheet[f'G{idx}'] = readability_metrix['Average sentance length']
    sheet[f'H{idx}'] = readability_metrix['Percentage of complex word']
    sheet[f'I{idx}'] = readability_metrix['FOG index']
    sheet[f'J{idx}'] = readability_metrix['Average Words Per Sentence']
    sheet[f'K{idx}'] = readability_metrix['Number of complex word']
    sheet[f'L{idx}'] = readability_metrix['Word Count']
    sheet[f'M{idx}'] = readability_metrix['Syllable Count Per Word']
    sheet[f'N{idx}'] = readability_metrix['Personal Pronouns Count']
    sheet[f'O{idx}'] = readability_metrix['Average Word Length']
    
    workbook.save(excel_file_path)

    print(f"Data has been written to the Excel file: {file_name}")


