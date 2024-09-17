Data Extraction and NLP Text Analysis
This project involves extracting text from articles using web scraping techniques and performing text analysis to compute various linguistic variables such as sentiment scores, readability metrics, and more. The project was developed using Python and leverages libraries like BeautifulSoup, Selenium, and others to crawl the web and analyze the data.

Objective
The goal of this project is to extract textual data from provided URLs and analyze the text to compute a range of text-based variables including:

Positive Score
Negative Score
Polarity Score
Subjectivity Score
Average Sentence Length
Percentage of Complex Words
Gunning Fog Index
Average Number of Words Per Sentence
Complex Word Count
Word Count
Syllables Per Word
Personal Pronouns
Average Word Length
Features
Data Extraction: Scrapes articles from the URLs provided in an Excel sheet (input.xlsx). Only the article title and text are extracted, excluding irrelevant content like headers, footers, or advertisements.

Text Analysis: Performs sentiment analysis and computes various text-based metrics such as readability and sentence complexity.

Output: The results are saved in a structured format as per the Output Data Structure.xlsx.

Libraries Used
beautifulsoup4==4.12.3
blinker==1.8.2
distlib==0.3.8
httpx==0.26.0
nltk==3.9.1
openpyxl==3.1.5
packaging==23.2
pandas==2.2.1
referencing==0.33.0
requests==2.31.0
