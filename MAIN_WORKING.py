import os
import pandas as pd
import requests
from bs4 import BeautifulSoup
import nltk
from nltk.tokenize import word_tokenize, sent_tokenize
import re

# Download NLTK resources
nltk.download('punkt')

# Read input Excel file
input_file = "input.xlsx"
output_file = "output.xlsx"

# Read input Excel file
df_input = pd.read_excel(input_file)

# Function to extract article text from URL
def extract_article_text(url):
    try:
        response = requests.get(url)
        soup = BeautifulSoup(response.content, 'html.parser')
        # Find article content based on HTML structure of the website
        article_text = ""
        article_content = soup.find('div', class_='article-content')
        if article_content:
            # Extract text from paragraphs inside article content
            paragraphs = article_content.find_all('p')
            for paragraph in paragraphs:
                article_text += paragraph.text.strip() + "\n"
        else:
            # If no specific class or structure is found, extract all text
            article_text = soup.get_text(separator="\n")
        return article_text
    except Exception as e:
        print(f"Error extracting text from {url}: {str(e)}")
        return None

# Function to combine stop words from multiple files into a single set
def read_stop_words(folder_path):
    stop_words = set()
    for filename in os.listdir(folder_path):
        file_path = os.path.join(folder_path, filename)
        with open(file_path, "r", encoding="Latin-1") as f:
            words = f.read().splitlines()
            stop_words.update(words)
    return stop_words

# Read stop words from the provided folder
stop_words_folder = r"C:\Users\KIIT\Desktop\Pre-INTERNSHIP BlackCoffer\stopwords"
stop_words = read_stop_words(stop_words_folder)

# Read positive and negative words from text files
with open("positive-words.txt", "r") as f:
    positive_words = set(f.read().splitlines())

with open("negative-words.txt", "r") as f:
    negative_words = set(f.read().splitlines())

# Function to count syllables in a word
def count_syllables(word):
    word = word.lower()
    count = 0
    vowels = "aeiouy"
    if word[0] in vowels:
        count += 1
    for index in range(1, len(word)):
        if word[index] in vowels and word[index - 1] not in vowels:
            count += 1
    if word.endswith("e") and word[-2] not in vowels:
        count -= 1
    if word.endswith("es"):
        count -= 1
    if word.endswith("ed"):
        count -= 1
    return max(count, 1)

# Function to compute text analysis variables
def compute_text_variables(text):
    words = word_tokenize(text)
    sentences = sent_tokenize(text)
    
    word_count = len(words)
    sentence_count = len(sentences)
    avg_sentence_length = sum(len(sent.split()) for sent in sentences) / sentence_count
    
    # Calculate percentage of complex words
    complex_words = [word for word in words if word.lower() not in stop_words and len(word) > 6]
    percentage_of_complex_words = (len(complex_words) / word_count) * 100
    
    # Calculate Fog Index
    fog_index = 0.4 * (avg_sentence_length + percentage_of_complex_words)
    
    # Calculate Positive Score and Negative Score
    positive_score = sum(1 for word in words if word.lower() in positive_words)
    negative_score = sum(1 for word in words if word.lower() in negative_words)
    
    # Calculate Polarity Score
    polarity_score = (positive_score - negative_score) / ((positive_score + negative_score) + 0.000001)
    
    # Calculate Subjectivity Score
    subjectivity_score = (positive_score + negative_score) / (word_count + 0.000001)
    
    # Calculate Personal Pronouns count
    personal_pronouns = re.findall(r'\b(?:I|we|my|ours|us)\b', text, flags=re.IGNORECASE)
    personal_pronouns_count = len(personal_pronouns)
    
    # Calculate Syllables per Word
    syllables_per_word = sum(count_syllables(word) for word in words) / word_count
    
     # Calculate Average Word Length
    total_word_length = sum(len(word) for word in words)
    average_word_length = total_word_length / word_count if word_count > 0 else 0

    return {
        'WORD_COUNT': word_count,
        'SENTENCE_COUNT': sentence_count,
        'AVG_SENTENCE_LENGTH': avg_sentence_length,
        'PERCENTAGE_OF_COMPLEX_WORDS': percentage_of_complex_words,
        'FOG_INDEX': fog_index,
        'POSITIVE_SCORE': positive_score,
        'NEGATIVE_SCORE': negative_score,
        'POLARITY_SCORE': polarity_score,
        'SUBJECTIVITY_SCORE': subjectivity_score,
        'PERSONAL_PRONOUNS': personal_pronouns_count,
        'SYLLABLE_PER_WORD': syllables_per_word,
        'AVERAGE_WORD_LENGTH': average_word_length
    }

# Iterate over URLs, extract article text, and compute variables
output_data = []
for index, row in df_input.iterrows():
    url_id = row['URL_ID']
    url = row['URL']
    try:
        article_text = extract_article_text(url)
        if article_text:
            text_variables = compute_text_variables(article_text)
            text_variables['URL_ID'] = url_id
            output_data.append(text_variables)
    except Exception as e:
        print(f"Error processing URL {url_id}: {str(e)}")

# Create DataFrame from output_data
df_output = pd.DataFrame(output_data)

# Merge input and output DataFrames
df_merged = pd.merge(df_input, df_output, on='URL_ID')

# Save output to Excel file
df_merged.to_excel(output_file, index=False)
