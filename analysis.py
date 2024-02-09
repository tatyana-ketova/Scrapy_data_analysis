import pandas as pd
import os
import nltk
from nltk.tokenize import word_tokenize
import ssl
import re
from nltk.tokenize import word_tokenize, sent_tokenize


nltk.data.path.append("nltk_data")

# Disable SSL certificate verification
try:
    _create_unverified_https_context = ssl._create_unverified_context
except AttributeError:
    pass
else:
    ssl._create_default_https_context = _create_unverified_https_context

# Now download the necessary NLTK resources
nltk.download('punkt', download_dir="nltk_data")



# Load stop words lists
stop_words_files = [
    'StopWords_Auditor.txt',
    'StopWords_Currencies.txt',
    'StopWords_DatesandNumbers.txt',
    'StopWords_Generic.txt',
    'StopWords_GenericLong.txt',
    'StopWords_Geographic.txt',
    'StopWords_Names.txt'
]
stop_words = set()
for file in stop_words_files:
    try:
        with open(os.path.join('StopWords', file), 'r', encoding='utf-8') as f:
            stop_words.update(f.read().splitlines())
    except UnicodeDecodeError:
        print("Failed to decode file with UTF-8 encoding. Trying alternative encodings...")

# Load master dictionary
positive_words = set()
negative_words = set()
try:
    with open(os.path.join('MasterDictionary', 'positive-words.txt'), 'r', encoding='utf-8') as f:
        positive_words.update(f.read().splitlines())
except UnicodeDecodeError:
    print("Failed to decode file with UTF-8 encoding. Trying alternative encodings...")
try:
    with open(os.path.join('MasterDictionary', 'negative-words.txt'), 'r', encoding='utf-8') as f:
        negative_words.update(f.read().splitlines())
except UnicodeDecodeError:
    print("Failed to decode file with UTF-8 encoding. Trying alternative encodings...")


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
    if word.endswith("e"):
        count -= 1
    if word.endswith("le"):
        count += 1
    if count == 0:
        count += 1
    return count


def clean_text(text):
    # Tokenize text
    tokens = word_tokenize(text.lower())
    # Remove stop words
    cleaned_tokens = [word for word in tokens if word not in stop_words]
    return cleaned_tokens


# Function to calculate readability scores
def calculate_readability_scores(text):
    sentences = sent_tokenize(text)
    words = [word for sentence in sentences for word in word_tokenize(sentence)]

    total_words = len(words)
    total_sentences = len(sentences)
    total_syllables = sum(count_syllables(word) for word in words)
    complex_words = [word for word in words if count_syllables(word) > 2]
    total_complex_words = len(complex_words)
    total_personal_pronouns = len(re.findall(r'\b(?:I|we|my|ours|us)\b', text, re.IGNORECASE))
    average_sentence_length = total_words / total_sentences
    percentage_complex_words = (total_complex_words / total_words) * 100
    fog_index = 0.4 * (average_sentence_length + percentage_complex_words)
    average_word_length = sum(len(word) for word in words) / total_words
    average_words_per_sentence = total_words / total_sentences

    return average_sentence_length, percentage_complex_words, fog_index, average_word_length, average_words_per_sentence, total_complex_words, total_syllables, total_personal_pronouns


def calculate_scores(tokens):
    positive_score = sum(1 for word in tokens if word in positive_words)
    negative_score = sum(1 for word in tokens if word in negative_words)
    polarity_score = (positive_score - negative_score) / (positive_score + negative_score + 0.000001)
    subjectivity_score = (positive_score + negative_score) / (len(tokens) + 0.000001)
    return positive_score, negative_score, polarity_score, subjectivity_score


def main():
    # Read input data structure
    output_structure_df = pd.read_excel('Output Data Structure.xlsx')

    output_data = []

    for index, row in output_structure_df.iterrows():
        url_id = row['URL_ID']
        url = row['URL']

        # Read extracted article text
        try:
            with open(os.path.join('extracted_articles', f"{url_id}.txt"), 'r', encoding='utf-8') as file:
                text = file.read()
        except FileNotFoundError:
            print(f"File '{url_id}.txt' not found in 'extracted_articles' folder.")
            continue  # Пропустить этот файл и перейти к следующему итерации цикла.

        # Calculate readability scores

        average_sentence_length, percentage_complex_words, fog_index, average_word_length, average_words_per_sentence, total_complex_words, total_syllables, total_personal_pronouns = calculate_readability_scores(
            text)

        # Calculate sentiment scores
        cleaned_tokens = clean_text(text)
        positive_score, negative_score, polarity_score, subjectivity_score = calculate_scores(cleaned_tokens)

        output_data.append({
            'URL_ID': url_id,
            'URL': url,
            'POSITIVE SCORE': positive_score,
            'NEGATIVE SCORE': negative_score,
            'POLARITY SCORE': polarity_score,
            'SUBJECTIVITY SCORE': subjectivity_score,
            'AVERAGE SENTENCE LENGTH': average_sentence_length,
            'PERCENTAGE OF COMPLEX WORDS': percentage_complex_words,
            'FOG INDEX': fog_index,
            'AVERAGE WORD LENGTH': average_word_length,
            'AVERAGE WORDS PER SENTENCE': average_words_per_sentence,
            'COMPLEX WORD COUNT': total_complex_words,
            'SYLLABLE COUNT PER WORD': total_syllables,
            'PERSONAL PRONOUNS COUNT': total_personal_pronouns
        })

    # Create DataFrame from output data and save to Excel
    output_df = pd.DataFrame(output_data)
    output_df.to_excel('output_results.xlsx', index=False)


if __name__ == "__main__":
    main()
