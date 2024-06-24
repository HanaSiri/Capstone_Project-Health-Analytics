

import pandas as pd
import matplotlib
import matplotlib.pyplot as plt
import seaborn as sns
from textblob import TextBlob
import nltk
from nltk.sentiment.vader import SentimentIntensityAnalyzer
from wordcloud import WordCloud
import xlsxwriter
import re
from collections import Counter

file_path = r'D:\Python\pinterest_data_selenium_3.xlsx'
df_main = pd.read_excel(file_path)

print(df_main.columns)

hashtag_counts = df['Hashtag'].value_counts()
print(hashtag_counts)
# Create a DataFrame from counts
counts_df = pd.DataFrame(hashtag_counts).reset_index()
counts_df.columns = ['Hashtag', 'count']

# Merge counts back into original DataFrame
df_merge = pd.merge(df_main, counts_df, on='Hashtag')

print(df_merge.columns)

print(df_merge)

# Plotting
plt.figure(figsize=(12, 20))
plt.bar(df['Hashtag'], df['count'], color='skyblue')
plt.xlabel('Hashtag')
plt.ylabel('Count')
plt.title('Counts of Hashtags')
plt.xticks(rotation=90)

plt.savefig('Count_of_hashtags.png')
plt.clf()

def extract_hashtags(text):
    if pd.isna(text):
        return []
    hashtags = re.findall(r'#(\w+)', text)
    return hashtags
# Apply the function to extract hashtags from 'Title' column
df['Extracted_Hashtags'] = df['Title'].apply(extract_hashtags)

print(df[['Title', 'Extracted_Hashtags']])

df_merge['Extracted_Hashtags'] = df['Extracted_Hashtags']
print(df_merge.columns)

all_hashtags = [tag for tags in df['Extracted_Hashtags'] for tag in tags]
hashtag_counts = Counter(all_hashtags)

# Convert to DataFrame for plotting
counts_df = pd.DataFrame(hashtag_counts.items(), columns=['Extracted_Hashtags', 'Count'])

# Sort by Count descending for better visualization
counts_df = counts_df.sort_values(by='Count', ascending=False)

# Plotting
plt.figure(figsize=(30,30))
plt.bar(counts_df['Extracted_Hashtags'], counts_df['Count'], color='skyblue',width=0.8)
plt.xlabel('Extracted_Hashtags')
plt.ylabel('Count')
plt.title('Count of Extracted Hashtags')
plt.xticks(rotation=90)
plt.savefig('count_of_extracted_hashtags.png')
plt.clf()



print(counts_df)

text = " ".join(df['Title'].astype(str).tolist())
print(text)

nltk.download('vader_lexicon')
sia = SentimentIntensityAnalyzer()
sentiment = sia.polarity_scores(text)
print(sentiment)


blob = TextBlob(text)
sentiment1 = blob.sentiment
print(f"Polarity: {sentiment1.polarity}, Subjectivity: {sentiment1.subjectivity}")

wordcloud = WordCloud(width=800, height=400, background_color='white').generate(text)
# Display the generated word cloud image
plt.figure(figsize=(10, 5))
plt.imshow(wordcloud, interpolation='bilinear')
plt.axis('off')
plt.show()
plt.savefig('wordcloud_pinterest.jpg')
wordcloud_file = 'wordcloud_pinterest.jpg'
wordcloud.to_file(wordcloud_file)

plt.clf()

with pd.ExcelWriter('pinterest_analysis_results.xlsx', engine='xlsxwriter') as writer:
    df_main.to_excel(writer, sheet_name='Data', index=False)
    df_merge.to_excel(writer, sheet_name='Aggregated Data', index=False)

    workbook = writer.book
    summary_sheet = workbook.add_worksheet('Summary')
    images_sheet = workbook.add_worksheet('Images')
    images1_sheet = workbook.add_worksheet('Wordcloud')

    # Write sentiment analysis results to the Summary sheet
    summary_sheet.write('E1', 'Sentiment Analysis using VADER')
    summary_sheet.write('E2', f'Negative: {sentiment["neg"]}')
    summary_sheet.write('E3', f'Neutral: {sentiment["neu"]}')
    summary_sheet.write('E4', f'Positive: {sentiment["pos"]}')
    summary_sheet.write('E5', f'Compound: {sentiment["compound"]}')

    summary_sheet.write('H1', 'TextBlob Sentiment')
    summary_sheet.write('H2', 'Polarity')
    summary_sheet.write('H3', sentiment1.polarity)
    summary_sheet.write('H4', 'Subjectivity')
    summary_sheet.write('H5', sentiment1.subjectivity)
    counts_df.to_excel(writer,sheet_name='Summary')
    # Insert images into the Images sheet

    images_sheet.insert_image('A1', 'Count_of_hashtags.png')
    images_sheet.insert_image('A100', 'count_of_extracted_hashtags.png')
    images1_sheet.insert_image('A1', wordcloud_file)

