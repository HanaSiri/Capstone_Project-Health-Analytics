
import pandas as pd
import matplotlib
import matplotlib.pyplot as plt
import seaborn as sns
from textblob import TextBlob
import nltk
from nltk.sentiment.vader import SentimentIntensityAnalyzer
from wordcloud import WordCloud
import xlsxwriter

# Load data into a pandas DataFrame
file_path = r'C:\Users\siris\Downloads\Nutrition.xlsx'
df = pd.read_excel(file_path)

selected_columns = ['hashtag', 'viewsCount', 'likesCount', 'commentsCount', 'sharesCount', 'text']
df_selected = df[selected_columns]
print(df_selected)
with pd.ExcelWriter("Nutrition_cleaned.xlsx") as writer:
    df_selected.to_excel(writer)

views_summary = df_selected['viewsCount'].describe()
likes_summary = df_selected['likesCount'].describe()
comments_summary = df_selected['commentsCount'].describe()
shares_summary = df_selected['sharesCount'].describe()

print("Views Summary:\n", views_summary)
print("Likes Summary:\n", likes_summary)
print("Comments Summary:\n", comments_summary)
print("Shares Summary:\n", shares_summary)

# Aggregate data by 'tags'
tag_aggregated_data = df_selected.groupby('hashtag').agg({
    'viewsCount': 'sum',
    'likesCount': 'sum',
    'sharesCount': 'sum',
    'commentsCount': 'sum'
}).reset_index()

print(tag_aggregated_data)

tag_aggregated_data['hashtag'] = pd.Categorical(tag_aggregated_data['hashtag'])

# Comments Count
plt.figure(figsize=(12,10))
sns.barplot(x='hashtag', y='commentsCount', data=tag_aggregated_data)
plt.title('Comments Count by Tags')
plt.xticks(rotation=30)
plt.savefig('comments_count_by_tags.png')
plt.clf()
# Likes Count
plt.figure(figsize=(12,10))
sns.barplot(x='hashtag', y='likesCount', data=tag_aggregated_data)
plt.title('Likes Count by Tags')
plt.xticks(rotation=30)
plt.savefig('likes_count_by_tags.png')
plt.clf()

# Shares Count
plt.figure(figsize=(12,10))
sns.barplot(x='hashtag', y='sharesCount', data=tag_aggregated_data)
plt.title('Shares Count by Tags')
plt.xticks(rotation=30)
plt.savefig('shares_count_by_tags.png')
plt.clf()
# Views Count
plt.figure(figsize=(12,10))
sns.barplot(x='hashtag', y='viewsCount', data=tag_aggregated_data)
plt.title('Views Count by Tags')
plt.xticks(rotation=30)
plt.savefig('views_count_by_tags.png')
plt.clf()

text = " ".join(df_selected['text'].astype(str).tolist())
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
plt.savefig('wordcloud_Nutrition.jpg')
wordcloud_file = 'wordcloud_Nutrition.jpg'
wordcloud.to_file(wordcloud_file)

plt.clf()

with pd.ExcelWriter('Nutrition_Analysis.xlsx', engine='xlsxwriter') as writer:
    df_selected.to_excel(writer, sheet_name='Data', index=False)
    tag_aggregated_data.to_excel(writer, sheet_name='Aggregated Data', index=False)

    workbook = writer.book
    summary_sheet = workbook.add_worksheet('Summary')
    images_sheet = workbook.add_worksheet('Images')
    images1_sheet = workbook.add_worksheet('Wordcloud')




    # Write summary statistics to the Summary sheet
    summary_sheet.write('A1', 'Views Summary')
    for idx, value in enumerate(views_summary.index.values):
        summary_sheet.write(idx + 1, 0, value)
        summary_sheet.write(idx + 1, 1, views_summary[value])

    summary_sheet.write('A10', 'Likes Summary')
    for idx, value in enumerate(likes_summary.index.values):
        summary_sheet.write(idx + 11, 0, value)
        summary_sheet.write(idx + 11, 1, likes_summary[value])

    summary_sheet.write('A19', 'Comments Summary')
    for idx, value in enumerate(comments_summary.index.values):
        summary_sheet.write(idx + 20, 0, value)
        summary_sheet.write(idx + 20, 1, comments_summary[value])

    summary_sheet.write('A28', 'Shares Summary')
    for idx, value in enumerate(shares_summary.index.values):
        summary_sheet.write(idx + 29, 0, value)
        summary_sheet.write(idx + 29, 1, shares_summary[value])

    # Write sentiment analysis results to the Summary sheet
    summary_sheet.write('C1', 'Sentiment Analysis')
    summary_sheet.write('C2', 'VADER Sentiment')
    for idx, (key, value) in enumerate(sentiment.items()):
        summary_sheet.write(idx + 3, 2, key)
        summary_sheet.write(idx + 3, 3, value)

    summary_sheet.write('H1', 'TextBlob Sentiment')
    summary_sheet.write('H2', 'Polarity')
    summary_sheet.write('H3', sentiment1.polarity)
    summary_sheet.write('H4', 'Subjectivity')
    summary_sheet.write('H5', sentiment1.subjectivity)

    # Insert images into the Images sheet
    images_sheet.insert_image('A1', 'comments_count_by_tags.png')
    images_sheet.insert_image('A51', 'likes_count_by_tags.png')
    images_sheet.insert_image('A101', 'shares_count_by_tags.png')
    images_sheet.insert_image('A151', 'views_count_by_tags.png')
    images1_sheet.insert_image('A1', wordcloud_file)
