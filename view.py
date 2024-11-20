import pandas as pd
import matplotlib.pyplot as plt
import openpyxl
from openpyxl.drawing.image import Image

# Load the Excel file into a DataFrame
file_path = 'feedback_analysis.xlsx'  # Adjust this to the path of your Excel file
df = pd.read_excel(file_path)

# Ensure the expected columns exist
required_columns = ['id', 'review', 'category', 'importance', 'sentiment']
if all(col in df.columns for col in required_columns):
    print("Data loaded successfully!")
else:
    print("Error: Missing one or more required columns.")

# 1. Count the number of reviews per category
category_counts = df['category'].value_counts()
print("\nReview count per category:")
print(category_counts)

# 2. Create a bar chart for the 3 sentiments
sentiment_counts = df['sentiment'].value_counts()

# Plotting the sentiment distribution
plt.figure(figsize=(8, 6))
sentiment_counts.plot(kind='bar', color=['green', 'blue', 'red'])
plt.title('Sentiment Distribution')
plt.xlabel('Sentiment')
plt.ylabel('Number of Reviews')
plt.xticks(rotation=0)
sentiment_chart_path = "sentiment_distribution.png"
plt.savefig(sentiment_chart_path)
plt.close()  # Close the plot to prevent it from showing

# 3. Rank the importance of categories
importance_counts = df['importance'].value_counts()

# Plotting importance ranking
plt.figure(figsize=(8, 6))
importance_counts.plot(kind='bar', color=['orange', 'yellow', 'purple', 'gray'])
plt.title('Importance of Categories')
plt.xlabel('Importance Level')
plt.ylabel('Number of Reviews')
plt.xticks(rotation=0)
importance_chart_path = "importance_ranking.png"
plt.savefig(importance_chart_path)
plt.close()  # Close the plot to prevent it from showing

# 4. Add summary data to the dataframe in separate columns to the right

# Sentiment summary (to add to the right of the existing data)
df['Sentiment Count'] = df['sentiment'].map(sentiment_counts.to_dict())

# Category summary (to add to the right of the existing data)
df['Category Count'] = df['category'].map(category_counts.to_dict())

# Importance summary (to add to the right of the existing data)
df['Importance Count'] = df['importance'].map(importance_counts.to_dict())

# 5. Save the updated DataFrame back to Excel with new table for sentiment/category/importance

# Open the Excel file
wb = openpyxl.load_workbook(file_path)
ws = wb.active

# Add summary data as a new table at the bottom or in a new sheet
summary_ws = wb.create_sheet('Summary')

# Write column headers for the summary table
summary_ws.append(['Sentiment', 'Category', 'Importance'])
summary_ws.append([str(sentiment_counts.to_dict())])
summary_ws.append([str(category_counts.to_dict())])
summary_ws.append([str(importance_counts.to_dict())])

# 6. Insert charts (sentiment and importance) into the Excel file
# Insert Sentiment Chart
sentiment_image = Image(sentiment_chart_path)
ws.add_image(sentiment_image, 'G2')  # You can change the cell position as needed

# Insert Importance Chart
importance_image = Image(importance_chart_path)
ws.add_image(importance_image, 'G20')  # You can change the cell position as needed

# Save the workbook with all changes
wb.save('updated_feedback_analysis_with_charts.xlsx')

print("\nUpdated Excel file saved as 'updated_feedback_analysis_with_charts.xlsx'.")
