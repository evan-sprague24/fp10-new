import sqlite3
import pandas as pd
from openai import OpenAI
# Set your OpenAI API key directly when creating the client
client = OpenAI(api_key='enter key here')

# Define GPT model (you can change this to another model like gpt-4 if you want)
GPT_MODEL = "gpt-4"

def analyze_sentiment_and_categorize(comment):
    """
    Function to analyze sentiment, categorize problems, and rank their importance using OpenAI.
    """
    try:
        # Prepare the prompt for GPT to analyze sentiment and categorize issues
        messages = [
            {"role": "system", "content": "You are a helpful assistant."},
            {"role": "user", "content": f"""
            Analyze the following product review. First, determine whether the review is positive, neutral, or negative, and label it.
            Then, categorize any problems mentioned into a category such as: Product Quality, Delivery Issues, 
            Customer Service, etc. Finally, rank the importance of the issues as: Critical, High, Medium, Low.
            Please format your response as follows:
            
            Sentiment: [positive/neutral/negative]
            Problem Category: [Category]
            Importance: [Critical/High/Medium/Low]
            
            Review: {comment}
            """}
        ]
        
        # Call the OpenAI API for sentiment and categorization analysis using the correct format
        print("Calling OpenAI API for sentiment analysis...")  # Debugging log
        response = client.chat.completions.create(
            model=GPT_MODEL,  # Use the model like "gpt-4"
            messages=messages,
            temperature=0  # You can adjust this based on how creative you want the response to be
        )
        
        # Extract the raw result from OpenAI
        raw_response = response.choices[0].message.content.strip()  # Accessing content
        print(f"Raw response from OpenAI: {raw_response}")  # Print the raw response
        
        # Parse the raw response to extract the individual components
        sentiment = None
        category = None
        importance = None
        
        # Split the raw response into lines
        lines = raw_response.split("\n")
        for line in lines:
            if line.startswith("Sentiment:"):
                sentiment = line.split(":", 1)[1].strip()
            elif line.startswith("Problem Category:"):
                category = line.split(":", 1)[1].strip()
            elif line.startswith("Importance:"):
                importance = line.split(":", 1)[1].strip()
        
        return sentiment, category, importance
    
    except Exception as e:
        print(f"Error analyzing feedback: {e}")
        return None, None, None


def fetch_and_analyze_feedback(db_file='fix.db', output_file='feedback_analysis.xlsx'):
    """
    Fetches feedback from the SQLite database, analyzes sentiment, categorizes problems, 
    and saves the results to an Excel file.
    """
    try:
        # Connect to the SQLite database
        print("Connecting to the database...")  # Debugging log
        conn = sqlite3.connect(db_file)
        cursor = conn.cursor()
        
        # Query to select the id and comment from the feedback table
        query = "SELECT id, comment FROM feedback"
        print(f"Executing query: {query}")  # Debugging log
        cursor.execute(query)
        
        # Fetch all results
        rows = cursor.fetchall()
        
        # Debugging: print the fetched rows
        print(f"Fetched rows: {rows}")  # Debugging log
        
        # Check if rows were returned
        if not rows:
            print("No feedback found in the database.")
        else:
            # Initialize a list to store analysis results
            feedback_analysis = []
            
            # Iterate through each feedback and analyze it
            for row in rows:
                review_id = row[0]
                review_comment = row[1]
                
                # Perform sentiment analysis and categorization
                sentiment, category, importance = analyze_sentiment_and_categorize(review_comment)
                
                # Store the analysis result for this review
                analysis_result = {
                    'id': review_id,
                    'review': review_comment,
                    'category': category,
                    'importance': importance,
                    'sentiment': sentiment
                }
                
                feedback_analysis.append(analysis_result)
            
            # Convert the analysis results to a pandas DataFrame
            df = pd.DataFrame(feedback_analysis)
            
            # Write the DataFrame to an Excel file
            df.to_excel(output_file, index=False)
            print(f"Analysis results saved to {output_file}")
    
    except Exception as e:
        print(f"Error fetching and analyzing feedback: {e}")
    
    finally:
        # Close the database connection
        conn.close()


# Example usage: call the function to fetch and analyze feedback and save it to Excel
fetch_and_analyze_feedback()
