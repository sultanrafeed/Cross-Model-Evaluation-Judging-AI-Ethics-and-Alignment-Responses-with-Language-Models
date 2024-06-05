import pandas as pd
import re

# Load the Excel file
df = pd.read_excel('D:\code\Mistral Phi3 Ourdataset Untrained.xlsx')

# Function to extract numeric value from a cell
def extract_numeric_value(x):
    if isinstance(x, str):
        # Remove leading/trailing spaces and non-numeric characters
        match = re.search(r'\b[0-5]\b', x.strip())
        if match:
            return int(match.group(0))
    elif isinstance(x, (int, float)) and not pd.isna(x) and 0 <= x <= 5:
        return int(x)
    return None


# Function to map numeric values to labels
def map_to_label(x):
    if x == 0:
        return 'Unusable'
    elif x == 1:
        return 'Poor'
    elif x == 2:
        return 'Below Average'
    elif x == 3:
        return 'Average'
    elif x == 4:
        return 'Good'
    elif x == 5:
        return 'Excellent'
    else:
        return 'Not Valid'

# Extract numeric values from 'Output' column
df['Numeric Output'] = df['Output'].apply(extract_numeric_value)

# Apply the function to classify the numeric values
df['Classified as'] = df['Numeric Output'].apply(map_to_label)

# Filter out rows with 'Not Valid' in 'Classified as' column
df = df[df['Classified as'] != 'Not Valid']

# Count the number of each score
score_count = df['Classified as'].value_counts()
print("Score Counts:")
print(score_count)

# Calculate the average score
average_score = df['Numeric Output'].mean()
print("\nAverage Score:")
print(average_score)

# Save the updated DataFrame to a new Excel file
df.to_excel('D:\code\Mistral Phi3 Ourdataset Untrained scores.xlsx', index=False)
