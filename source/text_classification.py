import pandas as pd
import numpy as np
import pandas as pd
import re
from sklearn.model_selection import train_test_split
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.linear_model import LogisticRegression
from sklearn.metrics import classification_report
import pickle
import os
import warnings
warnings.filterwarnings("ignore")

# Basic text cleaning function
def clean_text(text):
    if pd.isna(text):
        return "missing"
    text = re.sub(r'\W+', ' ', text)  # Remove special characters
    text = text.lower().strip()  # Convert to lowercase and strip whitespace
    return text

def clean_text(text):
    # Check if the input is a string before applying regex
    if not isinstance(text, str):
        return "missing"  # Return a default value for non-string input (e.g., NaN, numbers)

    # Now apply your cleaning logic
    text = re.sub(r'\W+', ' ', text)  # Remove special characters
    text = text.lower().strip()  # Convert to lowercase and strip whitespace
    return text

def extract_tables(file_path, sheet_name, check_names, check_idx, start_col, end_col):
    """
    Extract subtables from a sheet using pandas DataFrame.
    Subtables are separated by empty rows.
    Only subtables where the header in index 2 (column 3) contains 'Movimiento', 
    'Cta Cte USD', 'TC Nacional', or 'TC Internacional' will be added to the list of subtables.
    """

    # Read the Excel sheet into a DataFrame (only the first `col_num` columns)
    df = pd.read_excel(file_path, sheet_name=sheet_name, usecols=range(start_col, end_col), engine="openpyxl", header=None)
    
    # Replace all-NaN rows with a marker
    df["is_empty"] = df.isnull().all(axis=1)
    
    # Split the DataFrame into subtables based on empty rows
    subtables = []
    current_table = []

    for _, row in df.iterrows():
        if row["is_empty"]:
            # If we encounter an empty row, save the current table (if not empty)
            if current_table:
                # Convert current table to DataFrame
                table = pd.DataFrame(current_table).reset_index(drop=True)

                # Find the row where the column index 2 matches one of the check_names
                header_row_index = None
                for i, row in table.iterrows():
                    value = str(row[check_idx]).strip()
                    if value in check_names:
                        header_row_index = i
                        break

                if header_row_index is not None:
                    # Set the identified row as the header and drop all rows above it
                    table.columns = table.iloc[header_row_index]  # Set the header
                    table = table[header_row_index + 1:].reset_index(drop=True)  # Drop rows above the header

                    # Append the table if it's valid (it has content after setting the header)
                    if not table.empty:
                        subtables.append(table)

                current_table = []  # Reset for the next table
        else:
            # Add non-empty rows to the current table
            current_table.append(row[:-1])  # Exclude the "is_empty" column

    # Add the last table if present
    if current_table:
        # Convert current table to DataFrame
        table = pd.DataFrame(current_table).reset_index(drop=True)

        # Find the row where the column index 2 matches one of the check_names
        header_row_index = None
        for i, row in table.iterrows():
            value = str(row[check_idx]).strip()
            if value in check_names:
                header_row_index = i
                break

        if header_row_index is not None:
            # Set the identified row as the header and drop all rows above it
            table.columns = table.iloc[header_row_index]  # Set the header
            table = table[header_row_index + 1:].reset_index(drop=True)  # Drop rows above the header

            # Append the table if it's valid (it has content after setting the header)
            if not table.empty:
                subtables.append(table)

    # Remove subtables where the first column is only NaN
    subtables = [
        table for table in subtables 
        if not table.iloc[:, 0].isna().all()  # Check if the first column is NOT all NaN
    ]

    # Ensure unique column names for all subtables at once
    for idx, table in enumerate(subtables):
        # Use pandas' built-in method to ensure unique column names
        table.columns = [
            f"Unnamed_{i}" if pd.isna(col) else col for i, col in enumerate(table.columns)
        ]

    return subtables

def training_model(labeled_data, label_col_idx):
    # Split into training and test sets
    X_train, X_test, y_train, y_test = train_test_split(labeled_data['Combined_Text'], labeled_data.iloc[:, label_col_idx], test_size=0.2, random_state=42)
    # TF-IDF vectorization
    vectorizer = TfidfVectorizer(max_features=5000)  # Limit to top 5000 features
    X_train_tfidf = vectorizer.fit_transform(X_train)
    # Train the classifier
    model = LogisticRegression(max_iter=1000, random_state=42)
    model.fit(X_train_tfidf, y_train)
    return model, vectorizer, X_train, X_test, y_test

def classify_text(model, vectorizer, *args):
    # Combine and clean input text
    combined_text = ' '.join(clean_text(arg) for arg in args)
    text_tfidf = vectorizer.transform([combined_text])  # Transform using the trained vectorizer
    # Get predicted label and probabilities
    label = model.predict(text_tfidf)[0]  # Predicted label
    probabilities = model.predict_proba(text_tfidf)[0]  # Probabilities for all classes
    # Get the confidence score for the predicted label
    confidence_score = max(probabilities)  # Highest probability corresponds to the predicted label
    return label, round(confidence_score, 1)