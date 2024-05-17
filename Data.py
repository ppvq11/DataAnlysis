import os
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns

# Ensure that the plots show in a window
plt.ion()

def load_data(file_path):
    """Load Excel data from the given file path"""
    try:
        return pd.read_excel(file_path)
    except Exception as e:
        print(f"Error loading Excel file: {e}")
        return None

def clean_data(df):
    """Basic data cleaning steps"""
    for col in df.columns:
        # Attempt to convert object columns to numeric, if possible
        if df[col].dtype == 'object':
            df[col] = pd.to_numeric(df[col], errors='coerce')
        
        # Fill numeric columns with the median
        if df[col].dtype in [np.number]:
            df[col].fillna(df[col].median(), inplace=True)
        else:
            # Handle non-numeric columns by filling with mode or a placeholder
            mode_val = df[col].mode().iloc[0] if not df[col].mode().empty else "Unknown"
            df[col] = df[col].fillna(mode_val)
            
    return df

def analyze_data(df):
    """Perform basic analysis and generate plots"""
    print("\nBasic Descriptive Statistics:")
    print(df.describe())
    
    # Plot histograms for all numerical data
    num_cols = df.select_dtypes(include=[np.number]).columns.tolist()
    if num_cols:
        df[num_cols].hist(figsize=(10, 10))
        plt.show()
    else:
        print("No numeric columns to plot.")

def save_clean_data(df, output_path):
    """Save the cleaned data to an Excel file"""
    try:
        df.to_excel(output_path, index=False)
        print(f"Cleaned data saved to {output_path}")
    except Exception as e:
        print(f"Error saving Excel file: {e}")

def main():
    # User input for file path
    file_path = input("Enter the path to the Excel file: ")
    if not os.path.exists(file_path):
        print("File does not exist.")
        return
    
    df = load_data(file_path)
    if df is not None:
        df = clean_data(df)
        analyze_data(df)
        # Save cleaned data
        output_path = 'cleaned_data.xlsx'
        save_clean_data(df, output_path)

if __name__ == "__main__":
    main()
