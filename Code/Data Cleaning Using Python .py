#!/usr/bin/env python
# coding: utf-8

# In[1]:


get_ipython().system('pip install pandas matplotlib seaborn')


# In[4]:


"""
E-Commerce Data Analytics Pipeline
Author: Abdul Rahman
Dataset: Global Superstore 2016
"""

import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

def main():
    print("Starting Data Processing Pipeline...")

    # ==========================================
    # STEP 1: Load the Data
    # ==========================================
    print("\n1. Loading datasets...")
    try:
        orders_df = pd.read_csv(r'C:\Users\Abdul Rahman\OneDrive\Desktop\SkillCraft Technology\Task 2\global_superstore_2016.xlsx - Orders.csv')
        returns_df = pd.read_csv(r'C:\Users\Abdul Rahman\OneDrive\Desktop\SkillCraft Technology\Task 2\global_superstore_2016.xlsx - Returns.csv')
        people_df = pd.read_csv(r'C:\Users\Abdul Rahman\OneDrive\Desktop\SkillCraft Technology\Task 2\global_superstore_2016.xlsx - People.csv')
        print("Data loaded successfully.")
    except FileNotFoundError as e:
        print(f"Error loading files: {e}")
        return

    # ==========================================
    # STEP 2: Merge the Datasets
    # ==========================================
    print("\n2. Merging datasets...")

    # Merge Orders with Returns (Left Join on 'Order ID' and 'Region')
    # This brings the 'Returned' column into the main orders dataframe
    master_df = pd.merge(orders_df, returns_df, on=['Order ID', 'Region'], how='left')

    # Merge the result with People (Left Join on 'Region')
    # This brings the 'Person' (Manager) column into the dataframe
    master_df = pd.merge(master_df, people_df, on='Region', how='left')

    print(f"Merged dataset shape: {master_df.shape}")

    # ==========================================
    # STEP 3: Data Cleaning
    # ==========================================
    print("\n3. Cleaning data...")

    # Fill NaN values in the 'Returned' column with 'No'
    master_df['Returned'] = master_df['Returned'].fillna('No')

    # Convert Date columns to actual datetime objects
    master_df['Order Date'] = pd.to_datetime(master_df['Order Date'])
    master_df['Ship Date'] = pd.to_datetime(master_df['Ship Date'])

    # Check for any remaining critical missing values
    print("Null values in key columns after cleaning:")
    print(master_df[['Returned', 'Person']].isnull().sum())

    # ==========================================
    # STEP 4: Basic Analysis & Export
    # ==========================================
    print("\n4. Performing Basic Analysis...")

    # Example Analysis 1: Total Sales by Region
    sales_by_region = master_df.groupby('Region')['Sales'].sum().sort_values(ascending=False)
    print("\nTop 5 Regions by Sales:")
    print(sales_by_region.head())

    # Example Analysis 2: Return Rate
    total_orders = len(master_df)
    returned_orders = len(master_df[master_df['Returned'] == 'Yes'])
    return_rate = (returned_orders / total_orders) * 100
    print(f"\nOverall Return Rate: {return_rate:.2f}%")

    # Save the cleaned and merged dataset for Power BI or further ML use
    output_filename = "Cleaned_Global_Superstore.csv"
    master_df.to_csv(output_filename, index=False)
    print(f"\nPipeline complete. Cleaned data saved as '{output_filename}'.")

if __name__ == "__main__":
    main()


# In[6]:


# ==========================================
    # STEP 1: Load the Data
    # ==========================================
    print("\n1. Loading datasets...")
    try:
        # Notice the 'r' before the quotes to handle Windows backslashes safely
        orders_df = pd.read_csv(r'C:\Users\Abdul Rahman\OneDrive\Desktop\SkillCraft Technology\Task 2\global_superstore_2016.xlsx - Orders.csv')

        returns_df = pd.read_csv(r'C:\Users\Abdul Rahman\OneDrive\Desktop\SkillCraft Technology\Task 2\global_superstore_2016.xlsx - Returns.csv')

        people_df = pd.read_csv(r'C:\Users\Abdul Rahman\OneDrive\Desktop\SkillCraft Technology\Task 2\global_superstore_2016.xlsx - People.csv')

        print("Data loaded successfully.")
    except FileNotFoundError as e:
        print(f"Error loading files: {e}")
        return



# In[7]:


# ==========================================
    # STEP 1: Load the Data (Excel File)
    # ==========================================
    print("\n1. Loading datasets...")
    try:
        # Define the path to your single Excel file
        excel_file = r'C:\Users\Abdul Rahman\OneDrive\Desktop\SkillCraft Technology\Task 2\global_superstore_2016.xlsx'

        # Read each sheet into its own dataframe
        orders_df = pd.read_excel(excel_file, sheet_name='Orders')
        returns_df = pd.read_excel(excel_file, sheet_name='Returns')
        people_df = pd.read_excel(excel_file, sheet_name='People')

        print("Data loaded successfully.")
    except FileNotFoundError as e:
        print(f"Error loading files: {e}")
        return


# In[8]:


get_ipython().system('pip install openpyxl')


# In[9]:


# ==========================================
    # STEP 1: Load the Data (Excel File)
    # ==========================================
    print("\n1. Loading datasets...")
    try:
        # Define the path to your single Excel file
        excel_file = r'C:\Users\Abdul Rahman\OneDrive\Desktop\SkillCraft Technology\Task 2\global_superstore_2016.xlsx'

        # Read each sheet into its own dataframe
        orders_df = pd.read_excel(excel_file, sheet_name='Orders')
        returns_df = pd.read_excel(excel_file, sheet_name='Returns')
        people_df = pd.read_excel(excel_file, sheet_name='People')

        print("Data loaded successfully.")
    except FileNotFoundError as e:
        print(f"Error loading files: {e}")
        return


# In[10]:


"""
E-Commerce Data Analytics Pipeline
Author: Abdul Rahman
Dataset: Global Superstore 2016
"""

import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

def main():
    print("Starting Data Processing Pipeline...")

    # ==========================================
    # STEP 1: Load the Data (Excel Sheets)
    # ==========================================
    print("\n1. Loading datasets...")
    try:
        # The raw path to your Excel file
        excel_file = r'C:\Users\Abdul Rahman\OneDrive\Desktop\SkillCraft Technology\Task 2\global_superstore_2016.xlsx'

        # Reading specific sheets from the Excel file
        orders_df = pd.read_excel(excel_file, sheet_name='Orders')
        returns_df = pd.read_excel(excel_file, sheet_name='Returns')
        people_df = pd.read_excel(excel_file, sheet_name='People')

        print("Data loaded successfully.")
    except FileNotFoundError as e:
        print(f"Error loading files: {e}")
        print("Please check that the file path and name are 100% correct.")
        return
    except Exception as e:
        print(f"An unexpected error occurred while loading: {e}")
        return

    # ==========================================
    # STEP 2: Merge the Datasets
    # ==========================================
    print("\n2. Merging datasets...")

    # Merge Orders with Returns (Left Join on 'Order ID' and 'Region')
    master_df = pd.merge(orders_df, returns_df, on=['Order ID', 'Region'], how='left')

    # Merge the result with People (Left Join on 'Region')
    master_df = pd.merge(master_df, people_df, on='Region', how='left')

    print(f"Merged dataset shape: {master_df.shape}")

    # ==========================================
    # STEP 3: Data Cleaning
    # ==========================================
    print("\n3. Cleaning data...")

    # Fill NaN values in the 'Returned' column with 'No'
    if 'Returned' in master_df.columns:
        master_df['Returned'] = master_df['Returned'].fillna('No')

    # Convert Date columns to actual datetime objects
    master_df['Order Date'] = pd.to_datetime(master_df['Order Date'])
    master_df['Ship Date'] = pd.to_datetime(master_df['Ship Date'])

    # Check for any remaining critical missing values
    print("\nNull values in key columns after cleaning:")
    print(master_df[['Returned', 'Person']].isnull().sum())

    # ==========================================
    # STEP 4: Basic Analysis & Export
    # ==========================================
    print("\n4. Performing Basic Analysis...")

    # Example Analysis 1: Total Sales by Region
    sales_by_region = master_df.groupby('Region')['Sales'].sum().sort_values(ascending=False)
    print("\nTop 5 Regions by Sales:")
    print(sales_by_region.head())

    # Example Analysis 2: Return Rate
    total_orders = len(master_df)
    if 'Returned' in master_df.columns:
        returned_orders = len(master_df[master_df['Returned'] == 'Yes'])
        return_rate = (returned_orders / total_orders) * 100
        print(f"\nOverall Return Rate: {return_rate:.2f}%")

    # Save the cleaned and merged dataset as a CSV for easy use in Power BI or future ML models
    output_filename = "Cleaned_Global_Superstore.csv"
    master_df.to_csv(output_filename, index=False)
    print(f"\nPipeline complete. Cleaned data saved as '{output_filename}' in your current directory.")

# This ensures the main function runs when you execute the script
if __name__ == "__main__":
    main()


# In[ ]:




