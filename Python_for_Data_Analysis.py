import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# Load the data from the Excel file
file_path = 'orderdataset.xlsx'
data = pd.read_excel(file_path)

# Split the single column into multiple columns if necessary
if data.shape[1] == 1:
    data = data.iloc[:, 0].str.split(';', expand=True)
    data.columns = [
        'order_id', 'quantity', 'product_id', 'price', 'seller_id',
        'freight_value', 'customer_id', 'order_status', 'purchase_date',
        'payment_type', 'product_category_name', 'product_weight_gram'
    ]

# Convert numeric columns to the appropriate data types
data[['quantity', 'price', 'freight_value', 'product_weight_gram']] = data[
    ['quantity', 'price', 'freight_value', 'product_weight_gram']
].apply(pd.to_numeric, errors='coerce')

# Convert purchase_date to datetime
data['purchase_date'] = pd.to_datetime(data['purchase_date'], errors='coerce')

# 1. Check and prepare data to clean and handle missing values and ensure consistency
print("Initial Data Head:\n", data.head())  # Display the first few rows of the dataset
print("\nData Info:\n")
data.info()  # Display information about the dataset, including column types and non-null counts

# Check for missing values
missing_values = data.isnull().sum()  # Calculate the number of missing values in each column
print("\nMissing values before handling:\n", missing_values)

# Fill missing values
for column in data.columns:
    if data[column].dtype == 'object':
        # Fill missing values with mode (most frequent value) for categorical data
        mode_value = data[column].mode()[0]
        data[column] = data[column].fillna(mode_value)
    else:
        # Fill missing values with median for numerical data
        median_value = data[column].median()
        data[column] = data[column].fillna(median_value)

# Verify that there are no more missing values
missing_values_after = data.isnull().sum()  # Recalculate the number of missing values after filling
print("\nMissing values after handling:\n", missing_values_after)

# Save the cleaned data to a new Excel file
corrected_file_path = 'Corrected_File.xlsx'
data.to_excel(corrected_file_path, index=False)  # Save the cleaned dataset to an Excel file
print(f"Cleaned data saved to {corrected_file_path}")

# 2. Summarize the data with statistical analysis
data['total_sales'] = data['quantity'] * data['price']  # Calculate the total sales amount for each order

# Total sales amount
total_sales = data['total_sales'].sum()  # Calculate the total sales amount
print("\nTotal Sales Amount:", total_sales)

# Average sales amount per order
average_sales = data['total_sales'].mean()  # Calculate the average sales amount per order
print("Average Sales Amount per Order:", average_sales)

# Top product sales by quantity
top_products_quantity = data.groupby('product_id')['quantity'].sum().sort_values(ascending=False).head(10)  # Find the top 10 products by quantity sold
print("\nTop Products by Quantity:\n", top_products_quantity)

# Top product sales by revenue
top_products_revenue = data.groupby('product_id')['total_sales'].sum().sort_values(ascending=False).head(10)  # Find the top 10 products by total sales revenue
print("\nTop Products by Revenue:\n", top_products_revenue)

# 3. Use Statistical methods to identify significant correlations, comparisons, distributions, and trends
# Select only numeric columns for correlation matrix
numeric_columns = data.select_dtypes(include=['float64', 'int64'])  # Select only numeric columns for correlation matrix
correlations = numeric_columns.corr()  # Calculate the correlation matrix
print("\nCorrelation Matrix:\n", correlations)

# 4. Visualize the data with charts and graphs
# Top products by quantity
plt.figure(figsize=(10, 6))
top_products_quantity.plot(kind='bar')
plt.title('Top Products by Quantity')
plt.xlabel('Product ID')
plt.ylabel('Quantity Sold')
plt.savefig('top_products_quantity.png')  # Save the plot as a PNG file
plt.show()

# Top products by revenue
plt.figure(figsize=(10, 6))
top_products_revenue.plot(kind='bar')
plt.title('Top Products by Revenue')
plt.xlabel('Product ID')
plt.ylabel('Total Sales')
plt.savefig('top_products_revenue.png')  # Save the plot as a PNG file
plt.show()

# Sales trend over time
sales_trend = data.groupby(data['purchase_date'].dt.to_period('M'))['total_sales'].sum()  # Calculate monthly sales trend

plt.figure(figsize=(12, 6))
sales_trend.plot(kind='line')
plt.title('Sales Trend Over Time')
plt.xlabel('Month')
plt.ylabel('Total Sales')
plt.savefig('sales_trend.png')  # Save the plot as a PNG file
plt.show()

# Distribution of sales amount
plt.figure(figsize=(10, 6))
sns.histplot(data['total_sales'], bins=50, kde=True)  # Plot the distribution of total sales amount
plt.title('Distribution of Sales Amount')
plt.xlabel('Total Sales')
plt.ylabel('Frequency')
plt.savefig('distribution_of_sales.png')  # Save the plot as a PNG file
plt.show()

# Additional visualization: Correlation heatmap
plt.figure(figsize=(12, 8))
sns.heatmap(correlations, annot=True, cmap='coolwarm')  # Plot the correlation matrix as a heatmap
plt.title('Correlation Matrix Heatmap')
plt.savefig('correlation_heatmap.png')  # Save the plot as a PNG file
plt.show()

print("Exploratory Data Analysis complete. Graphs saved as PNG files.")
