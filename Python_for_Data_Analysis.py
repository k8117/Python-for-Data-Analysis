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

# Visualizing the top products by quantity sold
# - Set the figure size for better visibility
# - Plot a bar chart of top products by quantity
# - Add title, x-axis label, and y-axis label for clarity
# - Save the plot as a PNG file for future reference
# - Display the plot
plt.figure(figsize=(10, 6))
top_products_quantity.plot(kind='bar')
plt.title('Top Products by Quantity')
plt.xlabel('Product ID')
plt.ylabel('Quantity Sold')
plt.savefig('top_products_quantity.png')
plt.show()

# Visualizing the top products by revenue
# - Set the figure size for better visibility
# - Plot a bar chart of top products by revenue
# - Add title, x-axis label, and y-axis label for clarity
# - Save the plot as a PNG file for future reference
# - Display the plot
plt.figure(figsize=(10, 6))
top_products_revenue.plot(kind='bar')
plt.title('Top Products by Revenue')
plt.xlabel('Product ID')
plt.ylabel('Total Sales')
plt.savefig('top_products_revenue.png')
plt.show()


# Calculating and visualizing the monthly sales trend
# - Group data by month and calculate total sales for each month
# - Set the figure size for better visibility
# - Plot a line chart to show sales trend over time
# - Add title, x-axis label, and y-axis label for clarity
# - Save the plot as a PNG file for future reference
# - Display the plot
sales_trend = data.groupby(data['purchase_date'].dt.to_period('M'))['total_sales'].sum()

plt.figure(figsize=(12, 6))
sales_trend.plot(kind='line')
plt.title('Sales Trend Over Time')
plt.xlabel('Month')
plt.ylabel('Total Sales')
plt.savefig('sales_trend.png')
plt.show()


# Visualizing the distribution of total sales amount
# - Set the figure size for better visibility
# - Plot a histogram with 50 bins and a KDE overlay to show the distribution
# - Add title, x-axis label, and y-axis label for clarity
# - Save the plot as a PNG file for future reference
# - Display the plot
plt.figure(figsize=(10, 6))
sns.histplot(data['total_sales'], bins=50, kde=True)
plt.title('Distribution of Sales Amount')
plt.xlabel('Total Sales')
plt.ylabel('Frequency')
plt.savefig('distribution_of_sales.png')
plt.show()


# Visualizing the correlation matrix as a heatmap
# - Set the figure size for better visibility
# - Plot a heatmap with annotation for correlation coefficients
# - Use a color gradient for better differentiation of values
# - Add title for clarity
# - Save the plot as a PNG file for future reference
# - Display the plot
plt.figure(figsize=(12, 8))
sns.heatmap(correlations, annot=True, cmap='coolwarm')
plt.title('Correlation Matrix Heatmap')
plt.savefig('correlation_heatmap.png')
plt.show()

print("Exploratory Data Analysis complete. Graphs saved as PNG files.")


'''
Initial Data Head:
                            order_id  quantity                        product_id  ...     payment_type product_category_name  product_weight_gram
0  2e7a8482f6fb09756ca50c10d7bfc047         2  f293394c72c9b5fafd7023301fc21fc2  ...  virtual account               fashion               1800.0 
1  2e7a8482f6fb09756ca50c10d7bfc047         1  c1488892604e4ba5cff5b4eb4d595400  ...  virtual account            automotive               1400.0 
2  e5fa5a7210941f7d56d0208e4e071d35         1  f3c2d01a84c947b078e32bbef0718962  ...         e-wallet                  toys                700.0 
3  3b697a20d9e427646d92567910af6d57         1  3ae08df6bcbfe23586dd431c40bddbb7  ...         e-wallet             utilities                300.0 
4  71303d7e93b399f5bcd537d124c0bcfa         1  d2998d7ced12f83f9b832f33cf6507b6  ...         e-wallet               fashion                500.0 

[5 rows x 12 columns]

Data Info:

<class 'pandas.core.frame.DataFrame'>
RangeIndex: 49999 entries, 0 to 49998
Data columns (total 12 columns):
 #   Column                 Non-Null Count  Dtype
---  ------                 --------------  -----
 0   order_id               49999 non-null  object
 1   quantity               49999 non-null  int64
 2   product_id             49999 non-null  object
 3   price                  49999 non-null  int64
 4   seller_id              49999 non-null  object
 5   freight_value          49999 non-null  int64
 6   customer_id            49999 non-null  object
 7   order_status           49999 non-null  object
 8   purchase_date          19881 non-null  datetime64[ns]
 9   payment_type           49999 non-null  object
 10  product_category_name  49999 non-null  object
 11  product_weight_gram    49980 non-null  float64
dtypes: datetime64[ns](1), float64(1), int64(3), object(7)
memory usage: 4.6+ MB

Missing values before handling:
 order_id                     0
quantity                     0
product_id                   0
price                        0
seller_id                    0
freight_value                0
customer_id                  0
order_status                 0
purchase_date            30118
payment_type                 0
product_category_name        0
product_weight_gram         19
dtype: int64

Missing values after handling:
 order_id                 0
quantity                 0
product_id               0
price                    0
seller_id                0
freight_value            0
customer_id              0
order_status             0
purchase_date            0
payment_type             0
product_category_name    0
product_weight_gram      0
dtype: int64
Cleaned data saved to Corrected_File.xlsx

Total Sales Amount: 156053194000
Average Sales Amount per Order: 3121126.3025260507

Top Products by Quantity:
 product_id
422879e10f46682990de24d770e7f83d    464
99a4788cb24856965c36a24e339b6058    406
389d119b48cf3043d311335e499d9c6b    285
53759a2ecddad2bb87a079a1f1519f73    275
154e7e31ebfa092203795c972e5804a6    237
368c6c730842d78016ad823897a372db    231
d5991653e037ccb7af6ed7d94246b249    228
9571759451b1d780ee7c15012ea109d4    210
7c1bd920dbdf22470b68bde975dd3ccf    181
42a2c92a0979a949ca4ea89ec5c7b934    179
Name: quantity, dtype: int64

Top Products by Revenue:
 product_id
422879e10f46682990de24d770e7f83d    1243964000
99a4788cb24856965c36a24e339b6058    1028276000
389d119b48cf3043d311335e499d9c6b     780584000
53759a2ecddad2bb87a079a1f1519f73     715026000
9571759451b1d780ee7c15012ea109d4     627470000
368c6c730842d78016ad823897a372db     613350000
154e7e31ebfa092203795c972e5804a6     596916000
d5991653e037ccb7af6ed7d94246b249     496697000
7c1bd920dbdf22470b68bde975dd3ccf     471974000
270516a3f41dc035aa87d220228f844c     442546000
Name: total_sales, dtype: int64

Correlation Matrix:
                      quantity     price  freight_value  product_weight_gram  total_sales
quantity             1.000000 -0.001649      -0.009926            -0.009229     0.691926
price               -0.001649  1.000000       0.005095             0.002769     0.607551
freight_value       -0.009926  0.005095       1.000000            -0.005228    -0.004290
product_weight_gram -0.009229  0.002769      -0.005228             1.000000    -0.005351
total_sales          0.691926  0.607551      -0.004290            -0.005351     1.000000
Exploratory Data Analysis complete. Graphs saved as PNG files.
'''
