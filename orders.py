import pandas as pd
from openpyxl import workbook
df = pd.read_excel(r"C:\Users\nitee\Downloads\SuperstoreData (1).xlsx", engine= 'openpyxl')

print(df.head())

print(df.columns) 
print(df.describe())
print(df.isnull().sum())

#handling missing values if any
df= df.dropna() 

#convert 'order date' and ship date' to detetime
df['Order Date'] = pd.to_datetime(df['Order Date'])
df['Ship Date'] = pd.to_datetime(df['Ship Date'])

#create new features 
df['Order Month'] = df['Order Date'].dt.to_period('M')
df['Profit Margin'] = df['Profit']/ df['Sales']

#verify changes
print(df.head()) 

#sales and profit analysis 
import matplotlib.pyplot as plt
#sales trends over time 
monthly_sales = df.groupby('Order Month')['Sales'].sum()
monthly_sales.plot(kind='line',title='Monthly Sales Trend',xlabel='Month',ylabel='Sales')
plt.show()

#Profit analysis by category
category_profit = df.groupby('Category')['Profit'].sum()
category_profit.plot(kind='bar',title='Profit by Category',xlabel='Category',ylabel='profit')
plt.show() 

#Profit analysis by region
region_profit = df.groupby('Region')['Profit'].sum()
region_profit.plot(kind='bar',title='Profit by Region',xlabel = 'Region',ylabel = 'Profit')
plt.show() 

#customer and product insights
#Top customers by sales
top_customers = df.groupby('Customer Name')['Sales'].sum().sort_values(ascending=False).head(10)
top_customers.plot(kind='bar',title='Top 10 customers by sales',xlabel='Customer',ylabel='Sales')
plt.show()

#Best-selling products
best_selling_products = df.groupby('Product Name')['Sales'].sum().sort_values(ascending=False).head(10)
best_selling_products.plot(kind='bar',title='Top 10 best selling products',xlabel='Product',ylabel='Sales')
plt.show()

#Customer segmentation (example: by total sales)

df['Customer Segment'] = pd.qcut(df.groupby('Customer ID')['Sales'].transform('sum'), 4, labels=['Low', 'Medium', 'High', 'Very High']) 
df.groupby('Customer Segment')['Sales'].sum().plot(kind='pie', autopct= '%1.1f%%', title='Sales by Customer Segment')
plt.show()

# Calculate KPIs

total_sales = df['Sales'].sum()
average_order_value = df['Sales'].mean() 
profit_margin = df['Profit'].sum() /df['Sales'].sum()

print(f"Total Sales: {total_sales}")
print(f"Average Order Value: {average_order_value}")
print(f"Profit Margin: {profit_margin}")

# Customer Lifetime Value (CLV) estimate (simplified)

clv = df.groupby('Customer ID') ['Profit'].sum().mean()
print(f"Average Customer Lifetime Value: {clv}")