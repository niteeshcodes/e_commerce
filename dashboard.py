import pandas as pd
import dash
from dash import dcc, html
from dash.dependencies import Input, Output
import plotly.express as px

# Load the dataset
df = pd.read_excel(r"C:\Users\nitee\Downloads\SuperstoreData (1).xlsx", engine='openpyxl')

# Handle missing values
df = df.dropna()

# Convert 'Order Date' and 'Ship Date' to datetime
df['Order Date'] = pd.to_datetime(df['Order Date'])
df['Ship Date'] = pd.to_datetime(df['Ship Date'])

# Create new features
df['Order Month'] = df['Order Date'].dt.to_period('M')
df['Profit Margin'] = df['Profit'] / df['Sales']

# Initialize the Dash app
app = dash.Dash(__name__)

# Layout of the dashboard
app.layout = html.Div([
    html.H1("Superstore Dashboard"),
    
    dcc.Graph(id='monthly-sales-trend'),
    dcc.Graph(id='profit-by-category'),
    dcc.Graph(id='profit-by-region'),
    dcc.Graph(id='top-customers'),
    dcc.Graph(id='best-selling-products'),
    dcc.Graph(id='sales-by-customer-segment')
])

# Callbacks to update the graphs
@app.callback(
    Output('monthly-sales-trend', 'figure'),
    Input('monthly-sales-trend', 'id')
)
def update_monthly_sales_trend(_):
    monthly_sales = df.groupby('Order Month')['Sales'].sum().reset_index()
    monthly_sales['Order Month'] = monthly_sales['Order Month'].astype(str)
    fig = px.line(monthly_sales, x='Order Month', y='Sales', title='Monthly Sales Trend')
    return fig

@app.callback(
    Output('profit-by-category', 'figure'),
    Input('profit-by-category', 'id')
)
def update_profit_by_category(_):
    category_profit = df.groupby('Category')['Profit'].sum().reset_index()
    fig = px.bar(category_profit, x='Category', y='Profit', title='Profit by Category')
    return fig

@app.callback(
    Output('profit-by-region', 'figure'),
    Input('profit-by-region', 'id')
)
def update_profit_by_region(_):
    region_profit = df.groupby('Region')['Profit'].sum().reset_index()
    fig = px.bar(region_profit, x='Region', y='Profit', title='Profit by Region')
    return fig

@app.callback(
    Output('top-customers', 'figure'),
    Input('top-customers', 'id')
)
def update_top_customers(_):
    top_customers = df.groupby('Customer Name')['Sales'].sum().sort_values(ascending=False).head(10).reset_index()
    if 'Customer Name' not in top_customers.columns or 'Sales' not in top_customers.columns:
        raise ValueError("DataFrame does not contain the required columns.")
    fig = px.bar(top_customers, x='Customer Name', y='Sales', title='Top 10 Customers by Sales')
    return fig

@app.callback(
    Output('best-selling-products', 'figure'),
    Input('best-selling-products', 'id')
)
def update_best_selling_products(_):
    best_selling_products = df.groupby('Product Name')['Sales'].sum().sort_values(ascending=False).head(10).reset_index()
    fig = px.bar(best_selling_products, x='Product Name', y='Sales', title='Top 10 Best Selling Products')
    return fig

@app.callback(
    Output('sales-by-customer-segment', 'figure'),
    Input('sales-by-customer-segment', 'id')
)
def update_sales_by_customer_segment(_):
    df['Customer Segment'] = pd.qcut(df.groupby('Customer ID')['Sales'].transform('sum'), 4, labels=['Low', 'Medium', 'High', 'Very High'])
    customer_segment_sales = df.groupby('Customer Segment')['Sales'].sum().reset_index()
    fig = px.pie(customer_segment_sales, names='Customer Segment', values='Sales', title='Sales by Customer Segment')
    return fig

# Run the app
if __name__ == '__main__':
    app.run_server(debug=True)
