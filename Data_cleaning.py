import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns

df = pd.read_excel("/content/Raw_Data.xlsx")

df.head()

df.shape

df.info()

df.columns

# Transpose and check duplicates
print(df.T.duplicated())

df = df.loc[:, ~df.T.duplicated()]

df.columns

df.isna().sum()


# Fact table (sales transactions)
fact_sales = df[
    [
        'order_id',
        'order_item_id',
        'customer_id',
        'product_id',
        'quantity',
        'unit_price',
        'discount_pct',
        'discount_amount',
        'gross_amount',
        'net_amount',
        'shipping_cost',
        'line_total'
    ]
].drop_duplicates()


# Dimension tables

customers_table = df[
    [
        'customer_id',
        'customer_name',
        'gender',
        'age',
        'shipping_address',
        'city',
        'country',
        'signup_date',
        'customer_segment'
    ]
].drop_duplicates()


products_table = df[
    [
        'product_id',
        'product_name',
        'category',
        'brand',
        'supplier',
        'cost_price',
        'unit_price',
        'weight_kg'
    ]
].drop_duplicates()


orders_table = df[
    [
        'order_id',
        'order_date',
        'delivery_date',
        'order_status',
        'payment_method',
        'shipping_mode'
    ]
].drop_duplicates()


inventory_table = df[
    [
        'product_id',
        'total_stock_qty',
        'avg_reorder_level',
        'overall_stock_status',
        'warehouses_stocked'
    ]
].drop_duplicates()


reviews_table = df[
    [
        'customer_id',
        'product_id',
        'review_rating',
        'customer_review',
        'review_date'
    ]
].drop_duplicates()


# Export all tables into one Excel workbook

with pd.ExcelWriter("cleaned.xlsx") as writer:

    fact_sales.to_excel(
        writer,
        sheet_name="Fact_Sales",
        index=False
    )

    customers_table.to_excel(
        writer,
        sheet_name="Dim_Customers",
        index=False
    )

    products_table.to_excel(
        writer,
        sheet_name="Dim_Products",
        index=False
    )

    orders_table.to_excel(
        writer,
        sheet_name="Dim_Orders",
        index=False
    )

    inventory_table.to_excel(
        writer,
        sheet_name="Dim_Inventory",
        index=False
    )

    reviews_table.to_excel(
        writer,
        sheet_name="Dim_Reviews",
        index=False
    )
