import sqlite3
import pandas as pd

# Connect to SQLite database
conn = sqlite3.connect('churn_analysis.db')

# Query 1: Churn by Internet Service
query1 = """
SELECT 
    InternetService,
    COUNT(*) AS TotalCustomers,
    AVG(AttritionFlag) AS AttritionRate
FROM customers
GROUP BY InternetService;
"""
df1 = pd.read_sql_query(query1, conn)

# Query 2: Churn by Contract Type
query2 = """
SELECT 
    Contract,
    COUNT(*) AS TotalCustomers,
    AVG(AttritionFlag) AS AttritionRate
FROM customers
GROUP BY Contract;
"""
df2 = pd.read_sql_query(query2, conn)

# Query 3: Churn by Tenure Buckets
# First, create a temporary column in SQL using CASE
query3 = """
SELECT 
    CASE
        WHEN tenure <= 12 THEN 'Short Term (0-1 yr)'
        WHEN tenure BETWEEN 13 AND 36 THEN 'Mid Term (1-3 yrs)'
        ELSE 'Long Term (3+ yrs)'
    END AS TenureGroup,
    COUNT(*) AS TotalCustomers,
    AVG(AttritionFlag) AS AttritionRate
FROM customers
GROUP BY TenureGroup;
"""
df3 = pd.read_sql_query(query3, conn)

# Export all 3 DataFrames to the same Excel file (3 sheets)
with pd.ExcelWriter('Churn_Insights_Report.xlsx', engine='openpyxl') as writer:
    df1.to_excel(writer, sheet_name='By_InternetService', index=False)
    df2.to_excel(writer, sheet_name='By_Contract', index=False)
    df3.to_excel(writer, sheet_name='By_Tenure_Buckets', index=False)

# Close connection
conn.close()

print("✅ Churn Insights Excel report created successfully!")


