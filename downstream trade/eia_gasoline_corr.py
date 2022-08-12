import pandas as pd

# US EIA Weekly
# Gasoline Stocks - US Ending Stocks of Total Gasoline, Weekly (Thousand bbl)
# Gasoline Supplies - US Product Supplied of Finished Motor Gasoline, Weekly (Thousand b/d)
# Days of Supply = Stocks/Supplies (Days)

df = pd.read_excel('/Users/jingwen.wang/Desktop/US_EIA_Gasoline.xlsx', 'USEIA')

corr1 = df['Gasoline Stocks'].corr(df['Gasoline Supplies'], method="pearson")
print(f"the correlation between gasoline stocks and supplies is {corr1}")

corr2 = df['Gasoline Stocks'].corr(df['Days of Supply'])
print(f"the correlation between gasoline stocks and days of supply is {corr2}")

corr3 = df['Gasoline Supplies'].corr(df['Days of Supply'])
print(f"the correlation between gasoline supplies and days of supply is {corr3}")