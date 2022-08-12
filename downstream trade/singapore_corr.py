import pandas as pd

# singapore net trade (imports-exports) and stocks --------------------------
df = pd.DataFrame(
   {
      "Net Trade": [66,-17,-2,7,60,30,-55,42,10,-6,-20,-30,25],
      "Stocks": [14816,14559,13744,15461,15260,15121,15812,15212,15229,16839,15158,16206,18021]
   }
)

# pearson - interval scale (linear relationship between 2 continuous variables)
# spearman - ordinal scale (monotonic relationship, ranked values for each variable rather than raw data)
col1, col2 = "Net Trade", "Stocks"
corr = df[col1].corr(df[col2], method="pearson")
print("Correlation between", col1, "and", col2, "is:", round(corr, 2))

# singapore imports/exports --------------------------------------------------
df2 = pd.DataFrame(
   {
      "Imports": [93,46,73,71,82,68,24,94,20,35,37,57,89,69],
      "Exports": [28,63,75,64,22,38,79,52,10,41,57,87,64,42]
   }
)

col1, col2 = "Imports", "Exports"
corr = df2[col1].corr(df2[col2], method="pearson")
print("Correlation between", col1, "and", col2, "is:", round(corr, 2))