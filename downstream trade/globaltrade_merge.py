# purpose: merge all global trade import regional files to facilitate SQL cleaning 
# (i.e. removal of specific trade routes after, using SQL)

# merge excel files with specific tab name "LinksOut_PBI"
import pandas as pd
df1 = pd.read_excel('/Users/jingwen.wang/IHS Markit/Downstream_SG - Documents/Models/Country Product Trade/GTI/Global Trade/Asia_Trade.xlsx', 'LinksOut_PBI')
df2 = pd.read_excel('/Users/jingwen.wang/IHS Markit/Downstream_SG - Documents/Models/Country Product Trade/GTI/Global Trade/Middle East_Trade.xlsx', 'LinksOut_PBI')
df = pd.concat([df1,df2],ignore_index = True)
df.fillna(method='ffill')
df.to_csv('combined_trade_imports_AsiaME.csv', index = False)

#%% 
df4 = pd.read_excel('/Users/jingwen.wang/IHS Markit/Downstream_SG - Documents/Models/Country Product Trade/GTI/Global Trade/Europe_Trade.xlsx', 'LinksOut_PBI')
df6 = pd.read_excel('/Users/jingwen.wang/IHS Markit/Downstream_SG - Documents/Models/Country Product Trade/GTI/Global Trade/Latin America_Trade.xlsx', 'LinksOut_PBI') 
df = pd.concat([df4,df6],ignore_index = True)
df.fillna(method='ffill')
df.to_csv('combined_trade_imports_EuropeLatam.csv', index = False)

#%%
df3 = pd.read_excel('/Users/jingwen.wang/IHS Markit/Downstream_SG - Documents/Models/Country Product Trade/GTI/Global Trade/Africa_Trade.xlsx', 'LinksOut_PBI')  
df5 = pd.read_excel('/Users/jingwen.wang/IHS Markit/Downstream_SG - Documents/Models/Country Product Trade/GTI/Global Trade/North America_Trade.xlsx', 'LinksOut_PBI')      
df = pd.concat([df3, df5],ignore_index = True)
df.fillna(method='ffill')
df.to_csv('combined_trade_imports_AfricaNAM.csv', index = False)

# %%
# Removal of CIS data from import files (Check) ------------------------------
# C:\Users\jingwen.wang\Desktop\SQL
df_clean = pd.read_csv('/Users/jingwen.wang/Desktop/SQL/combined_trade_imports.csv', encoding="utf-8")

# column names
for col in df_clean.columns:
    print(col)

# Dropping strings containing CIS, Russia, Kazakhstan
df_clean[df_clean["Trade Partner"].str.contains("Russia|CIS|Kazakhstan") == False]

df_clean.to_csv('cleaned_trade_imports.csv', index = False)

# Check:
# Extract historical data from GTA Imports
# Extract Africa/Middle East data from GTA Exports

