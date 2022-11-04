# Analysis of annual sectoral demand against GDP, separated by refined products
# Correlation check

import pandas as pd
import seaborn as sns

df = pd.read_excel('/Users/jingwen.wang/Desktop/ASW/raw/ASIA.xlsx')

# Remove whitespaces
df = df.apply(lambda x: x.str.strip() if x.dtype == "object" else x)

# Remove whitespaces in header
df = df.rename(columns=lambda x:x.strip())

df.columns
# df.columns
"""
Index(['Year', 'GasolineTransformation', 'GasolineEnergy', 'GasolineIndustry',
       'GasolineTransportation', 'GasolineOther', 'GasolineTotalDomestic',
       'JetTransformation', 'JetEnergy', 'JetIndustry', 'JetTransportation',
       'JetOther', 'JetTotalDomestic', 'JetInternationalBunkers',
       'JetTotalPlusBunkers', 'DieselTransformation', 'DieselEnergy',
       'DieselIndustry', 'DieselTransportation', 'DieselOther',
       'DieselTotalDomestic', 'DieselInternationalBunkers',
       'DieselTotalPlusBunkers', 'FOTransformation', 'FOEnergy', 'FOIndustry',
       'FOTransportation', 'FOOther', 'FOTotalDomestic',
       'FOInternationalBunkers', 'FoTotalPlusBunkers', 'NaphthaTransformation',
       'NaphthaEnergy', 'NaphthaIndustry', 'NaphthaTransformation.1',
       'NaphthaOther', 'NaphthaTotalDomestic', 'TotalTransformation',
       'TotalEnergy', 'TotalIndustry', 'TotalTransportation', 'TotalOther',
       'TotalDomestic', 'TotalInternational', 'TotalPlusBunkers',
       'AustraliaGDP', 'BangladeshGDP', 'ChinaGDP', 'HongKongGDP', 'IndiaGDP',
       'IndonesiaGDP', 'JapanGDP', 'MalaysiaGDP', 'NewZealandGDP',
       'PakistanGDP', 'PhilippinesGDP', 'SingaporeGDP', 'SouthKoreaGDP',
       'SriLankaGDP', 'TaiwanGDP', 'ThailandGDP', 'VietnamGDP', 'TotalGDP'],
      dtype='object')
"""
# cm = df.corr()
# sns.heatmap(cm, annot=True)

# %%
# Using R value, also called Pearson's correlation coefficient
import matplotlib.pyplot as plt
plt.figure(figsize=(20, 10))

df_gasoline = df.loc[:,['GasolineTransformation', 'GasolineEnergy', 'GasolineIndustry',
                        'GasolineTransportation', 'GasolineOther', 'GasolineTotalDomestic', 'AustraliaGDP', 'BangladeshGDP', 'ChinaGDP', 'HongKongGDP', 'IndiaGDP',
                        'IndonesiaGDP', 'JapanGDP', 'MalaysiaGDP', 'NewZealandGDP',
                        'PakistanGDP', 'PhilippinesGDP', 'SingaporeGDP', 'SouthKoreaGDP',
                        'SriLankaGDP', 'TaiwanGDP', 'ThailandGDP', 'VietnamGDP', 'TotalGDP']]
cm_gasoline = df_gasoline.corr()
sns.heatmap(cm_gasoline, annot=True, cmap="rocket")

# R-squared 
r2 = cm_gasoline ** 2
plt.figure(figsize=(20, 10))
sns.heatmap(r2, annot=True, cmap="rocket")

# %%
plt.figure(figsize=(20, 10))
df_jet = df.loc[:,['JetTransformation', 'JetEnergy', 'JetIndustry', 'JetTransportation',
                   'JetOther', 'JetTotalDomestic', 'JetInternationalBunkers',
                   'JetTotalPlusBunkers', 'AustraliaGDP', 'BangladeshGDP', 'ChinaGDP', 'HongKongGDP', 'IndiaGDP',
                   'IndonesiaGDP', 'JapanGDP', 'MalaysiaGDP', 'NewZealandGDP',
                   'PakistanGDP', 'PhilippinesGDP', 'SingaporeGDP', 'SouthKoreaGDP',
                   'SriLankaGDP', 'TaiwanGDP', 'ThailandGDP', 'VietnamGDP', 'TotalGDP']]
cm_jet = df_jet.corr()
sns.heatmap(cm_jet, annot=True, cmap="rocket")

# R2
r2 = cm_jet ** 2
plt.figure(figsize=(20, 10))
sns.heatmap(r2, annot=True, cmap="rocket")

# %%
plt.figure(figsize=(20, 10))
df_diesel = df.loc[:,['DieselTransformation', 'DieselEnergy',
                      'DieselIndustry', 'DieselTransportation', 'DieselOther',
                      'DieselTotalDomestic', 'DieselInternationalBunkers',
                      'DieselTotalPlusBunkers', 'AustraliaGDP', 'BangladeshGDP', 'ChinaGDP', 'HongKongGDP', 'IndiaGDP',
                      'IndonesiaGDP', 'JapanGDP', 'MalaysiaGDP', 'NewZealandGDP',
                      'PakistanGDP', 'PhilippinesGDP', 'SingaporeGDP', 'SouthKoreaGDP',
                      'SriLankaGDP', 'TaiwanGDP', 'ThailandGDP', 'VietnamGDP', 'TotalGDP']]
cm_diesel = df_diesel.corr()
sns.heatmap(cm_diesel, annot=True, cmap="rocket")

# R2
r2 = cm_diesel ** 2
plt.figure(figsize=(20, 10))
sns.heatmap(r2, annot=True, cmap="rocket")

# %%
plt.figure(figsize=(20, 10))
df_FO = df.loc[:,['FOTransformation', 'FOEnergy', 'FOIndustry',
                  'FOTransportation', 'FOOther', 'FOTotalDomestic',
                  'FOInternationalBunkers', 'FoTotalPlusBunkers', 'AustraliaGDP', 'BangladeshGDP', 'ChinaGDP', 'HongKongGDP', 'IndiaGDP',
                  'IndonesiaGDP', 'JapanGDP', 'MalaysiaGDP', 'NewZealandGDP',
                  'PakistanGDP', 'PhilippinesGDP', 'SingaporeGDP', 'SouthKoreaGDP',
                  'SriLankaGDP', 'TaiwanGDP', 'ThailandGDP', 'VietnamGDP', 'TotalGDP']]
cm_FO = df_FO.corr()
sns.heatmap(cm_FO, annot=True, cmap="rocket")

# R2
r2 = cm_FO ** 2
plt.figure(figsize=(20, 10))
sns.heatmap(r2, annot=True, cmap="rocket")

# %%
plt.figure(figsize=(20, 10))
df_naphtha = df.loc[:,['NaphthaTransformation',
                       'NaphthaEnergy', 'NaphthaIndustry', 'NaphthaTransformation.1',
                       'NaphthaOther', 'NaphthaTotalDomestic', 'AustraliaGDP', 'BangladeshGDP', 'ChinaGDP', 'HongKongGDP', 'IndiaGDP',
                       'IndonesiaGDP', 'JapanGDP', 'MalaysiaGDP', 'NewZealandGDP',
                       'PakistanGDP', 'PhilippinesGDP', 'SingaporeGDP', 'SouthKoreaGDP',
                       'SriLankaGDP', 'TaiwanGDP', 'ThailandGDP', 'VietnamGDP', 'TotalGDP']]
cm_naphtha = df_naphtha.corr()
sns.heatmap(cm_naphtha, annot=True, cmap="rocket")

# R2
r2 = cm_naphtha ** 2
plt.figure(figsize=(20, 10))
sns.heatmap(r2, annot=True, cmap="rocket")