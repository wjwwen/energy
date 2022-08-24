import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

# Oil price data 1987 - present
# Gas price data 1997 - present
# Skip rows to start from the same point

brent_oil_prices = pd.read_excel("https://www.eia.gov/dnav/pet/hist_xls/RBRTEd.xls",
                                 sheet_name="Data 1",
                                 skiprows=2455,
                                 names=['Date', 'Brent_Price'])

WTI_oil_prices = pd.read_excel("https://www.eia.gov/dnav/pet/hist_xls/RWTCd.xls",
                                 sheet_name="Data 1",
                                 skiprows=2801,
                                 names=['Date', 'WTI_Price'])

gas_prices = pd.read_excel("https://www.eia.gov/dnav/ng/hist_xls/RNGWHHDd.xls",
                                 sheet_name="Data 1",
                                 skiprows=2,
                                 names=['Date', 'Henry_Hub_Price'])

# ------------- DATA CLEANING
print(brent_oil_prices.shape)
print(gas_prices.shape)
print(WTI_oil_prices.shape)

# Determine isnull values
brent_oil_prices.isna().sum()
WTI_oil_prices.isna().sum()
gas_prices.isna().sum()

print(gas_prices[gas_prices["Henry_Hub_Price"].isnull()])
# Fill the null value with the value from previous row
gas_prices.fillna(method = 'ffill', inplace=True)

# ------------- PLOT
plt.plot(brent_oil_prices["Brent_Price"])
plt.plot(WTI_oil_prices["WTI_Price"])

# ------------- Daily % Change in Brent, WTI, Henry Hub Gas Price
# Histogram
pc_brent = brent_oil_prices["Brent_Price"].pct_change()
pc_brent.hist(bins=100)

pc_WTI = WTI_oil_prices["WTI_Price"].pct_change()
pc_WTI.hist(bins=100)

pc_gas = gas_prices["Henry_Hub_Price"].pct_change()
pc_gas.hist(bins=100)

# ------------- Probability BRENT
print("The probability of Brent oil price changes between 1%% and -1%% is %1.2f%% "
      % (100*pc_brent[(pc_brent>-0.01) & (pc_brent<0.01)].shape[0]/pc_brent.shape[0]))

print("The probability of Brent oil price changes between 3%% and -3%% is %1.2f%% "
      % (100*pc_brent[(pc_brent>-0.03) & (pc_brent<0.03)].shape[0]/pc_brent.shape[0]))      

print("The probability of Brent oil price changes between 5%% and -5%% is %1.2f%% "
      % (100*pc_brent[(pc_brent>-0.05) & (pc_brent<0.05)].shape[0]/pc_brent.shape[0]))    

print("The probability of Brent oil price changes more than 5%% is %1.2f%% "
      % (100*pc_brent[(pc_brent>0.05)].shape[0]/pc_brent.shape[0]))

print("The probability of Brent oil price changes less than 5%% is %1.2f%% "
      % (100*pc_brent[(pc_brent<-0.05)].shape[0]/pc_brent.shape[0]))

# ------------- MIN/MAX for Brent, WTI, Henry Hub Gas TTF
# ------------- BRENT
pc_brent.min(),pc_brent.idxmin(),pc_brent.max(),pc_brent.idxmax()
# (-0.4746543778801844, 5904, 0.5098684210526316, 5905)

brent_oil_prices.iloc[[5903,5904,5905]]
# Brent oil price decreased from $17/bbl to $9.12/bbl, between 20 April to 21 April 2020
# Brent oil price decreased from $9.12/bbl to $13.77/bbl, between 21 April to 22 April 2020

# -------------- WTI
pc_WTI.min(),pc_WTI.idxmin(),pc_WTI.max(),pc_WTI.idxmax()
#  (-3.0196613872200984, 5844, 0.5308641975308643, 5846)

WTI_oil_prices.iloc[[5843,5844,5845]]
# WTI oil price decreased from $18.31/bbl to -$36.98/bbl between 17 April and 20 April 2020
WTI_oil_prices.iloc[[5845,5846,5847]]
# WTI oil price increased from $8.91/bbl to -$13.64/bbl between 21 April and 22 April 2020

# -------------- GAS
pc_gas.min(),pc_gas.idxmin(),pc_gas.max(),pc_gas.idxmax()
# (-0.6412405699916177, 6066, 1.1077738515901059, 6065)

gas_prices.iloc[[6065,6066,6067]]
# Henry Hub natural gas prices decreased from $23.6 MMBtu to $8.56 MMBtu from 17 February to 18 February 2021
# Climate conditions/domestic gas storage & infrastructure is important for stable gas prices
