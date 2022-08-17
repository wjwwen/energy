# LSTM

# 5 independent variables in csv
# 1. Gold Futures (GOLD)
# 2. S&P 500 Futures (SP500)
# 3. US Dollar Index (USDINDEX)
# 4. US 10-Year Bond Yield (US$10B)
# 5. Dow Jones Utilities Average (DJU)

import numpy as np 
import pandas as pd 
import matplotlib.pyplot as plt 
from sklearn.preprocessing import MinMaxScaler
from keras.models import Sequential
from keras.layers import Dense, LSTM, Dropout

df = pd.read_csv('prices.csv') 
df["WTI"]=pd.to_numeric(df.WTI,errors='coerce') 
df["Gold"]=pd.to_numeric(df["GOLD"],errors='coerce') 
df = df.dropna() 

# identify column headers
for col in df:
    print(col)

corr1 = df['WTI'].corr(df['GOLD'])
print(f"the correlation is {corr1}")

corr2 = df['WTI'].corr(df['SP 500'])
print(f"the correlation is {corr2}")

corr3 = df['WTI'].corr(df['US DOLLAR INDEX'])
print(f"the correlation is {corr3}")

corr4 = df['WTI'].corr(df['US 10YR BOND'])
print(f"the correlation is {corr4}")

corr5 = df['WTI'].corr(df['DJU'])
print(f"the correlation is {corr5}")

# %%
# DATA PREPROCESSING
# train 80%, test 20%
price_data = df.iloc[:, 1:2].values
price_data = price_data.reshape((-1,1)) 

split_percent = 0.80
split = int(split_percent*len(price_data))

price_train = price_data[:split]
price_test = price_data[split:]

date_train = df['Date'][:split]
date_test = df['Date'][split:]

print(len(price_train))
print(len(price_test))

# %%
trainData = price_train

# normalize features
sc = MinMaxScaler(feature_range=(0,1))
trainData = sc.fit_transform(trainData)

trainData.shape

X_train = []
y_train = []

for i in range (60, 3879): # 60: timestep // 3879: length of the data
    X_train.append(trainData[i-60:i,0]) 
    y_train.append(trainData[i,0])

X_train,y_train = np.array(X_train),np.array(y_train)

X_train = np.reshape(X_train,(X_train.shape[0],X_train.shape[1],1)) # adding the batch_size axis
X_train.shape

# Design network
model = Sequential()

model.add(LSTM(units=100, return_sequences = True, input_shape = (X_train.shape[1], 1)))
model.add(Dropout(0.2))

model.add(LSTM(units=100, return_sequences = True))
model.add(Dropout(0.2))

model.add(LSTM(units=100, return_sequences = True))
model.add(Dropout(0.2))

model.add(LSTM(units=100, return_sequences = False))
model.add(Dropout(0.2))

# Adam - adaptive learning rate optimization algorithm designed specifically for training deep neural networks
model.add(Dense(units = 1))
model.compile(optimizer='adam',loss="mean_squared_error")

# Fit network
hist = model.fit(X_train, y_train, epochs = 20, batch_size = 32, verbose=2)

# %%
testData = price_test
y_test = testData[60:,0:] # selecting the labels 

# Input array for the model
inputClosing = testData
inputClosing_scaled = sc.transform(inputClosing)
inputClosing_scaled.shape

X_test = []
length = len(testData)
timestep = 60
for i in range(timestep,length): # doing the same preprocessing 
    X_test.append(inputClosing_scaled[i-timestep:i,0])
X_test = np.array(X_test)
X_test = np.reshape(X_test,(X_test.shape[0],X_test.shape[1],1))
X_test.shape

# predicting the new values
y_pred = model.predict(X_test) 

# inverse the scaling transformation for plotting
predicted_price = sc.inverse_transform(y_pred) 

# %%
plt.plot(y_test, color = 'black', label = 'Actual Price')
plt.plot(predicted_price, color = 'red', label = 'Predicted Price')
plt.title('WTI/Gold - LSTM Model')
plt.xlabel('Time')
plt.ylabel('Price')
plt.legend()
plt.show()

plt.plot(hist.history['loss'])
plt.title('model loss')
plt.ylabel('loss')
plt.xlabel('epoch')
plt.legend(['train'], loc='upper left')
plt.show()

def rmse(predicted_price, y_test):
    return np.sqrt(((predicted_price - y_test) ** 2).mean())

rmse(predicted_price, y_test)