# Linear Regression of price against demand and price elasticity
# Monthly data between 2010-2021 for India
library(ggplot2)

data=read.csv(file.choose(),header=T)
demand_lm = lm(India.Demand~India.Price, data=data)
summary(demand_lm)

ggplot(data = data, aes(y = India.Demand, x = India.Price)) + 
  geom_point(col = 'blue') + geom_smooth(method = 'lm', col = 'red', size = 0.5) # fitted regression line
#### as price increases 1 unit, quantity demand increases by 7.6463 kbd 
KBD_India = 7.6463*1000 # 7646.3 b/d

avg_demand = mean(data$India.Demand)
avg_price = mean(data$India.Price)
# coefficient estimate * (avg price/avg demand)
elasticity = 7.6463*(avg_price/avg_demand) # 1.10862

# JAPAN RETAIL GASOLINE
data=read.csv(file.choose(),header=T)
demand_lm = lm(Japan.Demand~Japan.Price, data=data)
summary(demand_lm)

ggplot(data = data, aes(y = Japan.Demand, x = Japan.Price)) + 
  geom_point(col = 'blue') + geom_smooth(method = 'lm', col = 'red', size = 0.5) # fitted regression line
#### as price increases 1 unit, quantity demand increases by 0.634-kbd increase in demand kbd
KBD = 0.634*1000 # 634 barrels per day (634 b/d) 


avg_demand = mean(data$Japan.Demand)
avg_price = mean(data$Japan.Price)
# coefficient estimate * (avg price/avg demand)
elasticity = 0.6345*avg_price/avg_demand # 1.11394