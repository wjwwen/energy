library(xts)
library(ggplot2)
library(forecast)

dataurl <- 'https://.../mme.csv'

# Downloading/loading data
tmp <- tempfile()
download.file(dataurl, tmp, method = 'curl')
df <- read.csv(tmp)
unlink(tmp)

# Convert date strings to POSIX dates
df$date <- strptime(df$date, format = '%d-%m-%y')
# Day of the week
df$day <- as.factor(strftime(df$date, format = '%A'))
# Day of the year
df$yearday <- as.factor(strftime(df$date, format = '%m%d'))
# Final structure for the study
str(df)

df_test <- subset(df, date >= strptime('01-01-2011', format = '%d-%m-%Y'))
df <- subset(df, date < strptime('01-01-2011', format = '%d-%m-%Y'))
ts <- ts(df$demand, frequency = 1)

# Df and time series objects
demandts <- xts(df$demand, df$date)
plot(demandts, main = 'Energy demand evolution', xlab = 'Date', ylab = 'Demand (GWh)')

# Demand by day of the week
ggplot(df, aes(day, demand)) + geom_boxplot() + xlab('Day') + ylab('Demand (GWh)') + ggtitle('Demand per day of the week')

# Aggregating demand by day of the year (avg)
avg_demand_per_yearday <- aggregate(demand ~ yearday, df, 'mean')

# Computing the smooth curve for the time series. Data is replicated before computing the curve in order to achieve continuity
smooth_yearday <- rbind(avg_demand_per_yearday, avg_demand_per_yearday, avg_demand_per_yearday, avg_demand_per_yearday, avg_demand_per_yearday)
smooth_yearday <- lowess(smooth_yearday$demand, f = 1 / 45)
l <- length(avg_demand_per_yearday$demand)
l0 <- 2 * l + 1
l1 <- 3 * l
smooth_yearday <- smooth_yearday$y[l0:l1]

# Plot
par(mfrow = c(1, 1))
# Setting year to 2000 to allow existence of 29th February
dates <- as.Date(paste(levels(df$yearday), '2000'), format = '%m%d%Y')
plot(dates, avg_demand_per_yearday$demand, type = 'l', main = 'Average daily demand', xlab = 'Date', ylab = 'Demand (GWh)')
lines(dates, smooth_yearday, col = 'blue', lwd = 2)

par(mfrow = c(1, 2))
diff <- avg_demand_per_yearday$demand - smooth_yearday
abs_diff <- abs(diff)
barplot(diff[order(-abs_diff)], main = 'Smoothing error', ylab = 'Error')
boxplot(diff, main = 'Smoothing error', ylab = 'Error')

head(strftime(dates[order(-abs_diff)], format = '%B %d'), 10)

par(mfrow = c(2, 2))
acf(df$demand, 100, main = 'Autocorrelation')
acf(df$demand, 1500, main = 'Autocorrelation')
pacf(df$demand, 100, main = 'Partial autocorrelation')
pacf(df$demand, 1500, main = 'Partial autocorrelation')

wts <- ts(ts, frequency = 7)
dec_wts <- decompose(wts)
plot(dec_wts)

# Demand - week seasonal
df$demand_mws <- df$demand - as.numeric(dec_wts$season)

yts <- ts(subset(df, yearday != '0229')$demand_mws, frequency = 365)
dec_yts <- decompose(yts)
plot(dec_yts)
