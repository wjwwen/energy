# R Notes
# data.table
***DT[i, j, by]*** <br>
***i***: subset rows <br>
***j***: select columns <br>
***j***: compute <br>
***by***: Grouping <br>
***keyby***: sorted by <br>
***chaining***: DT[...][...][...] <br>
***.sd***: i.e. Subset of Data, multiple columns in j <br>
```R
file <- fread(file.csv)
```
``` R
# chaining
ans <- flights[carrier == "AA", .N, by = .(origin, dest)][order(origin, -dest)]
head(ans, 10)
#     origin dest    N
#  1:    EWR  PHX  121
#  2:    EWR  MIA  848
#  3:    EWR  LAX   62
#  4:    EWR  DFW 1618
#  5:    JFK  STT  229
#  6:    JFK  SJU  690
#  7:    JFK  SFO 1312
#  8:    JFK  SEA  298
#  9:    JFK  SAN  299
# 10:    JFK  ORD  432
```
``` R
# .sd
DT[, print(.SD), by = ID]
#    a b  c
# 1: 1 7 13
# 2: 2 8 14
# 3: 3 9 15
#    a  b  c
# 1: 4 10 16
# 2: 5 11 17
#    a  b  c
# 1: 6 12 18
# Empty data.table (0 rows and 1 cols): ID
```

## Categorical variables - Nominal v.s. Ordinal
```R
# Categorical: Nominal v.s. Ordinal (i.e. with Order - S, M, L or Bronze, Silver, Gold)

# Convert variable to nominal
flights.dt$month <- factor(flights.dt$month)
flights.dt$day <- factor(flights.dt$day)
summary(flights.dt$month)
summary(flights.dt$day)

# Convert variable to ordinal
X <- c("S", "M", "L")
class(X)
summary(X)

X <- factor(X, ordered=T, levels=c("S", "M", "L", "XL"))
class(X)
levels(X)
plot(X, main="Figure 4.1: Distribution of Shirt Size")
```