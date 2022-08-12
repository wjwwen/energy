library(dplyr)
library(ggplot2)
library(factoextra)
# factoextra - extract and visualize results of multivariate data analyses
# install.packages("dplyr")

df <- read.csv(file.choose(),header=T)

df %>%
  group_by(region,Year) %>%
  summarise(total=sum(Fossil.fuels..TWh.growth...sub.method.) , .groups = "drop_last") %>%
  ggplot(aes(Year,total , colour=region)) +
  geom_line()+
  ggtitle("Total fossil fuels in TWh consumed from 1966 to 2020 by continents")


df %>%
  filter(sub.region=="Eastern Asia") %>%
  group_by(Entity,Year) %>%
  summarise(total=sum(Fossil.fuels..TWh.growth...sub.method.) , .groups = "drop_last") %>%
  ggplot(aes(Year,total , colour=Entity)) +
  geom_line()+
  ggtitle("Total fossil fuels in TWh consumed from 1966 to 2020 in East Asia")

df %>%
  filter(sub.region=="Southern Asia") %>%
  group_by(Entity,Year) %>%
  summarise(total=sum(Fossil.fuels..TWh.growth...sub.method.) , .groups = "drop_last") %>%
  ggplot(aes(Year,total , colour=Entity)) +
  geom_line()+
  ggtitle("Total fossil fuels in Twh consumed from 1966 to 2020 in South Asia")