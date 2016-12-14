
setwd("U:/CityWide Performance/Human Resources/Data/WorkersCompensation")

library("xlsx")
library("plyr")
library("dplyr")
library("tidyr")
library("reshape")
library("reshape2")
library("stringr")
library("zoo")
library("lubridate")


comp  <-  read.csv(file="workersComp1.csv", header=TRUE, na.strings = c("", NA), stringsAsFactors = FALSE)

comp1  <-  read.csv(file="workersComp11_22_16.csv", header=TRUE, na.strings = c("", NA), stringsAsFactors = FALSE)
comp1 <- comp1[,c(-4)]

comp <- do.call("rbind", list(comp, comp1))

comp <- subset(comp, Injury.Date !="Year Total")

#comp <- na.omit(comp)
comp$Injured.Worker <- NULL
comp$Risk.Location.Address <- NULL

write.csv(comp,"U:/CityWide Performance/Human Resources/Data/WorkersCompensation/COMPENSATION.csv", row.names = FALSE)

##Write for CovStat
write.xlsx(comp,"O:/AllUsers/CovStat/Data Portal/Repository/Data/HumanResources/WorkersCompensation.xlsx", row.names = FALSE)