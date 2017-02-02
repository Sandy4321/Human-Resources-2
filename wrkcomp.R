
setwd("U:/CityWide Performance/CovStat/CovStat Projects/Human Resources/WorkersCompensation")

library("xlsx")
library("plyr")
library("dplyr")
library("tidyr")
library("reshape")
library("reshape2")
library("stringr")
library("zoo")
library("lubridate")

#comp  <-  read.csv(file="workersComp1.csv", header=TRUE, na.strings = c("", NA), stringsAsFactors = FALSE)

comp1  <-  read.csv(file="WC 1-30-17.csv", header=TRUE, na.strings = c("", NA), stringsAsFactors = FALSE)
comp1 <- subset(comp1, Injury.Date !="Year Total")
#comp <- do.call("rbind", list(comp, comp1))

comp_all <- comp1
comp_all$Count <- 1

#comp <- na.omit(comp)
######################
### SQLite Storage ###
######################
cons.hr <- dbConnect(drv=RSQLite::SQLite(), dbname="O:/AllUsers/CovStat/Data Portal/repository/Data/Database Files/HumanResources.db")
dbWriteTable(cons.hr, "WorkersComp", comp_all, overwrite = TRUE)
dbDisconnect(cons.hr)

######################################################
### Who files the most claims and how much incurred ##
######################################################
all_comp <- subset(comp_all, Policy.Year >= '2010')

## remove '$' and ',' from total incurred column and change class to numeric
all_comp$Total.Incurred <- as.numeric(gsub("\\$*\\,*", "", all_comp$Total.Incurred, x = all_comp$Total.Incurred))

## aggregate count and total incurred by staff name and department  
comp_agg <- aggregate(cbind(Count,Total.Incurred) ~ Injured.Worker + Department, all_comp, sum)
comp_agg <- comp_agg[order(-comp_agg$Count, -comp_agg$Total.Incurred, comp_agg$Injured.Worker),]

## get percentage of claims and costs for each staff ##
count_pct <- prop.table(comp_agg$Count)*100 
incurred_pct <- prop.table(comp_agg$Total.Incurred)*100
comp_agg <- cbind(comp_agg, count_pct, incurred_pct)

## assign multiple if more than 1 claim
for (i in 1:length(comp_agg$Count)){
  if(comp_agg$Count[i] > 1)
    comp_agg$Level[i] <- "multiple"
  else comp_agg$Level[i] <- "not multiple"
}

####################################
## Claims and costs by department ##
####################################
## Fire ##
##########
fire_comp <- all_comp
fire_comp$Department <- trimws(fire_comp$Department, "both")
fire_comp <- subset(fire_comp, Department == "Fire Department")
fire_agg <- aggregate(cbind(Count,Total.Incurred) ~ Injured.Worker + Department, fire_comp, sum)
fire_agg <- fire_agg[order(-fire_agg$Count, -fire_agg$Total.Incurred, fire_agg$Injured.Worker),]

## get percentage of claims and costs for each staff ##
Fcount_pct <- round(prop.table(fire_agg$Count)*100, 3)
Fincurred_pct <- round(prop.table(fire_agg$Total.Incurred)*100, 3)
fire_agg <- cbind(fire_agg, Fcount_pct, Fincurred_pct)

## assign multiple if more than 1 claim
for (i in 1:length(fire_agg$Count)){
  if(fire_agg$Count[i] > 1)
    fire_agg$Level[i] <- "multiple"
  else fire_agg$Level[i] <- "not multiple"
}
############
## Police ##
###########
police_comp <- all_comp
police_comp$Department <- trimws(police_comp$Department, "both")
police_comp <- subset(police_comp, Department == "Police Department")
police_agg <- aggregate(cbind(Count,Total.Incurred) ~ Injured.Worker + Department, police_comp, sum)
police_agg <- police_agg[order(-police_agg$Count, -police_agg$Total.Incurred, police_agg$Injured.Worker),]

## get percentage of claims and costs for each staff ##
Pcount_pct <- round(prop.table(police_agg$Count)*100, 3)
Pincurred_pct <- round(prop.table(police_agg$Total.Incurred)*100, 3)
police_agg <- cbind(police_agg, Pcount_pct, Pincurred_pct)

## assign multiple if more than 1 claim
for (i in 1:length(police_agg$Count)){
  if(police_agg$Count[i] > 1)
    police_agg$Level[i] <- "multiple"
  else police_agg$Level[i] <- "not multiple"
}

############
## Police ##
###########
dpi_comp <- all_comp
dpi_comp$Department <- trimws(dpi_comp$Department, "both")
dpi_comp <- subset(dpi_comp, Department == "Department of Public Improvements" | Department == "DPI")
dpi_agg <- aggregate(cbind(Count,Total.Incurred) ~ Injured.Worker + Department, dpi_comp, sum)
dpi_agg <- dpi_agg[order(-dpi_agg$Count, -dpi_agg$Total.Incurred, dpi_agg$Injured.Worker),]

## get percentage of claims and costs for each staff ##
DPIcount_pct <- round(prop.table(dpi_agg$Count)*100, 3)
DPIincurred_pct <- round(prop.table(dpi_agg$Total.Incurred)*100, 3)
dpi_agg <- cbind(dpi_agg, DPIcount_pct, DPIincurred_pct)

## assign multiple if more than 1 claim
for (i in 1:length(dpi_agg$Count)){
  if(dpi_agg$Count[i] > 1)
    dpi_agg$Level[i] <- "multiple"
  else dpi_agg$Level[i] <- "not multiple"
}


### write staff report of claims and incurred costs ###
fileName <- "C:/Users/tsink/CovStat Projects/Human Resources/WorkersCompensation/Staff Report Data/WrksCompStaffReport.xlsx"
write.xlsx(comp_agg, fileName, sheetName = "All_Since_2010", row.names = FALSE)
write.xlsx(fire_agg, fileName, sheetName = "Fire_Since2010", row.names = FALSE, append = TRUE)
write.xlsx(police_agg, fileName, sheetName = "Police_Since2010", row.names = FALSE, append = TRUE)
write.xlsx(dpi_agg, fileName, sheetName = "DPI_Since2010", row.names = FALSE, append = TRUE)
write.xlsx(all_comp, fileName, sheetName = "AllDetails2010", row.names = FALSE, append = TRUE)

#######################
## Email Data Report ##
#######################
library(RDCOMClient)
## init com api
OutApp <- COMCreate("Outlook.Application")
## create an email 
outMail = OutApp$CreateItem(0)
## configure  email parameter 
outMail[["To"]] = ""
outMail[["Cc"]] = "tsink@covingtonky.gov"
outMail[["subject"]] = "Workers Comp Staff Report <Sent Only To Stacey>"
outMail[["body"]] = "This is an automated email to distribute the staff report for workers compensation.\n
This report breaks down the count of claims and total costs incurred per Covington staff since 2010 for Police, Fire, and DPI departments. 
Percentages of all claims and incurred costs per staff are also provided."
outMail[["Attachments"]]$Add("C:/Users/tsink/CovStat Projects/Human Resources/WorkersCompensation/Staff Report Data/WrksCompStaffReport.xlsx")
## send it                     
outMail$Send()



## Drop name and address
comp1$Injured.Worker <- NULL
comp1$Name <- NULL
comp1$Risk.Location.Address <- NULL

## Write for Tableau ##
write.csv(comp1,"U:/CityWide Performance/CovStat/CovStat Projects/Human Resources/WorkersCompensation/COMPENSATION.csv", row.names = FALSE)

## Write for CovStat ##
write.xlsx(comp1,"O:/AllUsers/CovStat/Data Portal/Repository/Data/HumanResources/WorkersCompensation.xlsx", row.names = FALSE)

