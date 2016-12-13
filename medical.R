
setwd("U:/CityWide Performance/Human Resources/Data/Medical Claims")

library("xlsx", lib.loc="~/R/win-library/3.2")
library("plyr", lib.loc="~/R/win-library/3.2")
library("dplyr", lib.loc="~/R/win-library/3.2")
library("tidyr", lib.loc="~/R/win-library/3.2.5")
library("reshape", lib.loc="~/R/win-library/3.2")
library("reshape2", lib.loc="~/R/win-library/3.2.5")
library("stringr", lib.loc="~/R/win-library/3.2")
library("zoo", lib.loc="~/R/win-library/3.2")
library("lubridate", lib.loc="~/R/win-library/3.2")

medical <- read.csv("Medical.csv", header=TRUE, na.strings = c("", NA), stringsAsFactors = FALSE)
medical$Type <- "Medical"


dental <-  read.csv("Dental.csv", header=TRUE, na.strings = c("", NA), stringsAsFactors = FALSE)
dental$Type <- "Dental"

fees <-  read.csv("Fees.csv", header=TRUE, na.strings = c("", NA), stringsAsFactors = FALSE)
fees$Type <- "Fees"

rx.fees <-  read.csv("RX Fees.csv", header=TRUE, na.strings = c("", NA), stringsAsFactors = FALSE)
rx.fees$Type <- "RX Fees"

sub.fees <-  read.csv("Subrogation Fees.csv", header=TRUE, na.strings = c("", NA), stringsAsFactors = FALSE)
sub.fees$Type <- "Subrogation Fees"

medical.claims <- do.call("rbind", list(medical, dental, fees, rx.fees, sub.fees))

####
####   Monthly Update   #####
october16.medical <- read.xlsx("hr_medical.October16.xlsx", sheetName = "Medical", as.data.frame = TRUE, header = TRUE)


medical.claims$TOT.CHK.AMT <- as.numeric(gsub("[\\$-]", "", medical.claims$TOT.CHK.AMT))
medical.claims$SBM.CHRGS <- as.numeric(gsub("[\\$-]", "", medical.claims$SBM.CHRGS))
medical.claims$COPAY <- as.numeric(gsub("[\\$-]", "", medical.claims$COPAY))
medical.claims$DED.AMT <- as.numeric(gsub("[\\$-]", "", medical.claims$DED.AMT))
medical.claims$COINS <- as.numeric(gsub("[\\$-]", "", medical.claims$COINS))

medical.claims[is.na(medical.claims)] <- 0

#medical.claims$FROM.DOS <- gsub(pattern = "/", "", medical.claims$FROM.DOS)
#medical.claims$FROM.DOS <- dmy(medical.claims$FROM.DOS)

###Medical claims
#medical_claims  <- read.xlsx("medicalclaimsFrom7-14 to 6-16.xlsx", sheetName="Medical Claims", colIndex=1:10, as.data.frame=TRUE, header=TRUE)
#medical_claims$COVG.CATE.DESC <- "MEDICAL CLAIMS"
#medical_claims <- medical_claims[c("GROUP..", "ACCT", "ACCT.NAME", "CLM.NUM", "FROM.DOS", "THRU.DOS", "CHK.ISS", "TOT.CHK.AMT", "COVG.CATE.DESC", 
                   #        "TAX.ID", "PROV.PAYEE")]
#names(medical_claims) <- c("GROUP..", "ACCT", "ACCT.NAME", "CLM.NUM", "FROM.DOS", "THRU.DOS", "CHK.ISS", "TOT.CHK.AMT", "COVG.CATE.DESC", 
                    #       "TAX.ID", "PROV.PAYEE")
#medical_claims <- na.omit(medical_claims)

###Dental
#dental  <- read.xlsx("medicalclaimsFrom7-14 to 6-16.xlsx", sheetName="Dental", colIndex=1:10, as.data.frame=TRUE, header=TRUE)
#dental$COVG.CATE.DESC <- "DENTAL"
#dental <- dental[c("GROUP..", "ACCT", "ACCT.NAME", "CLM.NUM", "FROM.DOS", "THRU.DOS", "CHK.ISS", "TOT.CHK.AMT", "COVG.CATE.DESC", "TAX.ID", "PROV.PAYEE")]
#names(dental) <- c("GROUP..", "ACCT", "ACCT.NAME", "CLM.NUM", "FROM.DOS", "THRU.DOS", "CHK.ISS", "TOT.CHK.AMT", "COVG.CATE.DESC", 
                     #      "TAX.ID", "PROV.PAYEE") 
#dental <- na.omit(dental)

###Ortho
#ortho <- read.xlsx("medicalclaimsFrom7-14 to 6-16.xlsx", sheetName="Ortho", colIndex=1:11, as.data.frame=TRUE, header=TRUE)  
#ortho <- rename(ortho, c("COVG.CATE.DESC"="COVG CATE DESC"))
#ortho <- na.omit(ortho)

###Fees
#fees  <- read.xlsx("medicalclaimsFrom7-14 to 6-16.xlsx", sheetName="Fees", colIndex=1:10, as.data.frame=TRUE, header=TRUE)
#fees$TAX.ID <- 0
#fees$Ta.ID <- as.numeric(fees$TAX.ID)
#fees <- fees[c("GROUP..", "ACCT", "ACCT.NAME", "CLM.NUM", "FROM.DOS", "THRU.DOS", "CHK.ISS", "TOT.CHK.AMT", "COVG.CATE.DESC", "TAX.ID", "PROV.PAYEE")]
#names(fees) <- c("GROUP..", "ACCT", "ACCT.NAME", "CLM.NUM", "FROM.DOS", "THRU.DOS", "CHK.ISS", "TOT.CHK.AMT", "COVG.CATE.DESC", 
                   #"TAX.ID", "PROV.PAYEE") 
#fees <- na.omit(fees)

###RX
#rx  <- read.xlsx("medicalclaimsFrom7-14 to 6-16.xlsx", sheetName="Rx & Rx Fees", colIndex=1:11, as.data.frame=TRUE, header=TRUE)
#rx <- na.omit(rx)

###Subrogation
#subrogation  <- read.xlsx("medicalclaimsFrom7-14 to 6-16.xlsx", sheetName="Subrogation Fees", colIndex=1:11, as.data.frame=TRUE, header=TRUE)
#subrogation <- na.omit(subrogation)


###Put them together
#hr_medical <- do.call("rbind", list(medical_claims, dental, ortho, fees, rx, subrogation))

write.csv(medical.claims, "U:/CityWide Performance/Human Resources/Data/ForReporting/hr_medical.csv", row.names=FALSE)

