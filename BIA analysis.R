#load these libraries to manipulate the xlsx file
library(openxlsx)
library(readr)
library(readxl)
library(xlsx)
library(dplyr)

#read the file into an R dataframe
bia = read_xlsx("C:/Users/Elaine/Google Drive/BI analysis/PendaHealth_Data.xlsx", 
                sheet = "PatientPaidVisit", 
                range = "A1:K84365", 
                col_names = TRUE
                )
#view the dataframe
View(bia)

#view the structure of the dataframe
str(bia)
summary(bia)

#clean the dataset by replacing missing values

bia$`Visit Code`[is.na(bia$`Visit Code`)] = 0
bia$`Total invoice`[is.na(bia$`Total invoice`)] = 0
bia$`Payment Type`[is.na(bia$`Payment Type`)] = NA
bia$Gender[is.na(bia$Gender)] = NA
bia$Diagnosis[is.na(bia$Diagnosis)] = NA
bia$`New/Repeat`[is.na(bia$`New/Repeat`)] = NA
bia$`Visit location`[is.na(bia$`Visit location`)] = NA

#use the mutate_at() fuction to clean the dates
bia = bia %>% mutate_at(c("Visit date", "Date of Birth"), 
                        funs(ifelse(is.na(.), 
                                    yes = 0, 
                                    no = .
                                    )
                             )
                        )
#convert the 'Visit date' and 'Date of Birth' column values from numbers to date format

bia$`Date of Birth` = as.Date.POSIXct(bia$`Date of Birth`, 
                                      format = "%Y-%m-%d"
                                      )
#replace all typos of "tyfoid" to "typhoid" using gsub()
bia$Diagnosis = gsub(pattern = "tyfoid", 
                     replacement = "typhoid", 
                     x = bia$Diagnosis, 
                     ignore.case = TRUE
                     )

#overwrite the oldfile to a new-cleaned file

write.xlsx(bia, 
           file = "C:/Users/Elaine/Google Drive/BI analysis/PendaHealth_Data.xlsx", 
           sheetName = "clean BIA", 
           col.names = TRUE, 
           append = TRUE, 
           showNA = TRUE
           )


