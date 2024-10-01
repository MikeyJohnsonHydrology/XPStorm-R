# A simple scrip to read ".out" files from XPStorm.
# R scrips can be saved to the working directory wiht the XPSorm ".out" files.
# Data to be saved to MS Excel Workbooks.
# By: Mikey Johnson, mikey.yosemite@gmail.com

### Set working directory ######################################################
sfl <- dirname(rstudioapi::getActiveDocumentContext()$path)  # this is the source file location
setwd(sfl)
getwd()

### Loading Libraries ##########################################################
library(dplyr)
library(tidyr)
library(stringr)
library(writexl)

### Listing all the ".out" file in the current working directory ###############
out_file_names <- list.files(pattern=".out")
out_file_names
length(out_file_names)

### Start of the Loop to Produce all of the Outfiles ###########################
for (i in 1:length(out_file_names)){


### Read ".out" file ###########################################################
OF_dat <- read.delim(out_file_names[i], header=FALSE)

### Filter out the table information

table_e9_start <- which(OF_dat==" |        Table E9 - JUNCTION SUMMARY STATISTICS        |")
table_e10_start <- which(OF_dat==" |        Table E10 - CONDUIT SUMMARY STATISTICS        |")
table_e13_start <- which(OF_dat==" | Table E13. Channel losses(H), headwater depth (HW), tailwater |")
table_e13a_start <- which(OF_dat==" | Table E13a. CULVERT ANALYSIS CLASSIFICATION,             |")

table_e9_data <- data.frame(V1=OF_dat[(table_e9_start+9):(table_e10_start-2),1])
table_e13_data <- data.frame(V1=OF_dat[(table_e13_start+7):(table_e13a_start-2),1])

### Formatting Table E9 Data (Node Data) #######################################
df <- table_e9_data

# Function to clean and split data
clean_and_split <- function(text) {
  # Trim leading and trailing whitespace
  text <- str_trim(text)
  
  # Replace multiple spaces with a single space
  text <- str_replace_all(text, "\\s+", " ")
  
  # Split the text by space
  parts <- str_split(text, " ")[[1]]
  
  # Determine where the first numeric value starts
  # We assume the first numeric value starts after the code
  idx_first_numeric <- which(str_detect(parts, "^\\d+\\.\\d+"))[1]
  
  # Separate the code and values
  code <- paste(parts[1:(idx_first_numeric - 1)], collapse = " ")
  values <- parts[idx_first_numeric:length(parts)]
  
  # Combine the results
  return(c(code, values))
}

split_data <- df %>%
  rowwise() %>%
  mutate(split_data = list(clean_and_split(V1))) %>%
  unnest_wider(split_data, names_sep = "_")

# Standardize data by padding shorter rows with NA
max_cols <- 12  # Adjust based on the maximum number of expected columns

standardized_data <- split_data %>%
  mutate(across(starts_with("split_data_"), ~ replace_na(., NA))) %>%
  # Rename columns appropriately
  rename_with(~ paste0("Col_", seq_along(.)))

# Set specific column names
colnames(standardized_data) <- c("Original_Text",
                                 "Junction Name",
                                 "Ground Elevation feet",
                                 "Uppermost PipeCrown Elevation feet",
                                 "Maximum Junction Elevation feet",
                                 "Time of Occurence Hr.",
                                 "Time of Occurence Min.",
                                 "Feet of Surcharge at Max Elevation",
                                 "Freeboard of node feet",
                                 "Maximum Junction Area ft^2",
                                 "Maximum Gutter Depth feet",
                                 "Maximum Gutter Width feet",
                                 "Maximum Gutter Velocity ft/s")

# Final Data
Final_data <- standardized_data[,2:13]

# Write the data to an Excel file
write_xlsx(Final_data,path = paste("Table E9 ",
                                   gsub('.{4}$','',out_file_names[i]),
                                   ".xlsx",
                                   sep=""))


### Formatting Table E13 Data (Link Results) ###################################
df2 <- table_e13_data

split_data2 <- df2 %>%
  rowwise() %>%
  mutate(split_data = list(clean_and_split(V1))) %>%
  unnest_wider(split_data, names_sep = "_")

# Set specific column names
colnames(split_data2) <- c("Original_Text",
                           "Conduit Name",
                           "Conduit Name",
                           "Maximum Flow",
                           "Head Loss",
                           "Friction Loss",
                           "Critical Depth",
                           "Normal Depth",
                           "HW Elevation",
                           "TW Elevation",
                           "Max",
                           "Flow",
                           "Maximum Gutter Width feet",
                           "Maximum Gutter Velocity ft/s")

# Final Data
Final_data2 <- split_data2[,c(2,4:10)]

# Write the data to an Excel file
write_xlsx(Final_data2,path = paste("Table E13 ",
                                    gsub('.{4}$','',out_file_names[i]),
                                    ".xlsx",
                                    sep=""))

### Removing the data for the next 'i' loop
rm(OF_dat,
   table_e9_start,table_e10_start,table_e13_start,table_e13a_start,
   table_e9_data,table_e13_data,
   df,split_data,standardized_data,Final_data,
   df2,split_data2,Final_data2)



### End of loop ################################################################
}
