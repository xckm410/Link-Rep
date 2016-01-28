## incluede the packages in need
library( readxl )
library( dplyr )
library( solidearthtide )

## clear variables
rm( list = ls() )

###############################################################################
## read whole display table
###############################################################################
# set data source path
input_path <- "C:/Users/Jinwei.Xu/Desktop/from Stephen/28 Jan 2016/Generators/"

# read "udp" data
display_data_name <- "udp"
data_file_name <- dir( path = input_path, recursive = TRUE, ignore.case = TRUE, pattern = display_data_name )
udp_data <- read.csv( file = paste( input_path, data_file_name, sep = "" ), header = TRUE, check.names = FALSE )

# read "custom" data
display_data_name <- "custom"
data_file_name <- dir( path = input_path, recursive = TRUE, ignore.case = TRUE, pattern = display_data_name )
custom_data <- read.csv( file = paste( input_path, data_file_name, sep = "" ), header = TRUE, check.names = FALSE )

## do filtering to "Completed Application" (preserve "1')
completed_application_value_one_index <- which( udp_data$"Completed Application" == 1 )
filtered_udp_data <- udp_data[ completed_application_value_one_index, ]

###############################################################################
## generate correspondence between "custom" and "udp" data
###############################################################################
# get "custom" campaign names
custom_campaign_names <- as.character( custom_data$Campaign )

# get "udp" campaign names
udp_campaign_names <- as.character( filtered_udp_data$Campaign )


