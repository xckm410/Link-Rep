## incluede the packages in need
library( readxl )
library( dplyr )
library( solidearthtide )

## clear variables
rm( list = ls() )

################# function parts start ################# 
valid_colname_index_func <- function( name_list )
{
  valid_colname_index <- NULL
  for( name_index in 1 : length( name_list ) )
  {
    if( nchar( name_list[ name_index ] ) > 0 )
    {
      valid_colname_index <- c( valid_colname_index, name_index )
    }
  }
  return( valid_colname_index )
}

remove_us_dollar_symbol_file_name_func <- function( old_name_list )
{
  invalid_file_name_index <- grep( "/~", old_name_list )
  if( length( invalid_file_name_index ) )
  {
    new_name_list <- old_name_list[ -invalid_file_name_index ]
    return( new_name_list )
  }
  else
  {
    new_name_list <- old_name_list
    return( new_name_list )
  }
}

detect_data_range_start_row_func <-function( selected_sheet, col_name_list )
{
  valid_indicator <- NULL
  for( row in 1 : 10 )
  {
    valid_indicator <- rbind( valid_indicator, cbind( row, sum( as.character( selected_sheet[ row, ] ) == col_name_list, na.rm = TRUE ) ) )
  }
  col_name_row_index <- which( valid_indicator[ ,2 ] == length( col_name_list ) )
  if( length( col_name_row_index ) )
  {
    return( valid_indicator[ 1, col_name_row_index ] + 1 )
  }
  else
  {
    return( 1 )
  }
}

simpleCap <- function( x ) 
{
  s <- strsplit( x, " " )[[1]]
  paste( toupper( substring( s, 1, 1 ) ), substring( s, 2 ), sep = "", collapse = " " )
}

robust_process_date_func <- function( x )
{
  
  # get summary of date x
  date_x_summary <- summary( x )
  
  # get summary structure size
  summary_structure_size <- length( date_x_summary )
  
  # size = 6
  if( summary_structure_size == 6 )
  {
    temp_x <- date_x_summary[1]
    temp_date_string <- strsplit( temp_x, ":" )[[1]][ length( strsplit( temp_x, ":" )[[1]] ) ]
    temp_date_string <- gsub( " ", "", temp_date_string )
    return( temp_date_string )
  }
  else
  {
    temp_x <- as.numeric( x )
    temp_date_string <- as.character( as.Date( temp_x, origin = "1900-01-01" ) )
    return( temp_date_string )
  }
  
}
################# function parts end ################# 

## scan whole path
input_path <- "C:/Users/Jinwei.Xu/Desktop/from Stephen/28 Jan 2016/Generators/"
generator_file_name_list <- dir( path = input_path, recursive = TRUE, ignore.case = TRUE, pattern = "generator" )
generator_file_name_list <- remove_us_dollar_symbol_file_name_func( generator_file_name_list )

## input "custom" or "udp"
target_sheet_name <- "udp"

## select col names 
selected_col_names <- c( "date", "campaign", "publisher", "site", "creative message", "action type", "bank", "portfolio", "product", "device", 
                         "completed application", "actual publisher", "actual creative name", "actual product message",
                         "creative product + message", "application decision", "approval rate", "approved applications",
                         "activation rate", "approved activated accounts", "revenue muliplier", "vos multipler", "total revenue",
                         "total voS" )

#######################################################
## set unique col name list
unique_col_name <- selected_col_names

## set list for all individual col names
all_col_name_list <- vector( "list", length( generator_file_name_list ) )

## import source table in a batch processing way (start)
for( file_name_index in 1 : length( generator_file_name_list ) )
{
  
  # set source table name 
  source_table_name <- generator_file_name_list[ file_name_index ]
  
  # get table sheet names
  source_table_complete_path <- paste( input_path, source_table_name, sep = "" )
  source_table_sheet_name <- excel_sheets( source_table_complete_path )
  
  # get rough sheet id (could be more than one)
  target_sheet_id <- grep( target_sheet_name, source_table_sheet_name, ignore.case = TRUE )
  
  # should include "data" string
  selected_target_sheet_id <- grep( "data", source_table_sheet_name[ target_sheet_id ], ignore.case = TRUE )
  
  # get selected sheet name
  selected_sheet_name <- source_table_sheet_name[ target_sheet_id[ selected_target_sheet_id ] ]
  
  # judge whether there is no "selected sheet name" or not
  if( length( selected_sheet_name ) )
  {
  
    # import selected sheet name
    selected_sheet <- read_excel( source_table_complete_path, sheet = selected_sheet_name )
  
    # judge valid colume name
    temp_col_name_list <- colnames( selected_sheet )
    temp_valid_colname_index <- valid_colname_index_func( temp_col_name_list )
    if( length( temp_valid_colname_index ) < length( temp_col_name_list ) )
    {
      
      # invalid one
      col_name_list <- as.character( selected_sheet[ 1, ] )
  
    }
    else
    {
    
      # valid ones
      col_name_list <- colnames( selected_sheet )
      
    }  
  
    # get aggregated col name list  
    unique_col_name <- intersect( unique_col_name, tolower( col_name_list ) )
  
    # get all col name list
    all_col_name_list[ file_name_index ] <- list( col_name_list )
    
  }
  
}

#######################################################
# find overlapping col names within unique col name list
overlapping_col_name <- intersect( tolower( unique_col_name ), tolower( selected_col_names ) )

# generate overlapping selected col index list
overlapping_selected_col_index_list <- vector( "list", length( generator_file_name_list ) )
date_col_index_list <- vector( "list", length( generator_file_name_list ) )
for( file_name_index in 1 : length( generator_file_name_list ) )
{
  overlapping_selected_col_index_list[ file_name_index ] <- list( match( overlapping_col_name, tolower( all_col_name_list[[ file_name_index ]] ) ) )
  date_col_index_list[ file_name_index ] <- list( match( "date", tolower( all_col_name_list[[ file_name_index ]] ) ) )
}

#######################################################
# set an empty new table
new_table <- NULL
data_row_sum <- 0
Date_row_sum <- 0

for( file_name_index in 1 : length( generator_file_name_list ) )
{  
  
  # set source table name 
  source_table_name <- generator_file_name_list[ file_name_index ]
  
  # get table sheet names
  source_table_complete_path <- paste( input_path, source_table_name, sep = "" )
  source_table_sheet_name <- excel_sheets( source_table_complete_path )
  
  # get rough sheet id (could be more than one)
  target_sheet_id <- grep( target_sheet_name, source_table_sheet_name, ignore.case = TRUE )
  
  # should include "data" string
  selected_target_sheet_id <- grep( "data", source_table_sheet_name[ target_sheet_id ], ignore.case = TRUE )
  
  # get selected sheet name
  selected_sheet_name <- source_table_sheet_name[ target_sheet_id[ selected_target_sheet_id ] ]
  
  # judge whether there is no "selected sheet name" or not
  if( length( selected_sheet_name ) )
  {
  
    # import selected sheet name
    selected_sheet <- read_excel( source_table_complete_path, sheet = selected_sheet_name )  
  
    # generate new table
    temp_col_name_list <- all_col_name_list[[ file_name_index ]]
    data_start_row <- detect_data_range_start_row_func( selected_sheet, temp_col_name_list )
  
    ####################### print out intermediate status (start) ####################### 
    # print out the current excel table sheet
    print( "============================================================" )
    print( generator_file_name_list[ file_name_index ] )
    print( paste( "Data range starts from ", data_start_row, " row", sep = "" ) )
    print( length( data_start_row : dim( selected_sheet )[ 1 ] ) )
    print( "============================================================" )
    data_row_sum <- data_row_sum + length( data_start_row : dim( selected_sheet )[ 1 ] )
    ####################### print out intermediate status (end) #######################
  
    ####################### process date column (start) #######################
    # process date column
    temp_date_number_value <- selected_sheet[ data_start_row : dim( selected_sheet )[ 1 ], date_col_index_list[[ file_name_index ]] ]
  
    # judge whether it is a date
    temp_date_sequence_value <- NULL
    for( row_index in 1 : dim( temp_date_number_value )[ 1 ] )
    {
      temp_date_sequence_value[ row_index ] <- robust_process_date_func( temp_date_number_value[ row_index, 1 ] )
    }
  
    # completely transfer into date sequence
    Date_row_sum <- Date_row_sum + length( temp_date_sequence_value )
  
    # copy date sequence into table sheet
    selected_sheet[ data_start_row : dim( selected_sheet )[ 1 ], date_col_index_list[[ file_name_index ]] ] <- temp_date_sequence_value
    ####################### process date column (end) #######################
  
    # copy and paste
    temp_new_table <- selected_sheet[ data_start_row : dim( selected_sheet )[ 1 ], overlapping_selected_col_index_list[[ file_name_index ]] ]
    colnames( temp_new_table ) <- rep( "", length( overlapping_col_name ) )
    new_table <- rbind( new_table, temp_new_table ) 
    
  }
  else
  {
    
    # show "warning - no such table sheet exists"
    print( "============================================================" )
    print( paste( "No ", toupper( target_sheet_name ), " data sheet exists in ", generator_file_name_list[ file_name_index ], sep = "" ) )
    print( "============================================================" )
    
  }
  
}

# print results
print( paste( "Data range row sum = ", data_row_sum, sep = "" ) )
print( paste( "Date row sum = ", Date_row_sum, sep = "" ) )

# put column names back to new table
colnames( new_table ) <- sapply( unique_col_name, simpleCap )

# write "new table" into an Excel CSV table
csv_table_file_name <- paste( input_path, "Whole_", toupper( target_sheet_name ), "_Data.csv", sep = "" )
write.csv( new_table, file = csv_table_file_name, na = "", row.names = FALSE )

