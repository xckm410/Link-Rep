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
  new_name_list <- old_name_list[ -invalid_file_name_index ]
  return( new_name_list )
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
################# function parts end ################# 

## scan whole path
input_path <- "C:/Users/Jinwei.Xu/Desktop/Link's Job Task on 19 Jan 2016/Link/Sources/"
generator_file_name_list <- dir( path = input_path, recursive = TRUE, ignore.case = TRUE, pattern = "generator" )
generator_file_name_list <- remove_us_dollar_symbol_file_name_func( generator_file_name_list )

## input "custom" or "udp"
target_sheet_name <- "custom"

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
  
  # import selected sheet name
  selected_sheet <- read_excel( source_table_complete_path, sheet = selected_sheet_name )
  
  # judge valid colume name
  temp_col_name_list <- colnames( selected_sheet )
  temp_valid_colname_index <- valid_colname_index_func( temp_col_name_list )
  if( length( temp_valid_colname_index ) < length( temp_col_name_list ) )
  {
    # invalid one
    col_name_list <- as.character( selected_sheet[ 1, ] )
  }else
  {
    # valid ones
    col_name_list <- colnames( selected_sheet )
  }  
  
  # get aggregated col name list  
  unique_col_name <- intersect( unique_col_name, tolower( col_name_list ) )
  
  # get all col name list
  all_col_name_list[ file_name_index ] <- list( col_name_list )
  
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
  
  # import selected sheet name
  selected_sheet <- read_excel( source_table_complete_path, sheet = selected_sheet_name )  
  
  # print out the current excel table sheet
  print( generator_file_name_list[ file_name_index ] )
  
  # generate new table
  temp_col_name_list <- all_col_name_list[[ file_name_index ]]
  data_start_row <- detect_data_range_start_row_func( selected_sheet, temp_col_name_list )
  
  # process "date" column
  temp_date_number_value <- selected_sheet[ data_start_row : dim( selected_sheet )[ 1 ], date_col_index_list[[ file_name_index ]] ]
  for( date_index in data_start_row : dim( selected_sheet )[ 1 ] )
  {
    temp_real_date_value <- as.Date( as.numeric( temp_date_number_value[ date_index, 1 ] ), origin = "1900-01-01" )
  }
  
  
  
  selected_sheet[ data_start_row : dim( selected_sheet )[ 1 ], date_col_index_list[[ file_name_index ]] ] <- as.Date( as.numeric( selected_sheet[ data_start_row : dim( selected_sheet )[ 1 ], date_col_index_list[[ file_name_index ]] ] ), origin = "1899-12-30" )
  
  # copy and paste
  temp_new_table <- selected_sheet[ data_start_row : dim( selected_sheet )[ 1 ], overlapping_selected_col_index_list[[ file_name_index ]] ]
  colnames( temp_new_table ) <- rep( "", length( overlapping_col_name ) )
  new_table <- rbind( new_table, temp_new_table ) 

}

