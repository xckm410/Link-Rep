## incluede the packages in need
library( readxl )
library( dplyr )

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

################# find all generators under specific path (start) #################
input_path <- "C:/Users/Jinwei.Xu/Desktop/Link's Job Task on 19 Jan 2016/Link/Sources/"
generator_file_name_list <- dir( path = input_path, recursive = TRUE, ignore.case = TRUE, pattern = "generator" )
generator_file_name_list <- remove_us_dollar_symbol_file_name_func( generator_file_name_list )
################# find all generators under specific path (end) #################

## input "custom" or "udp"
target_sheet_name <- "udp"

## select col names 
selected_col_names <- c( "date", "campaign", "publisher", "site", "creative message", "action type", "bank", "portfolio", "product", "device", 
                         "completed application", "actual publisher", "actual creative name", "actual product message",
                         "creative product + message", "application decision", "approval rate", "approved applications",
                         "activation rate", "approved activated accounts", "revenue muliplier", "vos multipler", "total revenue",
                         "total voS" )

## set empty new table
new_table <- NULL

########################################################
## import source table in a batch processing way (start)
for( file_name_index in 1 : length( generator_file_name_list ) )
{
  
  ## set source table name 
  source_table_name <- generator_file_name_list[ file_name_index ]
  
  ## get table sheet names
  source_table_complete_path <- paste( input_path, source_table_name, sep = "" )
  source_table_sheet_name <- excel_sheets( source_table_complete_path )
#   print( source_table_sheet_name )
#   print( "=========================" )

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
  }
  else
  {
    # valid ones
    col_name_list <- colnames( selected_sheet )
  }  
#   print( col_name_list )
#   print( "==================================================" )

  # get bank name and portfolio name from the generator's name
  temp_source_table_name <- gsub( "-", " ", source_table_name )
  temp_split_source_table_name <- strsplit(temp_source_table_name, " ")[[1]]
  temp_valid_split_source_table_name_index <- valid_colname_index_func( temp_split_source_table_name )
  bank_name <- temp_split_source_table_name[ temp_valid_split_source_table_name_index[ 1 ] ]
  portfolio_name <- temp_split_source_table_name[ temp_valid_split_source_table_name_index[ 2 ] ]

  # get selected colom index
  selected_col_index_list <- NULL
  for( index in 1 : length( selected_col_names ) )
  {
    selected_col_index_list <- c( selected_col_index_list, which( tolower( col_name_list ) == selected_col_names[ index ] ) )
  }
  
  # generate new table
  data_start_row <- detect_data_range_start_row_func( selected_sheet, col_name_list )
  temp_new_table <- selected_sheet[ data_start_row : dim( selected_sheet )[ 1 ], selected_col_index_list ]
  # names( temp_new_table ) <- NULL
  new_table <- rbind( new_table, temp_new_table ) 
    
#   ##############################################
#   # for processing "custom data"
#   if( tolower( target_sheet_name ) == "custom" )  
#   {
#     
#     # get selected colom index
#     selected_col_index_list <- NULL
#     for( index in 1 : length( selected_col_names ) )
#     {
#       selected_col_index_list <- c( selected_col_index_list, which( tolower( col_name_list ) == selected_col_names[ index ] ) )
#     }
#     
# #     # generate new table
# #     new_custom_table <- selected_sheet[ ,selected_col_index_list ]
# #     new_custom_table <- cbind( new_custom_table, rep( bank_name, dim( selected_sheet )[ 1 ] ), rep( portfolio_name, dim( selected_sheet )[ 1 ] ) )
# #     colnames( new_custom_table )[ dim( new_custom_table )[2] - 1 ] <- "Bank"
# #     colnames( new_custom_table )[ dim( new_custom_table )[2] ] <- "Portfolio"
#     
#     # generate new table
#     data_start_row <- detect_data_range_start_row_func( selected_sheet, col_name_list )
#     new_table <- selected_sheet[ data_start_row : dim( selected_sheet )[ 1 ], selected_col_index_list ]
#     
#   }
#   
#   ##############################################
#   # for processing "UDP data"
#   if( tolower( target_sheet_name ) == "udp" )  
#   {
#     
#     # get selected colom index
#     selected_col_index_list <- NULL
#     for( index in 1 : length( selected_col_names ) )
#     {
#       selected_col_index_list <- c( selected_col_index_list, which( tolower( col_name_list ) == selected_col_names[ index ] ) )
#     }
#     
#     # generate new table
#     data_start_row <- detect_data_range_start_row_func( selected_sheet, col_name_list )
#     new_table <- selected_sheet[ data_start_row : dim( selected_sheet )[ 1 ], selected_col_index_list ]
#     
#   }  
    
}


  


