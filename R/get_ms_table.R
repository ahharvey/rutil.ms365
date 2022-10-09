#' Fetch experience from OneDrive database
#'
#' Performs a GET request to the Microsoft Graph API. 
#' Returns a tibble containing the data within the specified
#' xlsx table.
#'
#' @param dir R6 object of class ms_drive representing the target directory
#' @param file String filename of xlsx source
#' @param table String MS Graph ID for the Excel table within the xlsx file
#' 
#' @export
get_ms_table <- function(dir = NULL, file = NULL, table = NULL) {

    # get the workbook matching the filename from src dir
    file <- dir$get_item(file)

    # build the API endpoint string
    table <- glue::glue("workbook/tables('{table}')/columns")

    # peform a custom GET operation to return
    # the given tables colums as JSON
    json <- file$do_operation(table)

    # build an empty list to receive the table items
    tbl_list <- list()

    # iterate through the JSON response and insert each item into exp_list
    for (i in json$value) {
        # the first item is the column name, the remainder are the values
        tbl_list[[unlist(i$value[1])]] <- unlist(i$value[-1])
    }

    # return tbl_list as a tibble
    dplyr::as_tibble(tbl_list)

}