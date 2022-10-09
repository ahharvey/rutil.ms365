#' Get the user's OneDrive for Business drive
#'
#' Authenticates with the Microsoft Graph API, and returns
#' an object representing the OneDrive of the given user.
#'
#' @param tennant String indicating the Sharepoint tennant
#' @param app String Azure app id
#' @param pwd String Azure app secret token
#' @param email String MS email for the OneDrive user
#'
#' @export
get_ms_drive <- function(tennant = NULL, app = NULL,
    pwd = NULL, email = NULL) {

    # get an Azure token without interaction
    # (i.e. suitable for headless unattended)
    token <- AzureGraph::create_graph_login(tennant, app = app,
        password = pwd, auth_type = "client_credentials")

    # get the user matching the provided name
    user <- token$get_user(email)

    # get the OneDrive of the given user
    drive <- user$get_drive()

    drive

}