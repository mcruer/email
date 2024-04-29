#' @import RDCOMClient
NULL

#' Get the Inbox Folder from Outlook
#'
#' This function initializes a COM object for Outlook and returns the Inbox folder.
#'
#' @param com.object A character string specifying the COM object for Outlook. Default is "Outlook.Application".
#' @param namespace.name A character string specifying the namespace for Outlook. Default is "MAPI".
#'
#' @return A COM object representing the Inbox folder in Outlook.
#'
#' @examples
#' \dontrun{
#' inbox_folder <- inbox()
#' }
#'
#' @export
inbox <-
  function (com.object = "Outlook.Application",
            namespace.name = "MAPI") {
    # Initialize Outlook and get Namespace
    Outlook <- RDCOMClient::COMCreate(com.object)
    if (is.null(Outlook)) {
      stop("Failed to create COM object. Make sure Outlook is installed and running.")
    }

    namespace <- Outlook$GetNameSpace(namespace.name)
    if (is.null(namespace)) {
      stop("Failed to get namespace. Make sure the namespace name is correct.")
    }

    # Define constant for Inbox folder
    INBOX_FOLDER <- 6

    # Get the Inbox folder
    inbox <- namespace$GetDefaultFolder(INBOX_FOLDER)

    return (inbox)
  }

#' Get any Folder from Outlook
#'
#' This function initializes a COM object for Outlook and returns any folder.
#'
#' @param folder.name Name of Outlook folder. Default is "Inbox".
#' @param account.name Name of Outlook account as it appears in Outlook. Default is NULL, which uses user's default Outlook account.
#' @param com.object A character string specifying the COM object for Outlook. Default is "Outlook.Application".
#' @param namespace.name A character string specifying the namespace for Outlook. Default is "MAPI".
#'
#' @return A COM object representing the folder in Outlook.\
#'
#' @import RDCOMClient
#'
#' @examples
#' \dontrun{
#' folder <- outlook_folder(folder.name = "Sent Items")
#' folder <- outlook_folder(folder.name = "Inbox", account.name = "myOtherAccount@example.com")
#' }
#'
#' @export
outlook_folder <- function (folder.name = "Inbox",
                            account.name = NULL,
                            com.object = "Outlook.Application",
                            namespace.name = "MAPI") {
    # Initialize Outlook and get Namespace
    Outlook <- RDCOMClient::COMCreate(com.object)
    if (is.null(Outlook)) {
      stop("Failed to create COM object. Make sure Outlook is installed and running.")
    }

    namespace <- Outlook$GetNameSpace(namespace.name)
    if (is.null(namespace)) {
      stop("Failed to get namespace. Make sure the namespace name is correct.")
    }

    #Fill in default account
    if (is.null(account.name)) {
      account.name = namespace$Accounts()$Item(1)$DisplayName()
    }
    if (is.null(account.name)) {
      stop("Couldn't find Outlook account!")
    }

    #Get into the account first:
    account <- namespace$Folders(account.name)

    #Then into the specific folder
    folder <- account$Folders(folder.name)

    return (folder)
  }
