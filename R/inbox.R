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
    Outlook <- COMCreate(com.object)
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
