#' Get Most Recent Emails from Outlook Inbox
#'
#' This function retrieves the most recent emails from the Outlook inbox.
#'
#' @param most.recent.n An integer specifying the number of most recent emails to retrieve. Default is NULL, which retrieves all emails.
#'
#' @return A list of COM objects representing the most recent emails in the Outlook inbox.
#'
#' @importFrom purrr map
#'
#' @examples
#' \dontrun{
#' recent_emails <- get_emails(most.recent.n = 10)
#' }
#'
#' @export
get_emails <- function (most.recent.n = NULL) {
  inbox <- email::inbox()  # Explicitly state that it's from the same package
  if (is.null(inbox)) {
    stop("Failed to get the inbox. Make sure Outlook is installed and running.")
  }

  email_handle <- inbox$Items()
  if (is.null(email_handle)) {
    stop("Failed to get email handle.")
  }

  n_emails <- email_handle$Count()
  if (is.null(most.recent.n)) most.recent.n <- n_emails

  if (!is.numeric(most.recent.n) || most.recent.n <= 0 || most.recent.n > n_emails) {
    stop("Invalid value for most.recent.n.")
  }

  emails <- purrr::map((n_emails - most.recent.n + 1):n_emails, ~email_handle$Item(.), .progress = "text")

  return(emails)
}





