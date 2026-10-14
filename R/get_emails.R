#' Get Most Recent Emails from Outlook Folder
#'
#' This function retrieves emails from an Outlook folder, with option to filter by given number of most recent emails, emails only after a certain date, emails with a certain subject, or all of the above.
#'
#' @param most.recent.n An integer specifying the number of most recent emails to retrieve. Default is NULL, which retrieves all emails.
#' @param after.this.date A POSIXct date/time. Only emails newer than this date/time will be returned. Default is NULL which retrieves all emails.
#' @param subject A character string with the subject line to be searched for. Default is NULL which returns emails with any subject.
#' @param subject.exact.match Indicates whether the subject line search will look for exact matches only, or whether it will check if the subject string is in the email's subject. Default is "FALSE" which indicates do not look for exact match.
#' @param folder.name Name of folder within account in which to search. Default is "Inbox".
#' @param account.name Name of account in which to search. Default is NULL, which searches in the user's default Outlook account.
#'
#' @return A list of COM objects representing the most recent emails in the Outlook inbox.
#'
#' @importFrom purrr map
#'
#' @examples
#' \dontrun{
#' recent_emails <- get_emails(most.recent.n = 10)
#' emails_sent_after <-  get_emails(account.name = "me@example.com", folder.name = "Sent Items", after.this.date = as.POSIXct("2000-12-31 10:30:00",tz="EST") )
#' }
#'
#' @export
get_emails <- function (most.recent.n = NULL,
                        after.this.date = NULL,
                        subject = NULL,
                        subject.exact.match = FALSE,
                        folder.name = "Inbox",
                        account.name = NULL) {

  if (folder.name == "Inbox" & is.null(account.name)) {
    folder <- email::inbox()  # Explicitly state that it's from the same package
  } else {
    folder <- email::outlook_folder(folder.name = folder.name, account.name = account.name)
  }

  if (is.null(folder)) {
    stop("Failed to get the folder. Make sure Outlook is installed and running.")
  }

  email_handle <- folder$Items()
  if (is.null(email_handle)) {
    stop("Failed to get email handle.")
  }

  # Validate inputs
  if (!is.null(most.recent.n) & (!is.numeric(most.recent.n) || most.recent.n <= 0)) {
    stop("Invalid value for most.recent.n. Please input a positive integer.")
  }
  if (!is.null(after.this.date) & !('POSIXct' %in% class(after.this.date))) {
    stop("Invalid date for after.this.date. Please use POSIXct format.")
  }
  if (typeof(subject.exact.match) != 'logical') {
    stop("Invalid input for subject_exact_match. Please use logical format.")
  }

  # Sort newest first. Outlook doesn't guarantee that the folder's index order
  # matches received order, so we sort explicitly. Sort() reorders this same
  # Items collection in place, so Item(1) is now the newest item.
  email_handle$Sort("[ReceivedTime]", TRUE)

  n_items <- email_handle$Count()
  if (is.null(most.recent.n)) most.recent.n <- n_items

  # Not everything in a folder is an email: read receipts, delivery reports,
  # task requests, etc. have no ReceivedTime and error if asked for one.
  # Return NA for those so they can be skipped.
  received_time <- function(item) {
    tryCatch(
      convert_comdate(item$ReceivedTime(), tz = 'EST'),
      error = function(e) NA
    )
  }

  accept_subject <- function(item) {
    if (is.null(subject)) return(TRUE)
    item_subject <- tryCatch(item$Subject(), error = function(e) NA_character_)
    isTRUE(grepl(subject, item_subject, fixed = subject.exact.match))
  }

  # Walk from newest to oldest, stopping once we have enough emails or hit one
  # older than after.this.date (everything after it is older still).
  emails <- list()
  index <- 1
  while (index <= n_items && length(emails) < most.recent.n) {
    item <- email_handle$Item(index)
    index <- index + 1

    item_date <- received_time(item)
    if (is.na(item_date)) next
    if (!is.null(after.this.date) && item_date <= after.this.date) break
    if (!accept_subject(item)) next

    emails <- c(emails, list(item))
  }

  # Return oldest to newest, matching previous behaviour
  return(rev(emails))
}
