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

  # Set up:
  n_emails <- n_valid_emails <- email_handle$Count()
  get_email <- folder$Items
  email_start_index = 1
  if (is.null(most.recent.n)) most.recent.n <- n_emails

  # DATE FILTER: Get index for where emails start to have date that is newer than after.this.date
  ##N2S: This assumes user wants more recent emails. But we should solidify this with a bisector method or something in case they're pulling 10 years' worth of emails
  if (!is.null(after.this.date)) {
    email_index <- n_emails
    email_date <- convert_comdate(get_email(n_emails)$ReceivedTime(), tz='EST')
    while (email_date > after.this.date) {
      email_index = email_index - 1
      if (email_index < 1) break
      email_date <- convert_comdate(get_email(email_index)$ReceivedTime(), tz='EST')
    }

    email_start_index = email_index + 1
    n_valid_emails = n_emails - email_start_index + 1
  }

  # SUBJECT FILTER: Make a helper function that checks if subject is what we want
  if (!is.null(subject)) {
    accept_subject <- function(email, index) {
      email_subject <- email::subject(email)
      if (grepl(subject, email_subject, fixed = subject.exact.match)) {
        return(index)
      }
      return(0)
    }

    indices_of_valid_subjects <- unlist(purrr::map(email_start_index:n_emails, ~accept_subject(get_email(.), .), .progress = "text"))
    indices_of_valid_subjects <- indices_of_valid_subjects[indices_of_valid_subjects!=0]

    n_valid_emails <- length(indices_of_valid_subjects)

  }

  # NUMBER FILTER: Get most recent number
  if (n_valid_emails < most.recent.n) {
    most.recent.n = n_valid_emails
  }

  # Get the emails:
  if (exists('indices_of_valid_subjects')) {
    emails <- purrr::map(tail(indices_of_valid_subjects, most.recent.n), ~email_handle$Item(.), .progress = "text")
  } else {
    emails <- purrr::map((n_emails - most.recent.n + 1):n_emails, ~email_handle$Item(.), .progress = "text")
  }

  return(emails)
}



