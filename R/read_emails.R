.data <- NULL

# Convert COMDate ---------

#' Convert COMDate to POSIXct
#'
#' This function converts a COMDate to a POSIXct object.
#'
#' @param comdate A numeric COMDate.
#' @param tz A character string representing the time zone. Default is "UTC".
#'
#' @return A POSIXct object representing the date and time.
#' @export
#'
#' @examples
#' \dontrun{
#' comdate <- get_comdate()  # Assume this function retrieves a COMDate
#' convert_comdate(comdate)
#' }
convert_comdate <- function(comdate, tz = "UTC") {
  epoch <- as.POSIXct("1899-12-30", tz = tz)
  days <- floor(comdate)
  fraction_of_day <- comdate - days
  seconds <- round(fraction_of_day * 24 * 60 * 60)
  datetime <- epoch + days * 24 * 60 * 60 + seconds
  return(datetime)
}


# Attachment Count --------
#' Get the Number of Attachments in an Email
#'
#' This function returns the number of attachments in a given email object.
#'
#' @param email A COMIDispatch object representing an email.
#'
#' @return An integer representing the number of attachments, or NA if not found.
#' @export
#'
#' @examples
#' \dontrun{
#' email_obj <- get_email_object()  # Assume this function retrieves an email object
#' attachment_count(email_obj)
#' }
attachment_count <- function(email) {
  if (!inherits(email, "COMIDispatch")) {
    stop("Invalid email object.")
  }

  result <- tryCatch(
    email$Attachments()$Count(),
    error = function(e) {
      warning("No attachments found.")
      return(NA_integer_)
    }
  )

  return(result)
}

# Save Attachments --------
#' Save an Email's Attachments to Disk
#'
#' This function saves one or more attachments from a given email object to a
#' folder on disk, optionally filtering which attachments to save and renaming
#' them on the way out. Useful for automated workflows such as pulling a
#' specific file out of a recurring weekly email.
#'
#' @param email A COMIDispatch object representing an email.
#' @param path A character string giving the directory to save attachment(s)
#'   into. The directory must already exist.
#' @param attachment.pattern Optional character string (regex) used to filter
#'   attachments by file name, e.g. "\\\\.zip$" to only save zip files.
#'   Default is NULL, which matches all attachments on the email.
#' @param file_name Optional character vector of new file name(s) to save the
#'   matched attachment(s) as (e.g. "weekly_report.zip"). If supplied, its
#'   length must equal the number of matched attachments. Default is NULL,
#'   which keeps each attachment's original file name.
#' @param overwrite Logical, whether to overwrite a file already existing at
#'   the destination path. Default is TRUE.
#'
#' @return A character vector of the full file paths the attachment(s) were
#'   saved to, returned invisibly.
#' @importFrom purrr walk2
#' @export
#'
#' @examples
#' \dontrun{
#' email_obj <- get_emails(most.recent.n = 1, subject = "Weekly Report")[[1]]
#'
#' # Save every attachment using its original name
#' save_attachments(email_obj, path = "C:/reports")
#'
#' # Save just the zip attachment under a new name
#' save_attachments(
#'   email_obj,
#'   path = "C:/reports",
#'   attachment.pattern = "\\.zip$",
#'   file_name = "weekly_report.zip"
#' )
#' }
save_attachments <- function(email,
                              path,
                              attachment.pattern = NULL,
                              file_name = NULL,
                              overwrite = TRUE) {
  if (!inherits(email, "COMIDispatch")) {
    stop("Invalid email object.")
  }
  if (!dir.exists(path)) {
    stop("The specified path does not exist: ", path)
  }

  attachments <- email$Attachments()
  n_attachments <- tryCatch(attachments$Count(), error = function(e) 0L)

  if (n_attachments == 0) {
    warning("No attachments found.")
    return(invisible(character(0)))
  }

  indices <- seq_len(n_attachments)
  original_names <- vapply(indices, function(i) attachments$Item(i)$FileName(), character(1))

  if (!is.null(attachment.pattern)) {
    matched <- grepl(attachment.pattern, original_names)
    indices <- indices[matched]
    original_names <- original_names[matched]
  }

  if (length(indices) == 0) {
    warning("No attachments matched attachment.pattern.")
    return(invisible(character(0)))
  }

  if (!is.null(file_name)) {
    if (length(file_name) != length(indices)) {
      stop(
        "file_name has length ", length(file_name),
        " but ", length(indices), " attachment(s) matched. ",
        "Use attachment.pattern to narrow down to the intended attachment(s), ",
        "or supply one file_name per matched attachment."
      )
    }
    save_names <- file_name
  } else {
    save_names <- original_names
  }

  save_paths <- file.path(path, save_names)

  purrr::walk2(indices, save_paths, function(i, save_path) {
    if (file.exists(save_path) && !overwrite) {
      warning("File already exists and overwrite = FALSE, skipping: ", save_path)
      return(invisible(NULL))
    }
    attachments$Item(i)$SaveAsFile(save_path)
  })

  return(invisible(save_paths))
}

# Sender ------

#' Get the Sender of an Email
#'
#' This function returns the sender's name of a given email object.
#'
#' @param email A COMIDispatch object representing an email.
#'
#' @return A character string representing the sender's name, or NA if not found.
#' @export
#'
#' @examples
#' \dontrun{
#' email_obj <- get_email_object()
#' sender(email_obj)
#' }
sender <- function(email) {
  if (!inherits(email, "COMIDispatch")) {
    stop("Invalid email object.")
  }

  result <- tryCatch(
    email$SenderName(),
    error = function(e) {
      warning("Sender name not found.")
      return(NA_character_)
    }
  )

  return(result)
}


#' Get the Sender's Email Address of an Email
#'
#' This function returns the sender's email address of a given email object.
#'
#' @param email A COMIDispatch object representing an email.
#'
#' @return A character string representing the sender's email address, or NA if not found.
#' @export
#'
#' @examples
#' \dontrun{
#' email_obj <- get_email_object()
#' sender_address(email_obj)
#' }
sender_address <- function(email) {
  if (!inherits(email, "COMIDispatch")) {
    stop("Invalid email object.")
  }

  result <- tryCatch(
    email$SenderEmailAddress(),
    error = function(e) {
      warning("Sender email address not found.")
      return(NA_character_)
    }
  )

  return(result)
}


# Subject -------

#' Get the Subject of an Email
#'
#' This function returns the subject of a given email object.
#'
#' @param email A COMIDispatch object representing an email.
#'
#' @return A character string representing the subject, or NA if not found.
#' @export
#'
#' @examples
#' \dontrun{
#' email_obj <- get_email_object()
#' subject(email_obj)
#' }
subject <- function(email) {
  if (!inherits(email, "COMIDispatch")) {
    stop("Invalid email object.")
  }

  result <- tryCatch(
    email$Subject(),
    error = function(e) {
      warning("Subject not found.")
      return(NA_character_)
    }
  )

  return(result)
}

# Body ---------

#' Get the Body of an Email
#'
#' This function returns the body of a given email object.
#'
#' @param email A COMIDispatch object representing an email.
#'
#' @return A character string representing the body, or NA if not found.
#' @export
#'
#' @examples
#' \dontrun{
#' email_obj <- get_email_object()
#' body(email_obj)
#' }
body <- function(email) {
  if (!inherits(email, "COMIDispatch")) {
    stop("Invalid email object.")
  }

  result <- tryCatch(
    email$Body(),
    error = function(e) {
      warning("Body not found.")
      return(NA_character_)
    }
  )

  return(result)
}


# Date Received ---------

#' Get the Received Date of an Email
#'
#' This function returns the received date of a given email object.
#'
#' @param email A COMIDispatch object representing an email.
#'
#' @return A POSIXct object representing the received date, or NA if not found.
#' @export
#'
#' @examples
#' \dontrun{
#' email_obj <- get_email_object()
#' date_received(email_obj)
#' }
date_received <- function(email) {
  if (!inherits(email, "COMIDispatch")) {
    stop("Invalid email object.")
  }

  result <- tryCatch(
    convert_comdate(email$ReceivedTime()),
    error = function(e) {
      warning("Received date not found.")
      return(NA)
    }
  )

  return(result)
}

# Date Sent ---------

#' Get the Sent Date of an Email
#'
#' This function returns the sent date of a given email object.
#'
#' @param email A COMIDispatch object representing an email.
#'
#' @return A POSIXct object representing the sent date, or NA if not found.
#' @export
#'
#' @examples
#' \dontrun{
#' email_obj <- get_email_object()
#' date_sent(email_obj)
#' }
date_sent <- function(email) {
  if (!inherits(email, "COMIDispatch")) {
    stop("Invalid email object.")
  }

  result <- tryCatch(
    convert_comdate(email$SentOn()),
    error = function(e) {
      warning("Sent date not found.")
      return(NA)
    }
  )

  return(result)
}

# read_emails ---------
#' Read Emails and Extract Information
#'
#' This function reads emails from an Outlook folder, with option to filter by given number of most recent emails, emails only after a certain date, emails with a certain subject, or all of the above. It extracts specified information.
#'
#' @param most.recent.n Number of most recent emails to read. Default is NULL, which reads all emails.
#' @param after.this.date A POSIXct date/time. Only emails newer than this date/time will be returned. Default is NULL which retrieves all emails.
#' @param with.subject.like A character string with the subject line to be searched for. Default is NULL which returns emails with any subject.
#' @param subject.exact.match Indicates whether the subject line search will look for exact matches only, or whether it will check if the subject string is in the email's subject. Default is "FALSE" which indicates do not look for exact match.
#' @param in.folder.name Name of folder within account in which to search. Default is "Inbox".
#' @param in.account.name Name of account in which to search. Default is NULL, which searches in the user's default Outlook account.
#' @param date_sent Logical, whether to include the sent date. Default is TRUE.
#' @param date_received Logical, whether to include the received date. Default is TRUE.
#' @param body Logical, whether to include the email body. Default is TRUE.
#' @param subject Logical, whether to include the email subject. Default is TRUE.
#' @param sender Logical, whether to include the sender's name. Default is TRUE.
#' @param attachment_count Logical, whether to include the attachment count. Default is TRUE.
#' @return A tibble containing the extracted email information.
#' @importFrom dplyr filter select
#' @importFrom purrr map
#' @importFrom rlang set_names
#' @importFrom tidyr unnest
#' @importFrom tibble as_tibble tribble
#' @importFrom tidyselect everything

#' @export
read_emails <- function (
  most.recent.n = NULL,
  after.this.date = NULL,
  with.subject.like = NULL,
  subject.exact.match = FALSE,
  in.folder.name = "Inbox",
  in.account.name = NULL,
  date_sent = TRUE,
  date_received = TRUE,
  body = TRUE,
  subject = TRUE,
  sender = TRUE,
  attachment_count = TRUE
){

  # Validate arguments
  if (!is.null(most.recent.n) & (!is.numeric(most.recent.n) || most.recent.n <= 0)) {
    stop("Invalid value for most.recent.n. Please input a positive integer.")
  }
  if (!is.null(after.this.date) & !('POSIXct' %in% class(after.this.date))) {
    stop("Invalid date for after.this.date. Please use POSIXct format.")
  }
  if (typeof(subject.exact.match) != 'logical') {
    stop("Invalid input for subject_exact_match. Please use logical format.")
  }


  # Define which functions to run based on arguments
  func_list <- tibble::tribble(
    ~name,              ~argument,          ~func,
    "date_sent",        date_sent,          email::date_sent,
    "date_received",    date_received,      email::date_received,
    "body",             body,               email::body,
    "subject",          subject,            email::subject,
    "sender",           sender,             email::sender,
    "attachment_count", attachment_count,   email::attachment_count
  ) %>%
    dplyr::filter(.data$argument) %>%
    dplyr::select(-.data$argument)

  # Get emails
  emails <- tryCatch(
    email::get_emails(most.recent.n = most.recent.n,
                      after.this.date = after.this.date,
                      subject = with.subject.like,
                      subject.exact.match = subject.exact.match,
                      folder.name = in.folder.name,
                      account.name = in.account.name),
    error = function(e) {
      stop("Failed to get emails.")
    }
  )

  # Helper function to run each function on the list of emails
  run_functions <- function (func, emails) {
    purrr::map(emails, func)
  }

  # Run the functions and assemble the results
  purrr::map(func_list$func, run_functions, emails) %>%
    rlang::set_names(func_list$name) %>%
    tibble::as_tibble() %>%
    tidyr::unnest(cols = tidyselect::everything())
}
