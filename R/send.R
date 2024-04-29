#' Send an Email via Outlook with Optional Multiple Attachments
#'
#' This function sends an email using Microsoft Outlook. It allows you to specify recipients in the "To", "From", "CC", and "BCC" fields, and optionally attach multiple files.
#'
#' @param to A character vector of recipient email addresses for the "To" field.
#' @param from A character string specifying the sender's email address for the "From" field.
#' @param cc A character vector of recipient email addresses for the "CC" field. Default is NULL.
#' @param bcc A character vector of recipient email addresses for the "BCC" field. Default is NULL.
#' @param subject A character string specifying the email subject.
#' @param body A character string specifying the email body.
#' @param attachment_paths A character vector specifying the full file paths of the attachments. Default is NULL.
#'
#' @return A message indicating the email was sent successfully.
#'
#' @examples
#' \dontrun{
#' send(
#'   to = c("recipient1@example.com", "recipient2@example.com"),
#'   from = "sender@example.com",
#'   cc = c("cc1@example.com", "cc2@example.com"),
#'   bcc = c("bcc1@example.com", "bcc2@example.com"),
#'   subject = "Test Subject",
#'   body = "This is a test email.",
#'   attachment_paths = c("C:/path/to/your/file1.txt", "C:/path/to/your/file2.txt")
#' )
#' }
#'
#' @importFrom purrr walk
#'
#' @export
send <- function(to, from = NULL, cc = NULL, bcc = NULL, subject, body, attachment_paths = NULL) {
  # Initialize Outlook and get Namespace
  Outlook <- RDCOMClient::COMCreate("Outlook.Application")

  # Create a new MailItem object
  email <- Outlook$CreateItem(0)

  # Concatenate multiple email addresses with semicolons
  to <- paste(to, collapse = "; ")
  if (!is.null(cc)) cc <- paste(cc, collapse = "; ")
  if (!is.null(bcc)) bcc <- paste(bcc, collapse = "; ")

  # Set the properties of the email
  email[["To"]] <- to
  if (!is.null(from)) email[["SentOnBehalfOfName"]] = from
  if (!is.null(cc)) email[["CC"]] <- cc
  if (!is.null(bcc)) email[["BCC"]] <- bcc
  email[["Subject"]] <- subject
  email[["Body"]] <- body

  # Add attachments if specified
  if (!is.null(attachment_paths)) {
    purrr::walk(attachment_paths, ~email[["Attachments"]]$Add(.))
  }

  # Send the email
  email$Send()

  message("Email sent successfully.")
}
