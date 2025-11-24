#' Send an Email via Outlook with Optional HTML and Attachments
#'
#' This function sends an email using Microsoft Outlook. At least one recipient
#' must be supplied through `to`, `cc`, or `bcc`. The message body must be
#' provided either as plain text (`body`) or HTML (`html_body`), but not both.
#'
#' @param to Optional. A character vector of primary recipient email addresses.
#'   At least one of `to`, `cc`, or `bcc` must be specified.
#' @param from Optional. A character string specifying the sender to send on
#'   behalf of (Outlook "SentOnBehalfOfName").
#' @param cc Optional. A character vector of CC recipients.
#' @param bcc Optional. A character vector of BCC recipients.
#' @param subject A character string specifying the email subject.
#' @param body Plain-text body of the email. Must not be supplied together with
#'   `html_body`.
#' @param html_body HTML body of the email. Must not be supplied together with
#'   `body`.
#' @param attachment_paths Optional character vector of file paths to attach.
#'
#' @return A message indicating that the email was sent successfully.
#'
#' @examples
#' \dontrun{
#' # Plain text email
#' send(
#'   to = "recipient@example.com",
#'   subject = "Test",
#'   body = "This is a plain text message."
#' )
#'
#' # HTML email with CC only
#' send(
#'   cc = "someone@example.com",
#'   subject = "Test HTML",
#'   html_body = "<b>Hello!</b>"
#' )
#'
#' # BCC-only email (allowed)
#' send(
#'   bcc = "hidden@example.com",
#'   subject = "BCC announcement",
#'   body = "This message has no To/CC recipients."
#' )
#' }
#'
#' @importFrom purrr walk
#' @export
send <- function(
    to = NULL,
    from = NULL,
    cc = NULL,
    bcc = NULL,
    subject,
    body,
    html_body = NULL,
    attachment_paths = NULL
) {
  # Validation: enforce exactly one body type
  if (!missing(body) && !missing(html_body)) {
    stop("Supply only one of `body` or `html_body`, not both.")
  }

  if (missing(body) && missing(html_body)) {
    stop("You must supply either `body` (plain text) or `html_body` (HTML).")
  }

  # Must have at least one recipient
  if (is.null(to) && is.null(cc) && is.null(bcc)) {
    stop("You must supply at least one of `to`, `cc`, or `bcc`.")
  }


  # --- Helpers ---
  to_chr_or_null <- function(x) {
    if (is.null(x)) return(NULL)
    out <- as.character(x)
    if (length(out) == 0) return(NULL)
    out
  }

  collapse_recipients <- function(x) {
    if (is.null(x)) return(NULL)
    if (length(x) == 0) return(NULL)
    paste(x, collapse = "; ")
  }

  set_com_field <- function(email, field, value) {
    if (is.null(value)) return(invisible(email))
    if (length(value) == 0) return(invisible(email))
    email[[field]] <- value
    invisible(email)
  }

  # --- Type corrections ---
  to        <- to_chr_or_null(to)
  cc        <- to_chr_or_null(cc)
  bcc       <- to_chr_or_null(bcc)
  from      <- to_chr_or_null(from)
  subject   <- to_chr_or_null(subject)
  body      <- to_chr_or_null(body)
  html_body <- to_chr_or_null(html_body)

  # --- Initialize COM ---
  Outlook <- RDCOMClient::COMCreate("Outlook.Application")
  email   <- Outlook$CreateItem(0)

  # --- Collapse recipient lists ---
  to  <- collapse_recipients(to)
  cc  <- collapse_recipients(cc)
  bcc <- collapse_recipients(bcc)

  # --- COM field assignments ---
  set_com_field(email, "To", to)
  set_com_field(email, "SentOnBehalfOfName", from)
  set_com_field(email, "CC", cc)
  set_com_field(email, "BCC", bcc)
  set_com_field(email, "Subject", subject)

  # --- Body (HTML or plain text) ---
  if (!is.null(html_body)) {
    email[["BodyFormat"]] <- 2
    email[["HTMLBody"]]   <- html_body
  } else {
    set_com_field(email, "Body", body)
  }

  # --- Attachments ---
  if (!is.null(attachment_paths)) {
    purrr::walk(attachment_paths, ~ email[["Attachments"]]$Add(.))
  }

  # --- Send ---
  email$Send()
  message("Email sent successfully.")
}
