#' Execute code and ping me if something goes wrong (or right)
#'
#' Wraps a code block and sends email notifications based on execution outcome.
#' Useful for automated scripts where you want to be notified of failures
#' (or successes).
#'
#' @param code A code block (in curly braces) or expression to execute
#' @param to Email address(es) to notify. Can be a single address or character vector.
#' @param subject_on_error Subject line if code errors. Default: "Code Execution Error"
#' @param subject_on_success Subject line if code succeeds. Default: "Code Execution Successful"
#' @param body_on_error Body text if code errors. The R error message will be
#'   appended automatically. Default: "An error occurred during code execution."
#' @param body_on_success Body text if code succeeds. Default: "Code executed successfully."
#' @param stop_on_error Logical. If TRUE (default), stops execution after sending
#'   the error notification. If FALSE, sends notification but returns NULL and
#'   continues.
#' @param notify_on_success Logical. If TRUE, sends an email when code runs
#'   successfully. Default is FALSE.
#'
#' @return If successful, returns the result of the code block invisibly.
#'   If error and stop_on_error = FALSE, returns NULL invisibly.
#'
#' @examples
#' \dontrun{
#' # Wrap a code block
#' ping_me({
#'   data <- read.csv("important_data.csv")
#'   process(data)
#'   save_results()
#' }, to = "admin@example.com")
#'
#' # Wrap a single function call
#' ping_me(
#'   source("daily_report.R"),
#'   to = "admin@example.com",
#'   notify_on_success = TRUE
#' )
#'
#' # Continue on error instead of stopping
#' result <- ping_me({
#'   risky_operation()
#' }, to = "admin@example.com", stop_on_error = FALSE)
#'
#' # Custom messages for context
#' ping_me({
#'   source("daily_report.R")
#' },
#'   to = "team@example.com",
#'   subject_on_error = "Daily Report Failed",
#'   body_on_error = "The daily report script failed to complete.",
#'   notify_on_success = TRUE,
#'   subject_on_success = "Daily Report Complete",
#'   body_on_success = "The daily report ran successfully."
#' )
#' }
#'
#' @export
ping_me <- function(
    code,
    to,
    subject_on_error = "Code Execution Error",
    subject_on_success = "Code Execution Successful",
    body_on_error = "An error occurred during code execution.",
    body_on_success = "Code executed successfully.",
    stop_on_error = TRUE,
    notify_on_success = FALSE
) {

  code_expr <- substitute(code)

  tryCatch({
    result <- eval(code_expr, envir = parent.frame())

    if (notify_on_success) {
      for (recipient in to) {
        send(
          to = recipient,
          subject = subject_on_success,
          body = body_on_success
        )
      }
    }

    invisible(result)

  }, error = function(e) {
    error_body <- paste0(
      body_on_error,
      "\n\nR error message:\n",
      conditionMessage(e)
    )

    for (recipient in to) {
      send(
        to = recipient,
        subject = subject_on_error,
        body = error_body
      )
    }

    if (stop_on_error) {
      stop("Execution stopped after error notification was sent.", call. = FALSE)
    } else {
      invisible(NULL)
    }
  })
}
