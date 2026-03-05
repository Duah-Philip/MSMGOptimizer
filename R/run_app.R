
#' Launch the MSMG SimaPro CSV Optimizer
#'
#' @description Starts the MSMG SimaPro CSV Optimizer Shiny application.
#' @param ... Arguments passed to \code{shiny::runApp()}.
#' @export
#' @importFrom shiny runApp
ShinyMSMGOptimizer <- function(...) {
  app_dir <- system.file("shiny", package = "MSMGOptimizer")
  if (app_dir == "") {
    stop("Could not find Shiny app directory. Try re-installing the package.",
         call. = FALSE)
  }
  shiny::runApp(app_dir, ...)
}

#' @importFrom dplyr mutate filter select
#' @importFrom readxl read_excel
#' @importFrom DT datatable
#' @importFrom shinydashboard dashboardPage
#' @importFrom waiter Waiter
#' @importFrom htmltools tags
#' @importFrom zip zip
NULL
