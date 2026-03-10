#' Launch the MSMG 'SimaPro' CSV Optimizer 'Shiny' Application
#'
#' @description
#' Launches a 'Shiny' (web application framework for R) application that
#' converts 'Excel'-based Life Cycle Inventory (LCI) data into
#' 'SimaPro' CSV (Comma-Separated Values) format compatible with the
#' 'SimaPro' Life Cycle Assessment (LCA) software. The app provides
#' an interactive interface for uploading 'Excel' files, previewing
#' worksheet contents, converting all sheets to individual CSV files,
#' and downloading the results as a ZIP (compressed archive) file.
#' Developed by the Mine Sustainability Modeling Group (MSMG) at
#' Missouri University of Science and Technology.
#'
#' @return No return value. This function is called for its side effect
#'   of launching an interactive 'Shiny' application in the user's
#'   default web browser or 'RStudio' viewer pane.
#'
#' @examples
#' if (interactive()) {
#'   ShinyMSMGOptimizer()
#' }
#'
#' @seealso \code{\link{msmg_example}} for accessing bundled example files.
#'
#' @export
ShinyMSMGOptimizer <- function() {
  appDir <- system.file("shiny", package = "MSMGOptimizer")
  if (appDir == "") {
    stop("Could not find 'Shiny' app directory. Try re-installing 'MSMGOptimizer'.",
         call. = FALSE)
  }
  shiny::runApp(appDir, display.mode = "normal")
}

#' @importFrom shiny runApp
#' @importFrom dplyr mutate filter select
#' @importFrom readxl read_excel excel_sheets
#' @importFrom DT datatable renderDataTable dataTableOutput
#' @importFrom shinydashboard dashboardPage dashboardHeader dashboardSidebar
#'   dashboardBody sidebarMenu menuItem tabItems tabItem box
#' @importFrom waiter Waiter useWaiter spin_flower
#' @importFrom htmltools tags tagList
#' @importFrom zip zip
NULL
