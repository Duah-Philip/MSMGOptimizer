#' Get Path to MSMG (Mine Sustainability Modeling Group) Example Files
#'
#' @description
#' Returns the file path to bundled tutorial and example setup files
#' included with the MSMGOptimizer package. These files are stored in
#' the package's \code{inst/extdata} directory and can be used as
#' templates for preparing Life Cycle Inventory (LCI) data for
#' conversion with the \code{ShinyMSMGOptimizer()} application.
#'
#' Available files:
#' \itemize{
#'   \item \code{"Setup_File.xlsx"} — Standard single-product LCI
#'         (Life Cycle Inventory) template
#'   \item \code{"Multiproduct_Setup_File.xlsx"} — Template for
#'         processes with multiple co-products
#'   \item \code{"Sheet_by_Sheet_Setup_File.xlsx"} — Template where
#'         each worksheet represents one LCI process
#'   \item \code{"Tutorial.docx"} — Step-by-step tutorial document
#' }
#'
#' @param file Character string. Name of a specific file to retrieve.
#'   If \code{NULL} (default), returns the path to the \code{extdata}
#'   directory containing all bundled files.
#'
#' @return A character string giving the absolute file path to the
#'   requested file or directory. If \code{file = NULL}, returns the
#'   path to the \code{extdata} directory. If a specific file is
#'   requested, returns the full path to that file. Raises an error
#'   if the requested file does not exist in the package.
#'
#' @examples
#' # List all bundled example files
#' list.files(msmg_example())
#'
#' # Get the path to a specific setup file
#' msmg_example("Setup_File.xlsx")
#'
#' # Check that the tutorial file exists
#' file.exists(msmg_example("Tutorial.docx"))
#'
#' @export
msmg_example <- function(file = NULL) {
  if (is.null(file)) {
    dir <- system.file("extdata", package = "MSMGOptimizer")
    return(dir)
  }
  path <- system.file("extdata", file, package = "MSMGOptimizer")
  if (path == "") {
    stop(paste("File not found:", file,
               "\nAvailable files: Setup_File.xlsx,",
               "Multiproduct_Setup_File.xlsx,",
               "Sheet_by_Sheet_Setup_File.xlsx,",
               "Tutorial.docx"),
         call. = FALSE)
  }
  return(path)
}
