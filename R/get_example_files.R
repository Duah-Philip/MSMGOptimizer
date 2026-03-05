#' Get path to MSMG example and tutorial files
#'
#' @description Returns the file path to bundled tutorial
#'     and example setup files included with MSMGOptimizer.
#'     Available files:
#'     \itemize{
#'       \item \code{"Setup File.xlsx"}
#'       \item \code{"Multiproduct_Setup File.xlsx"}
#'       \item \code{"Sheet_by_Sheet_Setup File.xlsx"}
#'       \item \code{"Tutorial.docx"}
#'     }
#' @param file Name of the file. If \code{NULL} returns
#'     the directory path.
#' @return A character string with the full file path.
#' @examples
#' \dontrun{
#'   # Open the tutorial
#'   shell.exec(msmg_example("Tutorial.docx"))
#'
#'   # Load a setup file
#'   readxl::read_excel(msmg_example("Setup File.xlsx"))
#'
#'   # See all available files
#'   list.files(msmg_example())
#' }
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
