#####################################################################

library(openxlsx)

#####################################################################
# read xlsx files to dfs list
#####################################################################

#' Read-in Excel file (workbook) as a list of data frames.
#'
#' @param xlsxPath A path to the Excel file, a character.
#' @param rowNames Whether to read-in row names, a boolean.
#' @param colNames Whether to read-in column names, a boolean.
#' @param ... ... passed to read.xlsx function in the openxlsx package.
#' @return A list of data frames, each representing a sheet in the Excel file (sheet names are list element names).
#' @import openxlsx
#' @examples
#' # create example file
#' df1 <- data.frame(A=c(1, 2), B=c(3, 4))
#' df2 <- data.frame(C=c(5, 6), D=c(7, 8))
#' xlsx_fpath <- file.path(tempdir(), "testout.xlsx")
#' dfs2xlsx(withNames("sheet1", df1, "sheet2", df2), xlsx_fpath)
#' # read created file
#' dfs <- xlsx2dfs(xlsx_fpath)
#' file.remove(xlsx_fpath)
#' @export
xlsx2dfs <- function(xlsxPath, rowNames = TRUE, colNames = TRUE, ...) {
  wb <- openxlsx::loadWorkbook(xlsxPath)
  sheetNames <- names(wb)
  res <- lapply(sheetNames, function(sheetName) {
    openxlsx::read.xlsx(wb, sheet = sheetName, rowNames = rowNames, colNames = colNames, ...)
  })
  names(res) <- sheetNames
  res
}


#####################################################################
# printing dfs to xlsx files
#####################################################################

#' Helper function for more convenient input (sheet name, data frame, sheet name, data frame, ...).
#'
#' @param ... alterning arguments: sheet name 1, data frame 1, sheet name 2, data frame 2, ...
#' @return A list of the data frames with the names given each before.
#' @examples
#' df1 <- data.frame(A=c(1, 2), B=c(3, 4))
#' df2 <- data.frame(C=c(5, 6), D=c(7, 8))
#' xlsx_fpath <- file.path(tempdir(), "testout.xlsx")
#' dfs2xlsx(withNames("sheet1", df1, "sheet2", df2), xlsx_fpath)
#' file.remove(xlsx_fpath)
#' @export
withNames <- function(...) {
  p.l <- list(...)
  len <- length(p.l)
  if (len %% 2 == 1) {
    stop("withNames call with odd numbers of arguments")
    print()
  }
  seconds <- p.l[seq(2, len, 2)]
  firsts <- p.l[seq(1, len, 2)]
  names(seconds) <- unlist(firsts)
  seconds
}


#' Write a list of data frames into an excel file with each data frame in a new sheet and the list element name as its sheet name.
#'
#' @param dfs A list of data frames (names in the list are the names of the sheets).
#' @param fpath A character string representing path and filename of the output.
#' @param rowNames A boolean indicating whether the first column of a table in every sheet contains row names of the table.
#' @param colNames A boolean indicating whether the first line of a table in every sheet contains a header.
#' @return Nothing. Writes out data frames into specified Excel file.
#' @import openxlsx
#' @examples
#' df1 <- data.frame(A=c(1, 2), B=c(3, 4))
#' df2 <- data.frame(C=c(5, 6), D=c(7, 8))
#' xlsx_fpath <- file.path(tempdir(), "testout.xlsx")
#' dfs2xlsx(withNames("sheet1", df1, "sheet2", df2), xlsx_fpath)
#' file.remove(xlsx_fpath)
#' @export
dfs2xlsx <- function(dfs, fpath, rowNames=TRUE, colNames=TRUE) {
  wb <- createWorkbook()
  Map(function(data, name) {
    openxlsx::addWorksheet(wb, name)
    openxlsx::writeData(wb, name, data, rowNames = rowNames, colNames = colNames)
  }, dfs, names(dfs))
  openxlsx::saveWorkbook(wb, file = fpath, overwrite = TRUE)
}
