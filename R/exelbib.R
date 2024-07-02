#' Title
#'
#' @param excel_file The name or directory of the Excel file you use for your References. In the file, the bibtex-entries should be arranged in a column, one entry per cell.
#' @param bib_file The name of your .bib-file. Defaults to "bibliography.bib". The file will be added to your working directory.
#' @param sheet The sheet number from which you want to import your references. Defaults to 2.
#' @param column The column from which you want to import your references. The default is 1.
#' @param first_row The first row that contains a bibtex entry that you want to cite.
#'
#' @return An error message if the file does not exist.
#' @export
#'
#' @examples
#' download.file(url="https://raw.githubusercontent.com/Enno-W/excelbib/main/Excel_References.xlsx", destfile = "Excel_References.xlsx", mode ="wb")
#' xlsx_to_bib ("Excel_References.xlsx", bib_file="my_bibliography.bib", sheet = 2, column = 1, first_row = 2)
xlsx_to_bib <- function(excel_file, bib_file = "bibliography.bib", sheet = 2, column = 1, first_row = 2) {
  requireNamespace("xlsx")
  if (!file.exists(excel_file)) {
    stop("The specified Excel file does not exist.")
  }
  excel_data <- xlsx::read.xlsx(excel_file, sheetIndex = sheet, colIndex = column, startRow = first_row, as.data.frame = TRUE, header = F)

  citation_keys <- as.character(stats::na.omit(excel_data[, 1]))

  writeLines(citation_keys, bib_file)

  message("Bibliography has been successfully written to ", bib_file)
}

#' Title
#'
#' @param bib_file_path
#' @param excel_file_path
#' @param sheet_name
#'
#' @return References in your Excel file!
#' @export
#'
#' @examples
#' download.file(url="https://raw.githubusercontent.com/Enno-W/excelbib/main/Excel_References.xlsx", destfile = "Excel_References.xlsx", mode ="wb")
#' download.file(url="https://raw.githubusercontent.com/Enno-W/excelbib/main/export.bib", destfile = "Excel_References.xlsx", mode ="wb")
process_bib_to_excel <- function(bib_file_path = "export.bib", excel_file_path = "Excel_References.xlsx", sheet_name = "Import") {
  requireNamespace("openxlsx")
  bib_file <- readLines(bib_file_path, warn = FALSE)
  text <- paste(bib_file, collapse = " ")
  entries <- unlist(strsplit(text, split = "\\@"))
  entries <- entries[entries != ""]
  entries <- paste("@", entries, sep = "")
  data <- data.frame(entries)
  wb <- openxlsx::loadWorkbook(excel_file_path)
  writeData(wb, sheet_name, data, colNames = FALSE)
  openxlsx::saveWorkbook(wb, excel_file_path, overwrite = TRUE)
  message("Bibliography has been successfully written to ", excel_file_path)
}
