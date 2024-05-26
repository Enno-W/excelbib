xlsx_to_bib <- function(excel_file, bib_file = "bibliography.bib", sheet = 2, column = 1, first_row = 2) {

  if (!file.exists(excel_file)) {
    stop("The specified Excel file does not exist.")
  }
  excel_data <- xlsx::read.xlsx(excel_file, sheetIndex = sheet, colIndex = column, startRow = first_row, as.data.frame = TRUE, header = F)

  citation_keys <- as.character(stats::na.omit(excel_data[, 1]))

  writeLines(citation_keys, bib_file)

  message("Bibliography has been successfully written to ", bib_file)
}
