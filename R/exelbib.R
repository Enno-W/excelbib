#' Update your References from an Excel-file
#'
#' @param excel_file The name or directory of the Excel file you use for your References. In the file, the bibtex-entries should be arranged in a column, one entry per cell.
#' @param bib_file The name of your .bib-file. Defaults to "bibliography.bib". The file will be added to your working directory.
#' @param sheet The sheet number from which you want to import your references. Defaults to 2.
#' @param column The column from which you want to import your references. The default is 1.
#' @param first_row The first row that contains a bibtex entry that you want to cite.
#'
#' @return An error message if the file does not exist.
#' @export
#' @import xlsx openxlsx utils
#'
#' @examples
#' download.file(url="https://raw.githubusercontent.com/Enno-W/excelbib/main/Excel_References_example.xlsx", destfile = "Excel_References_example.xlsx", mode ="wb")
#' require("xlsx")
#' xlsx_to_bib ("Excel_References_example.xlsx", bib_file="example_bibliography.bib", sheet = 2, column = 1, first_row = 2)
xlsx_to_bib <- function(excel_file = "References.xlsx", bib_file = "bibliography.bib", sheet = 2, column = 1, first_row = 2) {
  if (!file.exists(excel_file)) {
    stop("The specified Excel file does not exist.")
  }
  wb <- openxlsx::loadWorkbook(excel_file)
  data <- openxlsx::readWorkbook(wb, sheet = sheet, colNames = FALSE, startRow = first_row)

  citation_keys <- as.character(stats::na.omit(data[[column]]))

  writeLines(citation_keys, bib_file)

  message("Bibliography has been successfully written to ", bib_file)
}

#' Import References to your Excel-File
#'
#' @param excel_file_path The name or directory of the Excel file you use for your References. In the file, the bibtex-entries should be arranged in a column, one entry per cell.
#' @param bib_file_path The name of your .bib-file. Defaults to "export.bib".
#' @param sheet_name Sheet name where you want the references imported to, default is "Import".
#'
#' @return References in your Excel file
#' @export
#' @import xlsx openxlsx utils
#'
#' @examples
#' download.file(url="https://raw.githubusercontent.com/Enno-W/excelbib/main/Excel_References_example.xlsx", destfile = "Excel_References_example.xlsx", mode ="wb")
#' download.file(url="https://raw.githubusercontent.com/Enno-W/excelbib/main/export_example.bib", destfile = "export_example.bib", mode ="wb")
#' require("openxlsx")
#' bib_to_xlsx(bib_file_path = "export_example.bib", excel_file_path = "Excel_References_example.xlsx")
bib_to_xlsx <- function(bib_file_path = "export.bib", excel_file_path = "References.xlsx", sheet_name = "Import") {
  if (!file.exists(bib_file_path)) {
    stop("The specified .bib file does not exist: ", bib_file_path)
  }
  bib_file <- readLines(bib_file_path, warn = FALSE)
  text <- paste(bib_file, collapse = " ")
  entries <- unlist(strsplit(text, split = "\\@"))
  entries <- entries[entries != ""]
  entries <- paste("@", entries, sep = "")
  data <- data.frame(entries)
  wb <- openxlsx::loadWorkbook(excel_file_path)
  openxlsx::writeData(wb, sheet_name, data, colNames = FALSE)
  openxlsx::saveWorkbook(wb, excel_file_path, overwrite = TRUE)
  message("Bibliography has been successfully imported to ", excel_file_path)
}

#' Add an Empty Excel-File and Start using Excel as Reference Software
#'
#' @param xlsx_name The name you want to use for your Excel-file where you organize your references.
#'
#' @return
#' @export
#' @import utils
#'
#' @examples
#' add_xlsx(xlsx_name= "Excel_References_example")
add_xlsx <- function(xlsx_name = "References.xlsx"){

  if (!file.exists(xlsx_name)) {
    download.file(url = "https://raw.githubusercontent.com/Enno-W/excelbib/main/Excel_References_example.xlsx",
                  destfile = xlsx_name,
                  mode = "wb")
    message("An empty Excel-file for your references was added")
  } else {
    message ("File already exists and will not be overwritten.")
  }

}
