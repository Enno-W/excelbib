This package lets you use Excel almost the same way as you would use a reference software, if you use Rmarkdown or Quarto. 
1. Install using  `devtools::install_github("Enno-W/excelbib")` and load `library(excelbib)`.
2. Create an excel file where you want to organise your references `add_xlsx()`. An .xlsx file will download from Github to your working directory. An existing file will not be overwritten. 
3. Add your references as BibTex-entries to the first column of sheet 2 or import them from a .bib -file using `bib_to_xlsx("your .bib file here")`. After importing, the references will appear in the first sheet. Copy and paste the ones that you want to cite to the first column of sheet 2. The citekey, authors, doi and other information will be extracted by Excel. For that, the BibTex-entry must be formatted with one space before and after the equal sign (=) and using curly brackets. You can use ChatGPT to create such BibTex entries from a reference, or use a doi to .bib-converter like https://www.doi2bib.org/. 
5. To write your references to the bibliography file, use `xlsx_to_bib()`. The default name for the .bib-file is "bibliography.bib"
6. Now you can use your references in Rmardown or Quarto, simply include `bibliography: bibliography.bib` in the YAML header.
7. See https://quarto.org/docs/authoring/citations.html for citation basics in Quarto, and https://en.wikibooks.org/wiki/LaTeX/Bibliography_Management#BibTeX for more information on BibTex. 
