## ----setup, include = FALSE---------------------------------------------------
library(openxlsx2)
options(rmarkdown.html_vignette.check_title = FALSE)
knitr::opts_chunk$set(
  collapse = TRUE,
  comment = "#>"
)

## -----------------------------------------------------------------------------
xlsxFile <- system.file("extdata", "openxlsx2_example.xlsx", package = "openxlsx2")

## -----------------------------------------------------------------------------
# import workbook
wb_to_df(xlsxFile)

## -----------------------------------------------------------------------------
# do not convert first row to colNames
wb_to_df(xlsxFile, colNames = FALSE)

## -----------------------------------------------------------------------------
# do not try to identify dates in the data
wb_to_df(xlsxFile, detectDates = FALSE)

## -----------------------------------------------------------------------------
# return the underlying Excel formula instead of their values
wb_to_df(xlsxFile, showFormula = TRUE)

## -----------------------------------------------------------------------------
# read dimension withot colNames
wb_to_df(xlsxFile, dims = "A2:C5", colNames = FALSE)

## -----------------------------------------------------------------------------
# read selected cols
wb_to_df(xlsxFile, cols = c("A:B", "G"))

## -----------------------------------------------------------------------------
# read selected rows
wb_to_df(xlsxFile, rows = c(2, 4, 6))

## -----------------------------------------------------------------------------
# convert characters to numerics and date (logical too?)
wb_to_df(xlsxFile, convert = FALSE)

## -----------------------------------------------------------------------------
# erase empty Rows from dataset
wb_to_df(xlsxFile, sheet = 1, skipEmptyRows = TRUE) |> tail()

## -----------------------------------------------------------------------------
# erase empty Cols from dataset
wb_to_df(xlsxFile, skipEmptyCols = TRUE)

## -----------------------------------------------------------------------------
# convert first row to rownames
wb_to_df(xlsxFile, sheet = 2, dims = "C6:G9", rowNames = TRUE)

## -----------------------------------------------------------------------------
# define type of the data.frame
wb_to_df(xlsxFile, cols = c(2, 5), types = c("Var1" = 0, "Var3" = 1))

## -----------------------------------------------------------------------------
# start in row 5
wb_to_df(xlsxFile, startRow = 5, colNames = FALSE)

## -----------------------------------------------------------------------------
# na string
wb_to_df(xlsxFile, na.strings = "")

## -----------------------------------------------------------------------------
# the file we are going to load
xlsxFile <- system.file("extdata", "openxlsx2_example.xlsx", package = "openxlsx2")
# loading the file into the workbook
wb <- wb_load(file = xlsxFile)

## ---- eval = FALSE------------------------------------------------------------
#  write_xlsx(mtcars, "mtcars.xlsx")

## ---- eval = FALSE------------------------------------------------------------
#  # replace the existing file
#  wb$save("mtcars.xlsx")
#  
#  # do not overwrite the exisisting file
#  try(wb$save("mtcars.xlsx", overwrite = FALSE))

