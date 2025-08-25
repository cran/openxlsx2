## ----setup, include = FALSE---------------------------------------------------
library(openxlsx2)
options(rmarkdown.html_vignette.check_title = FALSE)
knitr::opts_chunk$set(
  collapse = TRUE,
  comment = "#>"
)

## -----------------------------------------------------------------------------
file <- system.file("extdata", "openxlsx2_example.xlsx", package = "openxlsx2")

## -----------------------------------------------------------------------------
# import workbook
library(openxlsx2)
wb_to_df(file)

## -----------------------------------------------------------------------------
# do not convert first row to column names
wb_to_df(file, col_names = FALSE)

## -----------------------------------------------------------------------------
# do not try to identify dates in the data
wb_to_df(file, detect_dates = FALSE)

## -----------------------------------------------------------------------------
# return the underlying Excel formula instead of their values
wb_to_df(file, show_formula = TRUE)

## -----------------------------------------------------------------------------
# read dimension without column names
wb_to_df(file, dims = "A2:C5", col_names = FALSE)

## -----------------------------------------------------------------------------
# read dimension without column names with `wb_dims()`
wb_to_df(file, dims = wb_dims(rows = 2:5, cols = 1:3), col_names = FALSE)

## -----------------------------------------------------------------------------
# read selected cols
wb_to_df(file, cols = c("A:B", "G"))

## -----------------------------------------------------------------------------
# read selected rows
wb_to_df(file, rows = c(2, 4, 6))

## -----------------------------------------------------------------------------
# convert characters to numerics and date (logical too?)
wb_to_df(file, convert = FALSE)

## -----------------------------------------------------------------------------
# erase empty rows from dataset
tail(wb_to_df(file, sheet = 1, skip_empty_rows = TRUE))

## -----------------------------------------------------------------------------
# erase empty cols from dataset
wb_to_df(file, skip_empty_cols = TRUE)

## -----------------------------------------------------------------------------
# convert first row to rownames
wb_to_df(file, sheet = 2, dims = "C6:G9", row_names = TRUE)

## -----------------------------------------------------------------------------
# define type of the data.frame
wb_to_df(file, cols = c(2, 5), types = c("Var1" = 0, "Var3" = 1))

## -----------------------------------------------------------------------------
# start in row 5
wb_to_df(file, start_row = 5, col_names = FALSE)

## -----------------------------------------------------------------------------
# na strings
wb_to_df(file, na.strings = "")

