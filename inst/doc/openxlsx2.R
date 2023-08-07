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
wb_to_df(file, sheet = 1, skip_empty_rows = TRUE) |> tail()

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

## -----------------------------------------------------------------------------
# the file we are going to load
file <- system.file("extdata", "openxlsx2_example.xlsx", package = "openxlsx2")
# loading the file into the workbook
wb <- wb_load(file = file)

## ---- eval = FALSE------------------------------------------------------------
#  write_xlsx(x = mtcars, file = "mtcars.xlsx")

## ---- eval = FALSE------------------------------------------------------------
#  # replace the existing file
#  wb$save("mtcars.xlsx")
#  
#  # do not overwrite the existing file
#  try(wb$save("mtcars.xlsx", overwrite = FALSE))

## -----------------------------------------------------------------------------
# various options
wb_dims(from_row = 4)

wb_dims(rows = 4, cols = 4)
wb_dims(rows = 4, cols = "D")

wb_dims(rows = 4:10, cols = 5:9)

wb_dims(rows = 4:10, cols = "A:D") # same as below
wb_dims(rows = seq_len(7), cols = seq_len(4), from_row = 4)
# 10 rows and 15 columns from indice B2.
wb_dims(rows = 1:10, cols = 1:15, from_col = "B", from_row = 2)

# data + col names
wb_dims(x = mtcars)
# only data
wb_dims(x = mtcars, select = "data")

# The dims of the values of a column in `x`
wb_dims(x = mtcars, cols = "cyl")
# a column in `x` with the column name
wb_dims(x = mtcars, cols = "cyl", select = "x")
# rows in `x`
wb_dims(x = mtcars)

# in a wb chain
wb <- wb_workbook()$
  add_worksheet()$
  add_data(x = mtcars)$
  add_fill(
    dims = wb_dims(x = mtcars, rows = 1:5), # only 1st 5 rows of x data
    color = wb_color("yellow")
  )$
  add_fill(
    dims = wb_dims(x = mtcars, select = "col_names"), # only column names
    color = wb_color("cyan2")
  )

# or if the data's first coord needs to be located in B2.

wb_dims_custom <- function(...) {
  wb_dims(x = mtcars, from_col = "B", from_row = 2, ...)
}
wb <- wb_workbook()$
  add_worksheet()$
  add_data(x = mtcars, dims = wb_dims_custom())$
  add_fill(
    dims = wb_dims_custom(rows = 1:5),
    color = wb_color("yellow")
  )$
  add_fill(
    dims = wb_dims_custom(select = "col_names"),
    color = wb_color("cyan2")
  )

