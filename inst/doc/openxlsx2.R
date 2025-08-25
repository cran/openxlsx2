## ----setup, include = FALSE---------------------------------------------------
library(openxlsx2)
options(rmarkdown.html_vignette.check_title = FALSE)
knitr::opts_chunk$set(
  collapse = TRUE,
  comment = "#>"
)
modern_r <- getRversion() >= "4.1.0"

## ----eval = FALSE-------------------------------------------------------------
# install.packages("openxlsx2")

## ----eval = FALSE-------------------------------------------------------------
# library(openxlsx2)

## ----eval = FALSE-------------------------------------------------------------
# wb <- wb_load("your_file.xlsx")

## ----eval = FALSE-------------------------------------------------------------
# wb <- wb_workbook() |> wb_add_worksheet() |> wb_add_data(x = your_data)

## ----eval = FALSE-------------------------------------------------------------
# wb <- wb_add_data(wb_add_worksheet(wb_workbook()), x = your_data)

## ----eval = modern_r----------------------------------------------------------
wb <- wb_workbook() |> wb_add_worksheet() |> wb_add_data(x = mtcars)

## -----------------------------------------------------------------------------
wb

## ----eval = modern_r----------------------------------------------------------
wb <- wb_workbook() |> wb_add_worksheet() |> wb_add_worksheet() |> wb_add_data(x = mtcars)

## -----------------------------------------------------------------------------
wb

## ----eval = modern_r----------------------------------------------------------
wb |> wb_to_df() |> head()

## ----eval = modern_r----------------------------------------------------------
wb |> wb_to_df(sheet = "Sheet 2") |> head()

## ----eval = FALSE-------------------------------------------------------------
# wb |> wb_save(file = "my_first_worksheet.xlsx")

## ----eval = FALSE-------------------------------------------------------------
# wb |> wb_open()

## ----eval = FALSE-------------------------------------------------------------
# wb <- wb_workbook()
# wb_add_worksheet(wb, sheet = "USexp")
# wb_add_data(wb, "USexp", USPersonalExpenditure)

## ----echo = FALSE-------------------------------------------------------------
wb <- wb_workbook()
wb_add_worksheet(wb, sheet = "USexp")
wb_add_data(wb, "USexp", USPersonalExpenditure) |> try()

## ----eval = modern_r----------------------------------------------------------
wb |> wb_get_sheet_names()

## -----------------------------------------------------------------------------
wb <- wb_workbook()
wb <- wb_add_worksheet(wb, sheet = "USexp")
wb <- wb_add_data(wb, "USexp", USPersonalExpenditure)
wb_get_sheet_names(wb)
wb_to_df(wb)

## -----------------------------------------------------------------------------
wb <- wb_workbook()$add_worksheet("USexp")$add_data(x = USPersonalExpenditure)
wb$to_df()

## -----------------------------------------------------------------------------
# the file we are going to load
file <- system.file("extdata", "openxlsx2_example.xlsx", package = "openxlsx2")
# loading the file into the workbook
wb <- wb_load(file = file)

## ----eval = FALSE-------------------------------------------------------------
# write_xlsx(x = mtcars, file = "mtcars.xlsx")

## ----eval = FALSE-------------------------------------------------------------
# # replace the existing file
# wb$save("mtcars.xlsx")
# 
# # do not overwrite the existing file
# try(wb$save("mtcars.xlsx", overwrite = FALSE))

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

