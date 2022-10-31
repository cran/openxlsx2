## ----setup, include=FALSE-----------------------------------------------------
# library(openxlsx)
library(openxlsx2)

## ----read---------------------------------------------------------------------
xlsxFile <- system.file("extdata", "readTest.xlsx", package = "openxlsx2")

## ----old_read, eval = FALSE---------------------------------------------------
#  # read in openxlsx
#  openxlsx::read.xlsx(xlsxFile)

## ----new_read-----------------------------------------------------------------
# read in openxlsx2
openxlsx2::read_xlsx(xlsxFile)

## ----write--------------------------------------------------------------------
output <- temp_xlsx()

## ----old_write, eval = FALSE--------------------------------------------------
#  # write in openxlsx
#  openxlsx::write.xlsx(iris, file = output, colNames = TRUE)

## ----new_write----------------------------------------------------------------
# write in openxlsx2
openxlsx2::write_xlsx(iris, file = output, colNames = TRUE)

## ----old_workbook, eval = FALSE-----------------------------------------------
#  wb <- loadWorkbook(xlsxFile)

## ----workbook-----------------------------------------------------------------
wb <- wb_load(xlsxFile)

## ----old_style, eval = FALSE--------------------------------------------------
#  ## Create a new workbook
#  wb <- createWorkbook("My name here")
#  addWorksheet(wb, "Expenditure", gridLines = FALSE)
#  writeData(wb, sheet = 1, USPersonalExpenditure, rowNames = TRUE)
#  
#  ## style for body
#  bodyStyle <- createStyle(border = "TopBottom", borderColour = "#4F81BD")
#  addStyle(wb, sheet = 1, bodyStyle, rows = 2:6, cols = 1:6, gridExpand = TRUE)
#  
#  ## set column width for row names column
#  setColWidths(wb, 1, cols = 1, widths = 21)

## ----new_style----------------------------------------------------------------
border_color <- wb_colour(hex = "FF4F81BD")
wb <- wb_workbook("My name here")$
  add_worksheet("Expenditure", gridLines = FALSE)$
  add_data(x = USPersonalExpenditure, rowNames = TRUE)$
  add_border( # add the outer and inner border
    dims = "A1:F6",
    top_border = "thin", top_color = border_color,
    bottom_border = "thin", bottom_color = border_color,
    inner_hgrid = "thin", inner_hcolor = border_color,
    left_border = "", right_border = ""
  )$
  set_col_widths( # set column width
    cols = 1:6,
    widths = c("20", rep("10", 5))
  )$ # remove the value in A1
  add_data(dims = "A1", x = "")

## ----new_cf-------------------------------------------------------------------
wb <- wb_workbook()$
  add_worksheet("a")$
  add_data(x = 1:4, colNames = FALSE)$
  add_conditional_formatting(cols = 1, rows = 1:4, rule = ">2")

## ----old_dv, eval = FALSE-----------------------------------------------------
#  wb <- createWorkbook()
#  addWorksheet(wb, "Sheet 1")
#  writeDataTable(wb, 1, x = iris[1:30, ])
#  dataValidation(wb, 1,
#    col = 1:3, rows = 2:31, type = "whole",
#    operator = "between", value = c(1, 9)
#  )

## ----new_dv-------------------------------------------------------------------
wb <- wb_workbook()$
  add_worksheet("Sheet 1")$
  add_data_table(1, x = iris[1:30, ])$
  add_data_validation(1,
    col = 1:3, rows = 2:31, type = "whole",
    operator = "between", value = c(1, 9)
  )

