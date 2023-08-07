## ----setup, include=FALSE-----------------------------------------------------
# library(openxlsx)
library(openxlsx2)

## ----read---------------------------------------------------------------------
file <- system.file("extdata", "openxlsx2_example.xlsx", package = "openxlsx2")

## ----old_read, eval = FALSE---------------------------------------------------
#  # read in openxlsx
#  openxlsx::read.xlsx(xlsxFile = file)

## ----new_read-----------------------------------------------------------------
# read in openxlsx2
openxlsx2::read_xlsx(file = file)

## ----write--------------------------------------------------------------------
output <- temp_xlsx()

## ----old_write, eval = FALSE--------------------------------------------------
#  # write in openxlsx
#  openxlsx::write.xlsx(iris, file = output, colNames = TRUE)

## ----new_write----------------------------------------------------------------
# write in openxlsx2
openxlsx2::write_xlsx(iris, file = output, col_names = TRUE)

## ----old_workbook, eval = FALSE-----------------------------------------------
#  wb <- openxlsx::loadWorkbook(file = file)

## ----workbook-----------------------------------------------------------------
wb <- wb_load(file = file)

## ----old_style, eval = FALSE--------------------------------------------------
#  # openxlsx
#  ## Create a new workbook
#  wb <- createWorkbook(creator = "My name here")
#  addWorksheet(wb, "Expenditure", gridLines = FALSE)
#  writeData(wb, sheet = 1, USPersonalExpenditure, rowNames = TRUE)
#  
#  ## style for body
#  bodyStyle <- createStyle(border = "TopBottom", borderColor = "#4F81BD")
#  addStyle(wb, sheet = 1, bodyStyle, rows = 2:6, cols = 1:6, gridExpand = TRUE)
#  
#  ## set column width for row names column
#  setColWidths(wb, 1, cols = 1, widths = 21)

## ----new_style----------------------------------------------------------------
# openxlsx2 chained
border_color <- wb_color(hex = "#4F81BD")
wb <- wb_workbook(creator = "My name here")$
  add_worksheet("Expenditure", grid_lines = FALSE)$
  add_data(x = USPersonalExpenditure, row_names = TRUE)$
  add_border( # add the outer and inner border
    dims = "A1:F6",
    top_border = "thin", top_color = border_color,
    bottom_border = "thin", bottom_color = border_color,
    inner_hgrid = "thin", inner_hcolor = border_color,
    left_border = "", right_border = ""
  )$
  set_col_widths( # set column width
    cols = 1:6,
    widths = c(20, rep(10, 5))
  )$ # remove the value in A1
  add_data(dims = "A1", x = "")

## ----new_style_pipes----------------------------------------------------------
# openxlsx2 with pipes
border_color <- wb_color(hex = "4F81BD")
wb <- wb_workbook(creator = "My name here") %>%
  wb_add_worksheet(sheet = "Expenditure", grid_lines = FALSE) %>%
  wb_add_data(x = USPersonalExpenditure, row_names = TRUE) %>%
  wb_add_border( # add the outer and inner border
    dims = "A1:F6",
    top_border = "thin", top_color = border_color,
    bottom_border = "thin", bottom_color = border_color,
    inner_hgrid = "thin", inner_hcolor = border_color,
    left_border = "", right_border = ""
  ) %>%
  wb_set_col_widths( # set column width
    cols = 1:6,
    widths = c(20, rep(10, 5))
  ) %>% # remove the value in A1
  wb_add_data(dims = "A1", x = "")

## ----pipe_chain---------------------------------------------------------------
# openxlsx2
wbp <- wb_workbook() %>% wb_add_worksheet()
wbc <- wb_workbook()$add_worksheet()

# need to assign wbp
wbp <- wbp %>% wb_add_data(x = iris)
wbc$add_data(x = iris)

## ----new_cf-------------------------------------------------------------------
# openxlsx2 with chains
wb <- wb_workbook()$
  add_worksheet("a")$
  add_data(x = 1:4, col_names = FALSE)$
  add_conditional_formatting(dims = "A1:A4", rule = ">2")

# openxlsx2 with pipes
wb <- wb_workbook() %>%
  wb_add_worksheet("a") %>%
  wb_add_data(x = 1:4, col_names = FALSE) %>%
  wb_add_conditional_formatting(dims = "A1:A4", rule = ">2")

## ----old_dv, eval = FALSE-----------------------------------------------------
#  # openxlsx
#  wb <- createWorkbook()
#  addWorksheet(wb, "Sheet 1")
#  writeDataTable(wb, 1, x = iris[1:30, ])
#  dataValidation(wb, 1,
#    col = 1:3, rows = 2:31, type = "whole",
#    operator = "between", value = c(1, 9)
#  )

## ----new_dv-------------------------------------------------------------------
# openxlsx2 with chains
wb <- wb_workbook()$
  add_worksheet("Sheet 1")$
  add_data_table(1, x = iris[1:30, ])$
  add_data_validation(1,
    dims = wb_dims(rows = 2:31, cols = 1:3),
    # alternatively, dims can also be "A2:C31" if you know the span in your Excel workbook.
    type = "whole",
    operator = "between",
    value = c(1, 9)
  )

# openxlsx2 with pipes
wb <- wb_workbook() %>%
  wb_add_worksheet("Sheet 1") %>%
  wb_add_data_table(1, x = iris[1:30, ]) %>%
  wb_add_data_validation(
    sheet = 1,
    dims = "A2:C31", # alternatively, dims = wb_dims(rows = 2:31, cols = 1:3)
    type = "whole",
    operator = "between",
    value = c(1, 9)
  )

