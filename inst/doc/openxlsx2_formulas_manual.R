## ----include = FALSE----------------------------------------------------------
knitr::opts_chunk$set(
  collapse = TRUE,
  comment = "#>"
)
set.seed(123)

## ----setup--------------------------------------------------------------------
library(openxlsx2)

## -----------------------------------------------------------------------------
# Create artificial xlsx file
wb <- wb_workbook()$add_worksheet()$add_data(x = t(c(1, 1)), col_names = FALSE)$
  add_formula(dims = "C1", x = "A1 + B1")
# Users should never modify cc as shown here
wb$worksheets[[1]]$sheet_data$cc$v[3] <- 2

# we expect a value of 2
wb_to_df(wb, col_names = FALSE)

## -----------------------------------------------------------------------------
wb$add_data(x = 2)

# we expect 3
wb_to_df(wb, col_names = FALSE)

## -----------------------------------------------------------------------------
wb_to_df(wb, col_names = FALSE, show_formula = TRUE)

## -----------------------------------------------------------------------------
wb <- wb_workbook()$add_worksheet()$
  add_data(x = head(cars))$
  add_formula(x = "SUM(A2, B2)", dims = "D2")$
  add_formula(x = "A2 + B2", dims = "D3")
# wb$open()

## -----------------------------------------------------------------------------
wb <- wb_workbook()$add_worksheet()$
  add_data(x = head(cars))$
  add_formula(x = "A2:A7 * B2:B7", dims = "C2:C7", array = TRUE)
# wb$open()

## -----------------------------------------------------------------------------
m1 <- matrix(1:6, ncol = 2)
m2 <- matrix(7:12, nrow = 2)

wb <- wb_workbook()$add_worksheet()$
  add_data(x = m1)$
  add_data(x = m2, dims = wb_dims(from_col = 4))$
  add_formula(x = "MMULT(A2:B4, D2:F3)", dims = "H2:J4", array = TRUE)
# wb$open()

## -----------------------------------------------------------------------------
# we expect to find this in D1:E1
# coef(lm(head(cars)))
wb <- wb_workbook()$add_worksheet()$
  add_data(x = head(cars))$
  add_formula(x = "LINEST(A2:A7, B2:B7, TRUE)", dims = "D2:E2", array = TRUE)
# wb$open()

## -----------------------------------------------------------------------------
wb <- wb_workbook()$add_worksheet()$
  add_data(x = head(cars))$
  add_formula(x = "SUM(ABS(A2:A7))", dims = "D2", cm = TRUE)
# wb$open()

## -----------------------------------------------------------------------------
## creating example data
company_sales <- data.frame(
    sales_price = c(20, 30, 40),
    COGS = c(5, 11, 13),
    sales_quantity = c(1, 2, 3)
)

## write in the formula
company_sales$total_sales  <- paste(paste0("A", 1:3 + 1L), paste0("C", 1:3 + 1L), sep = " * ")
## add the formula class
class(company_sales$total_sales) <- c(class(company_sales$total_sales), "formula")

## write a workbook
wb <- wb_workbook()$
  add_worksheet("Total Sales")$
  add_data_table(x = company_sales)

## -----------------------------------------------------------------------------
## Because we want the `dataTable` formula to propagate down the entire column of the data
## we can assign the formula by itself to any column and allow that single string to be repeated for each row.

## creating example data
example_data <-
  data.frame(
    sales_price = c(20, 30, 40),
    COGS = c(5, 11, 13),
    sales_quantity = c(1, 2, 3)
  )

## base R method
example_data$gross_profit       <- "daily_sales[[#This Row],[sales_price]] - daily_sales[[#This Row],[COGS]]"
example_data$total_COGS        <- "daily_sales[[#This Row],[COGS]] * daily_sales[[#This Row],[sales_quantity]]"
example_data$total_sales       <- "daily_sales[[#This Row],[sales_price]] * daily_sales[[#This Row],[sales_quantity]]"
example_data$total_gross_profit <- "daily_sales[[#This Row],[total_sales]] - daily_sales[[#This Row],[total_COGS]]"

class(example_data$gross_profit)       <- c(class(example_data$gross_profit),       "formula")
class(example_data$total_COGS)        <- c(class(example_data$total_COGS),          "formula")
class(example_data$total_sales)       <- c(class(example_data$total_sales),         "formula")
class(example_data$total_gross_profit) <- c(class(example_data$total_gross_profit), "formula")

## -----------------------------------------------------------------------------
wb$
  add_worksheet("Daily Sales")$
  add_data_table(
    x           = example_data,
    table_style = "TableStyleMedium2",
    table_name  = "daily_sales"
  )

## -----------------------------------------------------------------------------
#### sum dataTable examples
wb$add_worksheet("sum_examples")

### Note: dataTable formula do not need to be used inside of dataTables. dataTable formula are for referencing the data within the dataTable.

### Note: dataTable formula do not need to be used inside of dataTables. dataTable formula are for referencing the data within the dataTable.
sum_examples <- data.frame(
  description = c("sum_sales_price", "sum_product_Price_Quantity"),
  formula = c("", "")
)

wb$add_data(x = sum_examples)

# add formulas
wb$add_formula(x = "sum(daily_sales[[#Data],[sales_price]])", dims = "B2")
wb$add_formula(x = "sum(daily_sales[[#Data],[sales_price]] * daily_sales[[#Data],[sales_quantity]])", dims = "B3", array = TRUE)

#### dataTable referencing
wb$add_worksheet("dt_references")

### Adding the headers by themselves.
wb$add_formula(
  x = "daily_sales[[#Headers],[sales_price]:[total_gross_profit]]",
  dims = "A1:G1",
  array = TRUE
)

### Adding the raw data by reference and selecting them directly.
wb$add_formula(
  x = "daily_sales[[#Data],[sales_price]:[total_gross_profit]]",
  start_row = 2,
  dims = "A2:G4",
  array = TRUE
)
# wb$open()

