## ---- include = FALSE---------------------------------------------------------
knitr::opts_chunk$set(
  collapse = TRUE,
  comment = "#>"
)
set.seed(123)

## ----setup--------------------------------------------------------------------
library(openxlsx2)

## -----------------------------------------------------------------------------
# Create artificial xlsx file
wb <- wb_workbook()$add_worksheet()$add_data(x = t(c(1, 1)), colNames = FALSE)$
  add_formula(dims = "C1", x = "A1 + B1")
# Users should never modify cc as shown here
wb$worksheets[[1]]$sheet_data$cc$v[3] <- 2

# we expect a value of 2
wb_to_df(wb, colNames = FALSE)

## -----------------------------------------------------------------------------
wb$add_data(x = 2)

# we expect 3
wb_to_df(wb, colNames = FALSE)

## -----------------------------------------------------------------------------
wb_to_df(wb, colNames = FALSE, showFormula = TRUE)

## -----------------------------------------------------------------------------
wb <- wb_workbook()$add_worksheet()$
  add_data(x = head(cars))$
  add_formula(dims = "D2", x = "SUM(A2, B2)")$
  add_formula(dims = "D3", x = "A2 + B2")
# wb$open()

## -----------------------------------------------------------------------------
wb <- wb_workbook()$add_worksheet()$
  add_data(x = head(cars))$
  add_formula(dims = "C2:C7", x = "A2:A7 * B2:B7", array = TRUE)
# wb$open()

## -----------------------------------------------------------------------------
m1 <- matrix(1:6, ncol = 2)
m2 <- matrix(7:12, nrow = 2)

wb <- wb_workbook()$add_worksheet()$
  add_data(x = m1, startCol = 1)$
  add_data(x = m2, startCol = 4)$
  add_formula(dims = "H2:J4", x = "MMULT(A2:B4, D2:F3)", array = TRUE)
# wb$open()

## -----------------------------------------------------------------------------
# we expect to find this in D1:E1
coef(lm(head(cars)))
wb <- wb_workbook()$add_worksheet()$
  add_data(x = head(cars))$
  add_formula(dims = "D2:E2", x = "LINEST(A2:A7, B2:B7, TRUE)", array = TRUE)
# wb$open()

## -----------------------------------------------------------------------------
wb <- wb_workbook()$add_worksheet()$
  add_data(x = head(cars))$
  add_formula(dims = "D2", x = 'SUM(ABS(A2:A7))', cm = TRUE)
# wb$open()

## -----------------------------------------------------------------------------
## creating example data
example_data <- data.frame(
    SalesPrice = c(20, 30, 40),
    COGS = c(5, 11, 13),
    SalesQuantity = c(1, 2, 3)
)

## write in the formula
example_data$Total_Sales  <- paste(paste0("A", 1:3 + 1L), paste0("C", 1:3 + 1L), sep = " + ")
## add the formula class
class(example_data$Total_Sales) <- c(class(example_data$Total_Sales), "formula")

## write a workbook
wb <- wb_workbook()$
  add_worksheet("Total Sales")$
  add_data_table(x = example_data)

## -----------------------------------------------------------------------------
## Because we want the `dataTable` formula to propagate down the entire column of the data
## we can assign the formula by itself to any column and allow that single string to be repeated for each row.

## creating example data
example_data <-
  data.frame(
    SalesPrice = c(20, 30, 40),
    COGS = c(5, 11, 13),
    SalesQuantity = c(1, 2, 3)
  )

## base R method
example_data$GrossProfit       <- "daily_sales[[#This Row],[SalesPrice]] - daily_sales[[#This Row],[COGS]]"
example_data$Total_COGS        <- "daily_sales[[#This Row],[COGS]] * daily_sales[[#This Row],[SalesQuantity]]"
example_data$Total_Sales       <- "daily_sales[[#This Row],[SalesPrice]] * daily_sales[[#This Row],[SalesQuantity]]"
example_data$Total_GrossProfit <- "daily_sales[[#This Row],[Total_Sales]] - daily_sales[[#This Row],[Total_COGS]]"

class(example_data$GrossProfit)       <- c(class(example_data$GrossProfit),       "formula")
class(example_data$Total_COGS)        <- c(class(example_data$Total_COGS),        "formula")
class(example_data$Total_Sales)       <- c(class(example_data$Total_Sales),       "formula")
class(example_data$Total_GrossProfit) <- c(class(example_data$Total_GrossProfit), "formula")

## -----------------------------------------------------------------------------
wb$
  add_worksheet('Daily Sales')$
  add_data_table(
    x          = example_data,
    tableStyle = "TableStyleMedium2",
    tableName  = 'daily_sales'
  )

## -----------------------------------------------------------------------------
#### sum dataTable examples
wb$add_worksheet('sum_examples')

### Note: dataTable formula do not need to be used inside of dataTables. dataTable formula are for referencing the data within the dataTable.
sum_examples <- data.frame(
    description = c("sum_SalesPrice", "sum_product_Price_Quantity"),
    formula = c(
      "sum(daily_sales[[#Data],[SalesPrice]])",
      "sum(daily_sales[[#Data],[SalesPrice]] * daily_sales[[#Data],[SalesQuantity]])"
    )
  )
class(sum_examples$formula) <- c(class(sum_examples$formula), "formula")

wb$add_data(x = sum_examples)

#### dataTable referencing
wb$add_worksheet('dt_references')

### Adding the headers by themselves.
wb$add_formula(
  x = "daily_sales[[#Headers],[SalesPrice]:[Total_GrossProfit]]",
)

### Adding the raw data by reference and selecting them directly.
wb$add_formula(
  x = "daily_sales[[#Data],[SalesPrice]:[Total_GrossProfit]]",
  startRow = 2
)
# wb$open()

