## ----setup, include = FALSE---------------------------------------------------
library(openxlsx2)
options(rmarkdown.html_vignette.check_title = FALSE)
knitr::opts_chunk$set(
  collapse = TRUE,
  comment = "#>"
)

## ----package------------------------------------------------------------------
library(openxlsx2) # openxlsx2 >= 1.26 for enharter support

## create a workbook
wb <- wb_workbook()

## ----plot---------------------------------------------------------------------
myplot <- tempfile(fileext = ".jpg")
jpeg(myplot)
print(plot(AirPassengers))
dev.off()

# Add basic plots to the workbook
wb$add_worksheet("add_image")$add_image(file = myplot)

## ----encharter----------------------------------------------------------------
if (requireNamespace("encharter")) {
library(encharter)

df_bar <- data.frame(
  Product = c("Software", "Services", "Hardware", "Support"),
  Q1      = c(310, 195, 140, 85),
  Q2      = c(340, 210, 130, 90),
  Q3      = c(375, 225, 125, 95),
  Q4      = c(420, 250, 120, 105)
)

wb <- wb_add_worksheet(wb, "add_encharter", grid_lines = FALSE)
wb <- wb_add_data_table(
  wb, sheet = "add_encharter", x = df_bar,
  dims = "A1", table_style = "TableStyleMedium2"
)
wb <- wb_set_col_widths(wb, sheet = "add_encharter", cols = 1:5, widths = c(12, 8, 8, 8, 8))
wb_df <- wb_data(wb)

chart <- ec("barChart")
chart$set_chart_title("Quarterly Revenue by Product (EUR k)", bold = TRUE)
chart$set_y_axis(min = 0, format = "#,##0", grid_lines = TRUE, grid_color = "EEEEEE")

colors    <- c("2E4057", "048A81", "E84855", "F4A261")
quarters  <- c("Q1", "Q2", "Q3", "Q4")
cols      <- c("B",  "C",  "D",  "E")
variables <- names(wb_df)
for (i in seq_along(quarters)) {
  chart$add_series(
    name   = variables[i + 1L],
    label  = variables[1L],
    data   = wb_df,
    color  = colors[i]
  )
}

chart$set_legend_style(pos = "bottom")

wb <- wb_add_encharter(wb, sheet = "add_encharter", graph = chart, dims = "G1:P18")
}

## ----chartsheet---------------------------------------------------------------
# add chartsheet
wb <- wb |>
  wb_add_chartsheet() |>
  wb_add_encharter(graph = chart)

## ----mschart------------------------------------------------------------------
if (requireNamespace("mschart")) {

library(mschart) # mschart >= 0.4 for openxlsx2 support

## create chart from mschart object (this creates new input data)
mylc <- ms_linechart(
  data = browser_ts,
  x = "date",
  y = "freq",
  group = "browser"
)

wb$add_worksheet("add_mschart")$add_mschart(dims = "A10:G25", graph = mylc)
}

