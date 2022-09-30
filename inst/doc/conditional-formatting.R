## ---- include = FALSE---------------------------------------------------------
knitr::opts_chunk$set(
  collapse = TRUE,
  comment = "#>"
)
set.seed(123)

## ----setup--------------------------------------------------------------------
library(openxlsx2)

## -----------------------------------------------------------------------------
wb <- wb_workbook()
negStyle <- create_dxfs_style(font_color = wb_colour(hex = "FF9C0006"), bgFill = wb_colour(hex = "FFFFC7CE"))
posStyle <- create_dxfs_style(font_color = wb_colour(hex = "FF006100"), bgFill = wb_colour(hex = "FFC6EFCE"))
wb$styles_mgr$add(negStyle, "negStyle")
wb$styles_mgr$add(posStyle, "posStyle")

## ----echo=FALSE, warning=FALSE------------------------------------------------
knitr::include_graphics("img/cf_cells.jpg")

## -----------------------------------------------------------------------------
wb$add_worksheet("cellIs")
wb$add_data("cellIs", -5:5)
wb$add_data("cellIs", LETTERS[1:11], startCol = 2)
wb$add_conditional_formatting(
  "cellIs",
  cols = 1,
  rows = 1:11,
  rule = "!=0",
  style = "negStyle"
)
wb$add_conditional_formatting(
  "cellIs",
  cols = 1,
  rows = 1:11,
  rule = "==0",
  style = "posStyle"
)

## ----echo=FALSE, warning=FALSE------------------------------------------------
knitr::include_graphics("img/cf_moving_row.jpg")

## -----------------------------------------------------------------------------
wb$add_worksheet("Moving Row")
wb$add_data("Moving Row", -5:5)
wb$add_data("Moving Row", LETTERS[1:11], startCol = 2)
wb$add_conditional_formatting(
  "Moving Row",
  cols = 1:2,
  rows = 1:11,
  rule = "$A1<0",
  style = "negStyle"
)
wb$add_conditional_formatting(
  "Moving Row",
  cols = 1:2,
  rows = 1:11,
  rule = "$A1>0",
  style = "posStyle"
)

## ----echo=FALSE, warning=FALSE------------------------------------------------
knitr::include_graphics("img/cf_moving_col.jpg")

## -----------------------------------------------------------------------------
wb$add_worksheet("Moving Col")
wb$add_data("Moving Col", -5:5)
wb$add_data("Moving Col", LETTERS[1:11], startCol = 2)
wb$add_conditional_formatting(
  "Moving Col",
  cols = 1:2,
  rows = 1:11,
  rule = "A$1<0",
  style = "negStyle"
)
wb$add_conditional_formatting(
  "Moving Col",
  cols = 1:2,
  rows = 1:11, 
  rule = "A$1>0",
  style = "posStyle"
)

## ----echo=FALSE, warning=FALSE------------------------------------------------
knitr::include_graphics("img/cf_dependent_on.jpg")

## -----------------------------------------------------------------------------
wb$add_worksheet("Dependent on")
wb$add_data("Dependent on", -5:5)
wb$add_data("Dependent on", LETTERS[1:11], startCol = 2)
wb$add_conditional_formatting(
  "Dependent on",
  cols = 1:2,
  rows = 1:11, 
  rule = "$A$1 < 0",
  style = "negStyle"
)
wb$add_conditional_formatting(
  "Dependent on",
  cols = 1:2,
  rows = 1:11,
  rule = "$A$1>0",
  style = "posStyle"
)

## -----------------------------------------------------------------------------
wb$add_data("Dependent on", data.frame(x = 1:10, y = runif(10)), startRow = 15)
wb$add_conditional_formatting(
  "Dependent on",
  cols = 1,
  rows = 16:25, 
  rule = "B16<0.5", 
  style = "negStyle"
)
wb$add_conditional_formatting(
  "Dependent on",
  cols = 1,
  rows = 16:25, 
  rule = "B16>=0.5", 
  style = "posStyle"
)

## ----echo=FALSE, warning=FALSE------------------------------------------------
knitr::include_graphics("img/cf_duplicates.jpg")

## -----------------------------------------------------------------------------
wb$add_worksheet("Duplicates")
wb$add_data("Duplicates", sample(LETTERS[1:15], size = 10, replace = TRUE))
wb$add_conditional_formatting(
  "Duplicates",
  cols = 1,
  rows = 1:10,
  type = "duplicatedValues"
)

## ----echo=FALSE, warning=FALSE------------------------------------------------
knitr::include_graphics("img/cf_contains_text.jpg")

## -----------------------------------------------------------------------------
fn <- function(x) paste(sample(LETTERS, 10), collapse = "-")
wb$add_worksheet("containsText")
wb$add_data("containsText", sapply(1:10, fn))
wb$add_conditional_formatting(
  "containsText",
  cols = 1,
  rows = 1:10,
  type = "contains",
  rule = "A"
)
wb$add_worksheet("notcontainsText")

## ----echo=FALSE, warning=FALSE------------------------------------------------
knitr::include_graphics("img/cf_contains_no_text.jpg")

## -----------------------------------------------------------------------------
fn <- function(x) paste(sample(LETTERS, 10), collapse = "-")
wb$add_data("notcontainsText", sapply(1:10, fn))
wb$add_conditional_formatting(
  "notcontainsText", 
  cols = 1,
  rows = 1:10, 
  type = "notContainsText", 
  rule = "A"
)

## ----echo=FALSE, warning=FALSE------------------------------------------------
knitr::include_graphics("img/cf_begins_with.jpg")

## -----------------------------------------------------------------------------
fn <- function(x) paste(sample(LETTERS, 10), collapse = "-")
wb$add_worksheet("beginsWith")
wb$add_data("beginsWith", sapply(1:100, fn))
wb$add_conditional_formatting(
  "beginsWith", 
  cols = 1,
  rows = 1:100, 
  type = "beginsWith", 
  rule = "A"
)

## ----echo=FALSE, warning=FALSE------------------------------------------------
knitr::include_graphics("img/cf_ends_with.jpg")

## -----------------------------------------------------------------------------
fn <- function(x) paste(sample(LETTERS, 10), collapse = "-")
wb$add_worksheet("endsWith")
wb$add_data("endsWith", sapply(1:100, fn))
wb$add_conditional_formatting(
  "endsWith", 
  cols = 1,
  rows = 1:100, 
  type = "endsWith",
  rule = "A"
)

## ----echo=FALSE, warning=FALSE, out.width="100%", fig.cap="Yep, that is a color scale image."----
knitr::include_graphics("img/cf_color_scale.jpg")

## -----------------------------------------------------------------------------
df <- read_xlsx(system.file("extdata", "readTest.xlsx", package = "openxlsx2"), sheet = 5)
wb$add_worksheet("colorScale", zoom = 30)
wb$add_data("colorScale", df, colNames = FALSE) ## write data.frame

## -----------------------------------------------------------------------------
wb$add_conditional_formatting(
  "colorScale",
  cols = seq_along(df), 
  rows = seq_len(nrow(df)),
  style = c("black", "white"),
  rule = c(0, 255),
  type = "colorScale"
)
wb$set_col_widths("colorScale", cols = seq_along(df), widths = 1.07)
wb$set_row_heights("colorScale", rows = seq_len(nrow(df)), heights = 7.5)

## ----echo=FALSE, warning=FALSE, out.width="100%"------------------------------
knitr::include_graphics("img/cf_databar.jpg")

## -----------------------------------------------------------------------------
wb$add_worksheet("databar")
## Databars
wb$add_data("databar", -5:5, startCol = 1)
wb <- wb_add_conditional_formatting(
  wb,
  "databar",
  cols = 1,
  rows = 1:11,
  type = "dataBar"
) ## Default colours

wb$add_data("databar", -5:5, startCol = 3)
wb <- wb_add_conditional_formatting(
  wb,
  "databar",
  cols = 3,
  rows = 1:11,
  type = "dataBar",
  params = list(
    showValue = FALSE,
    gradient = FALSE
  )
) ## Default colours

wb$add_data("databar", -5:5, startCol = 5)
wb <- wb_add_conditional_formatting(
  wb,
  "databar",
  cols = 5,
  rows = 1:11,
  type = "dataBar",
  style = c("#a6a6a6"),
  params = list(showValue = FALSE)
)

wb$add_data("databar", -5:5, startCol = 7)
wb <- wb_add_conditional_formatting(
  wb,
  "databar",
  cols = 7,
  rows = 1:11,
  type = "dataBar",
  style = c("red"),
  params = list(
    showValue = TRUE,
    gradient = FALSE
  )
)

# custom color
wb$add_data("databar", -5:5, startCol = 9)
wb <- wb_add_conditional_formatting(
  wb,
  "databar",
  cols = 9,
  rows = 1:11,
  type = "dataBar",
  style = c("#a6a6a6", "#a6a6a6"),
  params = list(showValue = TRUE, gradient = FALSE)
)

# with rule
wb$add_data(x = -5:5, startCol = 11)
wb <- wb_add_conditional_formatting(
  wb,
  "databar",
  cols = 11,
  rows = 1:11,
  type = "dataBar",
  rule = c(0, 5),
  style = c("#a6a6a6", "#a6a6a6"),
  params = list(showValue = TRUE, gradient = FALSE)
)

## ----echo=FALSE, warning=FALSE------------------------------------------------
knitr::include_graphics("img/cf_between.jpg")

## -----------------------------------------------------------------------------
wb$add_worksheet("between")
wb$add_data("between", -5:5)
wb$add_conditional_formatting(
  "between", 
  cols = 1,
  rows = 1:11,
  type = "between", 
  rule = c(-2,2)
)
wb$add_worksheet("topN")

## ----echo=FALSE, warning=FALSE------------------------------------------------
knitr::include_graphics("img/cf_top_n.jpg")

## -----------------------------------------------------------------------------
wb$add_data("topN", data.frame(x = 1:10, y = rnorm(10)))

## -----------------------------------------------------------------------------
wb$add_conditional_formatting(
  "topN", 
  cols = 1,
  rows = 2:11,
  style = "posStyle",
  type = "topN",
  params = list(rank = 5)
)

## -----------------------------------------------------------------------------
wb$add_conditional_formatting(
  "topN", 
  cols = 2,
  rows = 2:11,
  style = "posStyle", 
  type = "topN",
  params = list(rank = 20, percent = TRUE)
)
wb$add_worksheet("bottomN")

## ----echo=FALSE, warning=FALSE------------------------------------------------
knitr::include_graphics("img/cf_bottom_n.jpg")

## -----------------------------------------------------------------------------
wb$add_data("bottomN", data.frame(x = 1:10, y = rnorm(10)))

## -----------------------------------------------------------------------------
wb$add_conditional_formatting(
  "bottomN",
  cols = 1,
  rows = 2:11,
  style = "negStyle", 
  type = "bottomN",
  params = list(rank = 5)
)

## -----------------------------------------------------------------------------
wb$add_conditional_formatting(
  "bottomN", 
  cols = 2,
  rows = 2:11,
  style = "negStyle", 
  type = "bottomN",
  params = list(rank = 20, percent = TRUE)
)
wb$add_worksheet("logical operators")

## ----echo=FALSE, warning=FALSE------------------------------------------------
knitr::include_graphics("img/cf_logical_operators.jpg")

## -----------------------------------------------------------------------------
wb$add_data("logical operators", 1:10)
wb$add_conditional_formatting(
  "logical operators",
  cols = 1, 
  rows = 1:10,
  rule = "OR($A1=1,$A1=3,$A1=5,$A1=7)"
)

