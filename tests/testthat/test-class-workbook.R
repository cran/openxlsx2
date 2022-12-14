test_that("Workbook class", {
  expect_null(assert_workbook(wb_workbook()))
})


test_that("wb_set_col_widths", {
# TODO use wb$wb_set_col_widths()

  wb <- wbWorkbook$new()
  wb$add_worksheet("test")
  wb$add_data("test", mtcars)

  # set column width to 12
  expect_silent(wb$set_col_widths("test", widths = 12L, cols = seq_along(mtcars)))
  expect_equal(
    "<col min=\"1\" max=\"11\" bestFit=\"1\" customWidth=\"1\" hidden=\"false\" width=\"12.711\"/>",
    wb$worksheets[[1]]$cols_attr
  )

  # wrong sheet
  expect_error(wb$set_col_widths("test2", widths = 12L, cols = seq_along(mtcars)))

  # reset the column with, we do not provide an option ot remove the column entry
  expect_silent(wb$set_col_widths("test", cols = seq_along(mtcars)))
  expect_equal(
    "<col min=\"1\" max=\"11\" bestFit=\"1\" customWidth=\"1\" hidden=\"false\" width=\"9.141\"/>",
    wb$worksheets[[1]]$cols_attr
  )

  # create column width for column 25
  expect_silent(wb$set_col_widths("test", cols = "Y", widths = 22))
  expect_equal(
    c("<col min=\"1\" max=\"11\" bestFit=\"1\" customWidth=\"1\" hidden=\"false\" width=\"9.141\"/>",
      "<col min=\"12\" max=\"24\" width=\"8.43\"/>",
      "<col min=\"25\" max=\"25\" bestFit=\"1\" customWidth=\"1\" hidden=\"false\" width=\"22.711\"/>"),
    wb$worksheets[[1]]$cols_attr
  )

  # a few more errors
  expect_error(wb$set_col_widths("test", cols = "Y", width = 1:2))
  expect_error(wb$set_col_widths("test", cols = "Y", hidden = 1:2))




  wb <- wb_workbook()$
    add_worksheet()$
    set_col_widths(cols = 1:10, width = (8:17) + .5)$
    add_data(x = rbind(8:17), colNames = FALSE)

  exp <- c(
    "<col min=\"1\" max=\"1\" bestFit=\"1\" customWidth=\"1\" hidden=\"false\" width=\"9.211\"/>",
    "<col min=\"2\" max=\"2\" bestFit=\"1\" customWidth=\"1\" hidden=\"false\" width=\"10.211\"/>",
    "<col min=\"3\" max=\"3\" bestFit=\"1\" customWidth=\"1\" hidden=\"false\" width=\"11.211\"/>",
    "<col min=\"4\" max=\"4\" bestFit=\"1\" customWidth=\"1\" hidden=\"false\" width=\"12.211\"/>",
    "<col min=\"5\" max=\"5\" bestFit=\"1\" customWidth=\"1\" hidden=\"false\" width=\"13.211\"/>",
    "<col min=\"6\" max=\"6\" bestFit=\"1\" customWidth=\"1\" hidden=\"false\" width=\"14.211\"/>",
    "<col min=\"7\" max=\"7\" bestFit=\"1\" customWidth=\"1\" hidden=\"false\" width=\"15.211\"/>",
    "<col min=\"8\" max=\"8\" bestFit=\"1\" customWidth=\"1\" hidden=\"false\" width=\"16.211\"/>",
    "<col min=\"9\" max=\"9\" bestFit=\"1\" customWidth=\"1\" hidden=\"false\" width=\"17.211\"/>",
    "<col min=\"10\" max=\"10\" bestFit=\"1\" customWidth=\"1\" hidden=\"false\" width=\"18.211\"/>"
  )
  got <- wb$worksheets[[1]]$cols_attr
  expect_equal(exp, got)

})


# order -------------------------------------------------------------------

test_that("$set_order() works", {
  wb <- wb_workbook()
  wb$add_worksheet("a")
  wb$add_worksheet("b")
  wb$add_worksheet("c")

  expect_identical(wb$sheetOrder, 1:3)
  exp <- letters[1:3]
  names(exp) <- exp
  expect_identical(wb$get_sheet_names(), exp)

  wb$set_order(3:1)
  expect_identical(wb$sheetOrder, 3:1)
  exp <- letters[3:1]
  names(exp) <- exp
  expect_identical(wb$get_sheet_names(), exp)
})


# sheet names -------------------------------------------------------------

test_that("$set_sheet_names() and $get_sheet_names() work", {
  wb <- wb_workbook()$add_worksheet()$add_worksheet()
  wb$set_sheet_names(new = c("a", "b & c"))

  # return a names character vector
  res <- wb$get_sheet_names()
  exp <- c(a = "a", "b & c" = replace_legal_chars("b & c"))
  expect_identical(res, exp)

  # should be able to check the original values, too
  res <- wb$.__enclos_env__$private$get_sheet_index("b & c")
  expect_identical(res, 2L)

  # make sure that it works silently
  wb <- wb_load(file = system.file("extdata", "oxlsx2_sheet.xlsx", package = "openxlsx2"))
  expect_silent(wb$set_sheet_names(old = "SUM", new = "Sheet 1"))

  exp <- c(`Sheet 1` = "Sheet 1")
  got <- wb$get_sheet_names()
  expect_equal(exp, got)
})

# data validation ---------------------------------------------------------


test_that("data validation", {

  temp <- temp_xlsx()

  df <- data.frame(
    "d" = as.Date("2016-01-01") + -5:5,
    "t" = as.POSIXct("2016-01-01") + -5:5 * 10000
  )

  wb <- wb_workbook()$
    add_worksheet("Sheet 1")$
    add_data_table(x = iris)$
    # whole numbers are fine
    add_data_validation(col = 1:3, rows = 2:151, type = "whole",
                        operator = "between", value = c(1, 9)
    )$
    # text width 7-9 is fine
    add_data_validation(col = 5, rows = 2:151, type = "textLength",
                        operator = "between", value = c(7, 9)
    )$
    ## Date and Time cell validation
    add_worksheet("Sheet 2")$
    add_data_table(x = df)$
    # date >= 2016-01-01 is fine
    add_data_validation(col = 1, rows = 2:12, type = "date",
                        operator = "greaterThanOrEqual", value = as.Date("2016-01-01")
    )$
    # a few timestamps are fine
    add_data_validation(col = 2, rows = 2:12, type = "time",
                        operator = "between", value = df$t[c(4, 8)]
    )$
    ## validate list: validate inputs on one sheet with another
    add_worksheet("Sheet 3")$
    add_data_table(x = iris[1:30, ])$
    add_worksheet("Sheet 4")$
    add_data(x = sample(iris$Sepal.Length, 10))$
    add_data_validation("Sheet 3", col = 1, rows = 2:31, type = "list",
                        value = "'Sheet 4'!$A$1:$A$10")

  exp <- c(
    "<dataValidation type=\"whole\" operator=\"between\" allowBlank=\"1\" showInputMessage=\"1\" showErrorMessage=\"1\" sqref=\"A2:C151\"><formula1>1</formula1><formula2>9</formula2></dataValidation>",
    "<dataValidation type=\"textLength\" operator=\"between\" allowBlank=\"1\" showInputMessage=\"1\" showErrorMessage=\"1\" sqref=\"E2:E151\"><formula1>7</formula1><formula2>9</formula2></dataValidation>"
  )
  got <- wb$worksheets[[1]]$dataValidations
  expect_equal(exp, got)


  exp <- c(
    "<dataValidation type=\"date\" operator=\"greaterThanOrEqual\" allowBlank=\"1\" showInputMessage=\"1\" showErrorMessage=\"1\" sqref=\"A2:A12\"><formula1>42370</formula1></dataValidation>",
    "<dataValidation type=\"time\" operator=\"between\" allowBlank=\"1\" showInputMessage=\"1\" showErrorMessage=\"1\" sqref=\"B2:B12\"><formula1>42369.7685185185</formula1><formula2>42370.2314814815</formula2></dataValidation>"
  )
  got <- wb$worksheets[[2]]$dataValidations
  expect_equal(exp, got)


  exp <- c(
    "<dataValidation type=\"list\" allowBlank=\"1\" showInputMessage=\"1\" showErrorMessage=\"1\" sqref=\"A2:A31\"><formula1>'Sheet 4'!$A$1:$A$10</formula1></dataValidation>"
  )
  got <- wb$worksheets[[3]]$dataValidations
  expect_equal(exp, got)

  wb$save(temp)

  wb2 <- wb_load(temp)

  # wb2$add_data_validation("Sheet 3", col = 2, rows = 2:31, type = "list",
  #                         value = "'Sheet 4'!$A$1:$A$10")
  # wb2$save(temp)

  expect_equal(
    wb$worksheets[[1]]$dataValidations,
    wb2$worksheets[[1]]$dataValidations
  )

  expect_equal(
    wb$worksheets[[2]]$dataValidations,
    wb2$worksheets[[2]]$dataValidations
  )

  expect_equal(
    wb$worksheets[[3]]$dataValidations,
    wb2$worksheets[[3]]$dataValidations
  )

  wb2$add_data_validation("Sheet 3", col = 2, rows = 2:31, type = "list",
                          value = "'Sheet 4'!$A$1:$A$10")

  exp <- c(
    "<dataValidation type=\"list\" allowBlank=\"1\" showInputMessage=\"1\" showErrorMessage=\"1\" sqref=\"A2:A31\"><formula1>'Sheet 4'!$A$1:$A$10</formula1></dataValidation>",
    "<dataValidation type=\"list\" allowBlank=\"1\" showInputMessage=\"1\" showErrorMessage=\"1\" sqref=\"B2:B31\"><formula1>'Sheet 4'!$A$1:$A$10</formula1></dataValidation>"
  )
  got <- wb2$worksheets[[3]]$dataValidations
  expect_equal(exp, got)

  ### tests if conditions

  # test col2int
  wb <- wb_workbook()$
    add_worksheet("Sheet 1")$
    add_data_table(x = head(iris))$
    # whole numbers are fine
    add_data_validation(col = "A", rows = 2:151, type = "whole",
                        operator = "between", value = c(1, 9)
    )

  exp <- "<dataValidation type=\"whole\" operator=\"between\" allowBlank=\"1\" showInputMessage=\"1\" showErrorMessage=\"1\" sqref=\"A2:A151\"><formula1>1</formula1><formula2>9</formula2></dataValidation>"
  got <- wb$worksheets[[1]]$dataValidations
  expect_equal(exp, got)


  # to many values
  expect_error(
    wb <- wb_workbook()$
    add_worksheet("Sheet 1")$
    add_data_table(x = head(iris))$
    add_data_validation(col = "A", rows = 2:151, type = "whole",
                        operator = "between", value = c(1, 9, 19)
    ),
    "length <= 2"
  )

  # wrong type
  expect_error(
    wb <- wb_workbook()$
      add_worksheet("Sheet 1")$
      add_data_table(x = head(iris))$
      add_data_validation(col = "A", rows = 2:151, type = "even",
                          operator = "between", value = c(1, 9)
      ),
    "Invalid 'type' argument!"
  )

  # wrong operator
  expect_error(
    wb <- wb_workbook()$
      add_worksheet("Sheet 1")$
      add_data_table(x = head(iris))$
      add_data_validation(col = "A", rows = 2:151, type = "whole",
                          operator = "lower", value = c(1, 9)
      ),
    "Invalid 'operator' argument!"
  )

  # wrong value for date
  expect_error(
    wb <- wb_workbook()$
      add_worksheet("Sheet 1")$
      add_data_table(x = head(iris))$
      # whole numbers are fine
      add_data_validation(col = 1, rows = 2:12, type = "date",
                          operator = "greaterThanOrEqual", value = 7
      ),
    "If type == 'date' value argument must be a Date vector"
  )

  # wrong value for time
  expect_error(
    wb <- wb_workbook()$
      add_worksheet("Sheet 1")$
      add_data_table(x = head(iris))$
      # whole numbers are fine
      add_data_validation(col = 1, rows = 2:12, type = "time",
                          operator = "greaterThanOrEqual", value = 7
      ),
    "If type == 'time' value argument must be a POSIXct or POSIXlt vector."
  )


  # some more options
  wb <- wb_workbook()$
    add_worksheet("Sheet 1")$
    add_data(x = c(-1:1), colNames = FALSE)$
    # whole numbers are fine
    add_data_validation(col = 1, rows = 1:3, type = "whole",
                        operator = "greaterThan", value = c(0),
                        errorStyle = "information", errorTitle = "ERROR!",
                        error = "Some error ocurred!",
                        promptTitle = "PROMPT!",
                        prompt = "Choose something!"
    )

  exp <- "<dataValidation type=\"whole\" operator=\"greaterThan\" allowBlank=\"1\" showInputMessage=\"1\" showErrorMessage=\"1\" sqref=\"A1:A3\" errorStyle=\"information\" errorTitle=\"ERROR!\" error=\"Some error ocurred!\" promptTitle=\"PROMPT!\" prompt=\"Choose something!\"><formula1>0</formula1></dataValidation>"
  got <- wb$worksheets[[1]]$dataValidations
  expect_equal(exp, got)

  # add custom data
  wb <- wb_workbook()$
    add_worksheet("Sheet 1")$
    add_data(x = data.frame(x = 1, y = 2), colNames = FALSE)$
    # whole numbers are fine
    add_data_validation(col = 1, rows = 1:3, type = "custom", value = "A1=B1")

  exp <- "<dataValidation type=\"custom\" allowBlank=\"1\" showInputMessage=\"1\" showErrorMessage=\"1\" sqref=\"A1:A3\"><formula1>A1=B1</formula1></dataValidation>"
  got <- wb$worksheets[[1]]$dataValidations
  expect_equal(exp, got)

})


test_that("clone worksheet", {

  ## Dummy tests - not sure how to test these from R ##

  # # clone chartsheet ----------------------------------------------------
  fl <- system.file("extdata", "mtcars_chart.xlsx", package = "openxlsx2")
  wb <- wb_load(fl)
  # wb$get_sheet_names() # chartsheet has no named name?
  expect_silent(wb$clone_worksheet(1, "Clone 1"))
  expect_true(inherits(wb$worksheets[[5]], "wbChartSheet"))
  # wb$open()

  # clone pivot table and drawing -----------------------------------------
  fl <- system.file("extdata", "loadExample.xlsx", package = "openxlsx2")
  wb <- wb_load(fl)
  expect_silent(wb$clone_worksheet(4, "Clone 1"))

  # sheets 4 & 5 both reference the same pivot table in different drawing
  # once the file is opened, both pivot tables behave independently
  exp <- c(
    "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable\" Target=\"../pivotTables/pivotTable2.xml\"/>",
    "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing\" Target=\"../drawings/drawing5.xml\"/>"
  )
  got <- wb$worksheets_rels[[5]]
  expect_equal(exp, got)
  # wb$open()

  # clone drawing ---------------------------------------------------------
  fl <- system.file("extdata", "loadExample.xlsx", package = "openxlsx2")
  wb <- wb_load(fl)
  expect_silent(wb$clone_worksheet("testing", "Clone1"))

  expect_false(identical(wb$worksheets_rels[2], wb$worksheets_rels[5]))
  # wb$open()

  # clone sheet with table ------------------------------------------------
  fl <- system.file("extdata", "tableStyles.xlsx", package = "openxlsx2")
  wb <- wb_load(fl)
  expect_silent(wb$clone_worksheet(1, "clone"))

  expect_false(identical(wb$tables$tab_xml[1], wb$tables$tab_xml[2]))
  # wb$open()

  # clone sheet with chart ------------------------------------------------
  fl <- system.file("extdata", "mtcars_chart.xlsx", package = "openxlsx2")
  wb <- wb_load(fl)
  wb$clone_worksheet(2, "Clone 1")

  expect_true(grepl("test", wb$charts$chart[2]))
  expect_true(grepl("'Clone 1'", wb$charts$chart[3]))
  # wb$open()

  # clone slicer ----------------------------------------------------------
  fl <- system.file("extdata", "loadExample.xlsx", package = "openxlsx2")
  wb <- wb_load(fl)
  expect_warning(wb$clone_worksheet("IrisSample", "Clone1"),
                 "Cloning slicers is not yet supported. It will not appear on the sheet.")
  # wb$open()

})

test_that("set and remove row heights work", {

  ## add row heights
  wb <- wb_workbook()$
    add_worksheet()$
    set_row_heights(
      rows = c(1, 4, 22, 2, 19),
      heights = c(24, 28, 32, 42, 33)
    )

  exp <- structure(
    list(
      customHeight = c("1", "1", "1", "1", "1"),
      ht = c("24", "42", "28", "33", "32"),
      r = c("1", "2", "4", "19", "22")
    ),
    row.names = c(1L, 2L, 4L, 19L, 22L),
    class = "data.frame"
  )
  got <- wb$worksheets[[1]]$sheet_data$row_attr[c(1, 2, 4, 19, 22), c("customHeight", "ht", "r")]
  expect_equal(exp, got)

  ## remove row heights
  wb$remove_row_heights(rows = 1:21)
  exp <- structure(
    list(
      customHeight = c("", "", "", "", "1"),
      ht = c("", "", "", "", "32"),
      r = c("1", "2", "4", "19", "22")
    ),
    row.names = c(1L, 2L, 4L, 19L, 22L),
    class = "data.frame"
  )
  got <- wb$worksheets[[1]]$sheet_data$row_attr[c(1, 2, 4, 19, 22), c("customHeight", "ht", "r")]
  expect_equal(exp, got)

  expect_warning(
    wb$add_worksheet()$remove_row_heights(rows = 1:3),
    "There are no initialized rows on this sheet"
  )

  wb <- wb_workbook()$
    add_worksheet()$
    add_data(x = mtcars)$
    set_row_heights(rows = 5:15, hidden = TRUE)

  exp <- structure(
    c(22L, `1` = 11L),
    dim = 2L,
    dimnames = structure(
      list(
        c("", "1")
      ),
      names = ""
    ),
    class = "table"
  )
  got <- table(wb$worksheets[[1]]$sheet_data$row_attr$hidden)

  expect_equal(exp, got)

})

test_that("add_drawing works", {

  skip_if_not_installed("rvg")
  skip_if_not_installed("ggplot2")

  require(rvg)
  require(ggplot2)

  tmp <- tempfile(fileext = "drawing.xml")

  ## rvg example
  dml_xlsx(file =  tmp, fonts = list(sans = "Bradley Hand"))
  print(
    ggplot(data = iris,
           mapping = aes(x = Sepal.Length, y = Petal.Width)) +
      geom_point() + labs(title = "With font Bradley Hand") +
      theme_minimal(base_family = "sans", base_size = 18)
  )
  dev.off()

  wb <- wb_workbook()$
    add_worksheet()$
    add_drawing(xml = tmp)$
    add_drawing(xml = tmp, dims = "A1:H10")$
    add_drawing(xml = tmp, dims = "L1")$
    add_drawing(xml = tmp, dims = NULL)$
    add_drawing(xml = tmp, dims = "L19")

  expect_equal(1L, length(wb$drawings))

})

test_that("add_drawing works", {

  skip_if_not_installed("mschart")

  require(mschart)

  # write data starting at B2
  wb <- wb_workbook()$add_worksheet()$
    add_data(x = mtcars, dims = "B2")$
    add_data(x = data.frame(name = rownames(mtcars)), dims = "A2")

  # create wb_data object this will tell this mschart from this PR to create a file corresponding to openxlsx2
  dat <- wb_data(wb, 1)
  expect_equal(c(32L, 12L), dim(dat))

  dat <- wb_data(wb, 1, dims = "A2:G6")

  exp <- structure(
    list(
      name = c("Mazda RX4", "Mazda RX4 Wag", "Datsun 710", "Hornet 4 Drive"),
      mpg = c(21, 21, 22.8, 21.4),
      cyl = c(6, 6, 4, 6),
      disp = c(160, 160, 108, 258),
      hp = c(110, 110, 93, 110),
      drat = c(3.9, 3.9, 3.85, 3.08),
      wt = c(2.62, 2.875, 2.32, 3.215)
    ),
    row.names = 3:6,
    class = c("data.frame", "wb_data"),
    tt = structure(
      list(
        name = c("s", "s", "s", "s"),
        mpg = c("n", "n", "n", "n"),
        cyl = c("n", "n", "n", "n"),
        disp = c("n", "n", "n", "n"),
        hp = c("n", "n", "n", "n"),
        drat = c("n", "n", "n", "n"),
        wt = c("n", "n", "n", "n")
      ),
      row.names = 3:6,
      class = "data.frame"),
    types = c(A = 0, B = 1, C = 1, D = 1, E = 1, F = 1, G = 1),
    dims = structure(
      list(
        A = c("A2", "A3", "A4", "A5", "A6"),
        B = c("B2", "B3", "B4", "B5", "B6"),
        C = c("C2", "C3", "C4", "C5", "C6"),
        D = c("D2", "D3", "D4", "D5", "D6"),
        E = c("E2", "E3", "E4", "E5", "E6"),
        F = c("F2", "F3", "F4", "F5", "F6"),
        G = c("G2", "G3", "G4", "G5", "G6")
      ),
      row.names = 2:6,
      class = "data.frame"),
    sheet = "Sheet 1")

  expect_equal(exp, dat)

  # call ms_scatterplot
  scatter_plot <- ms_scatterchart(
    data = dat,
    x = "mpg",
    y = c("disp", "hp"),
    labels = c("disp", "hp")
  )

  # add the scatterplots to the data
  wb <- wb %>%
    wb_add_mschart(dims = "F4:L20", graph = scatter_plot)

  expect_equal(1L, NROW(wb$charts))

  chart_01 <- ms_linechart(
    data = us_indus_prod,
    x = "date", y = "value",
    group = "type"
  )

  wb$add_worksheet()
  wb$add_mschart(dims = "F4:L20", graph = chart_01)

  exp <- list(
    "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart\" Target=\"../charts/chart1.xml\"/>",
    "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart\" Target=\"../charts/chart2.xml\"/>"
  )
  got <- wb$drawings_rels
  expect_equal(exp, got)


  # write data starting at B2
  wb <- wb_workbook()$
    add_worksheet()$add_data(x = mtcars)$
    add_worksheet()$add_data(x = mtcars)$
    add_worksheet()$add_data(x = mtcars)$
    add_worksheet()$add_data(x = mtcars)$
    add_mschart(dims = "F4:L20", 2, graph = chart_01)$
    add_mschart(dims = "F4:L20", 3, graph = chart_01)

  exp <- list(
    character(0),
    "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing\" Target=\"../drawings/drawing2.xml\"/>",
    "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing\" Target=\"../drawings/drawing3.xml\"/>",
    character(0)
  )
  got <- wb$worksheets_rels
  expect_equal(exp, got)

  ## write different anchors
  wb <- wb_workbook()$
    add_worksheet()$add_data(x = mtcars)

  scatter_plot <- ms_scatterchart(
    data = wb_data(wb),
    x = "mpg",
    y = c("disp", "hp")
  )

  wb$
    add_mschart(graph = scatter_plot)$
    add_mschart(dims = "A1", graph = scatter_plot)$
    add_mschart(dims = "F4:L20", graph = scatter_plot)

  expect_true(grepl("absoluteAnchor", wb$drawings))
  expect_true(grepl("oneCellAnchor", wb$drawings))
  expect_true(grepl("twoCellAnchor", wb$drawings))

})

test_that("add_chartsheet works", {

  skip_if_not_installed("mschart")

  require(mschart)

  wb <- wb_workbook()$
    add_worksheet("A & B")$
    add_data(x = mtcars)$
    add_chartsheet(tabColour = "red")

  dat <- wb_data(wb, 1, dims = "A1:E6")

  # call ms_scatterplot
  data_plot <- ms_scatterchart(
    data = dat,
    x = "mpg",
    y = c("disp", "hp"),
    labels = c("disp", "hp")
  )

  wb$add_mschart(graph = data_plot)

  expect_equal(1, nrow(wb$charts))

  expect_true(grepl("A &amp; B", wb$charts$chart))

  expect_true(wb$is_chartsheet[[2]])

  # add new worksheet and replace chart on chartsheet
  wb$add_worksheet()$add_data(x = mtcars)
  dat <- wb_data(wb, dims = "A1:E1;A7:E15")
  data_plot <- ms_scatterchart(
    data = dat,
    x = "mpg",
    y = c("disp", "hp"),
    labels = c("disp", "hp")
  )
  wb$add_mschart(sheet = 2, graph = data_plot)

  expect_equal(2L, nrow(wb$charts))

  exp <- "xdr:absoluteAnchor"
  got <- xml_node_name(wb$drawings, "xdr:wsDr")
  expect_equal(exp, got)

})
